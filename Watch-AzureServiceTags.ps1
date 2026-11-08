<#
.SYNOPSIS
  Monitor Azure Service Tags (Public) and Microsoft 365 Worldwide endpoints.
  - Archives each new JSON (per-source retention enforced)
  - Diffs against the previous version
  - Exports CSV reports ONLY (no Markdown)
  - Optionally emails a summary + CSV attachments (now with HTML body)
  - Optionally posts to Microsoft Teams (MessageCard or Adaptive Card)
  - Azure region filter (default: Global + all U.S. regions via "us*")

.SOURCES
  Azure Service Tags (Public): https://www.microsoft.com/en-us/download/details.aspx?id=56519
  M365 Web Service (Worldwide): https://endpoints.office.com/endpoints/worldwide
  Docs:
    https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide
    https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide

.REQUIREMENTS
  - PowerShell 5.1+ or PowerShell 7+
  - Outbound HTTPS to microsoft.com, endpoints.office.com, and (optional) Teams webhook URL
  - SMTP relay or mailbox credentials if using -SendEmail
#>

[CmdletBinding()]
param(
  [string]$ArchiveRoot = "C:\AzureServiceTags\archive",
  [string]$ReportsRoot = "C:\AzureServiceTags\reports",
  [switch]$VerboseOutput,

  # --- Archive retention ---
  [ValidateScript({ $_ -eq 0 -or $_ -ge 2 })]
  [int]$ArchiveRetention = 5,              # Keep newest N files per archive; 0 disables pruning

  # --- Email options ---
  [switch]$SendEmail,
  [string]$SmtpServer = "mail.com",
  [int]$SmtpPort = 25,
  [string]$MailFrom = "MS-CloudReport@mail.com",
  [string[]]$MailTo = "some-one@mail.com",
  [string[]]$MailCc,
  [string]$SubjectPrefix = "[MS Cloud Endpoints]",
  [string]$SmtpUser,
  [securestring]$SmtpPassword,
  [switch]$UseSsl = $false,
  [switch]$AttachAllReports,
  [switch]$SendOnNoChange,

  # --- Optional source toggles ---
  [bool]$MonitorAzureServiceTags = $true,
  [bool]$MonitorM365Worldwide = $true,

  # --- Azure region filter (supports wildcards; case-insensitive) ---
  [string[]]$AzureRegions = @(
    "Global",   # include no-region services as Global
    "*us",      # ends with 'us' (eastus, centralus, westcentralus, ...)
    "*us2",     # ends with 'us2' (eastus2, westus2, ...)
    "*us3",     # ends with 'us3' (eastus3, westus3, southeastus3, ...)
    "*useuap",  # ends with 'useuap' (centraluseuap, eastus2euap, ...)
    "usstage*"  # starts with 'usstage' (usstagec, usstagee)
  ),
  
  # --- TEAMS: Webhook options ---
  [switch]$PostToTeams,                      # Enable Teams notification - Just set to $true to turn on - off by default
  [ValidateSet('MessageCard','AdaptiveCard')]
  [string]$TeamsPayload = 'MessageCard',     # MessageCard (Incoming Webhook) or AdaptiveCard (Workflows webhook)
  [string]$TeamsWebhookUrl,                  # Webhook URL
  [string]$TeamsTitlePrefix = "[Cloud Endpoints]",
  [int]$TeamsPreviewLimit = 10               # How many items to preview in card lists
)

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# Default to NO TLS on port 25 unless explicitly overridden
if ($SmtpPort -eq 25 -and -not $PSBoundParameters.ContainsKey('UseSsl')) { $UseSsl = $false }

# Ensure TLS 1.2 for WinPS 5.1 Invoke-WebRequest/Invoke-RestMethod
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Ensure base directories exist
New-Item -ItemType Directory -Force -Path $ArchiveRoot, $ReportsRoot | Out-Null

# ---------------- Azure Service Tags helpers ----------------

function Get-LatestJsonUrl {
    # The confirmation page exposes the weekly JSON link
    $confirmUrl = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=56519"
    if ($VerboseOutput) { Write-Host "Fetching confirmation page: $confirmUrl" -ForegroundColor Cyan }
    $resp = Invoke-WebRequest -Uri $confirmUrl -UseBasicParsing

    $jsonHref = ($resp.Links | Where-Object { $_.href -match '\.json$' } |
                 Select-Object -ExpandProperty href -First 1)

    if (-not $jsonHref) {
        if ($resp.Content -match 'href="(?<u>https[^"]+?ServiceTags_Public_\d{8}\.json)"') {
            $jsonHref = $Matches.u
        }
    }
    if (-not $jsonHref) { throw "Could not find Azure Service Tags JSON URL." }
    if ($VerboseOutput) { Write-Host "Azure JSON URL: $jsonHref" -ForegroundColor Green }
    return $jsonHref
}

function Save-LatestJson {
    $url = Get-LatestJsonUrl
    $fileName = [IO.Path]::GetFileName($url)  # e.g., ServiceTags_Public_YYYYMMDD.json
    $dest = Join-Path $ArchiveRoot $fileName

    if (-not (Test-Path $dest)) {
        Write-Host "Downloading Azure Service Tags: $url" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $url -OutFile $dest
    } elseif ($VerboseOutput) {
        Write-Host "Azure already archived: $dest" -ForegroundColor Yellow
    }
    return $dest
}

function Load-ServiceTags {
    param([Parameter(Mandatory)][string]$Path)

    $json = Get-Content -Raw -Path $Path | ConvertFrom-Json
    $meta = [pscustomobject]@{
        version      = $json.version
        changeNumber = $json.changeNumber
        cloud        = $json.cloud
        file         = (Split-Path $Path -Leaf)
    }

    # Map: "<Id>|<Region>" -> HashSet of prefixes
    $map = @{}
    foreach ($v in $json.values) {
        $id     = $v.id
        $props  = $v.properties
        $region = if ($props.region) { $props.region } else { "Global" }
        $key = "$id|$region"

        $prefixes = @()
        if ($props.addressPrefixes)     { $prefixes += $props.addressPrefixes }
        if ($props.addressPrefixesIpv6) { $prefixes += $props.addressPrefixesIpv6 }

        $set = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($p in $prefixes) { [void]$set.Add([string]$p) }
        $map[$key] = $set
    }

    [pscustomobject]@{ Meta = $meta; Map = $map }
}

function Compare-Maps {
    param(
      [Parameter(Mandatory)]$old,
      [Parameter(Mandatory)]$new
    )

    $oldKeys = [System.Collections.Generic.HashSet[string]]::new()
    $newKeys = [System.Collections.Generic.HashSet[string]]::new()
    $old.Map.Keys | ForEach-Object { [void]$oldKeys.Add($_) }
    $new.Map.Keys | ForEach-Object { [void]$newKeys.Add($_) }

    $addedKeys   = $newKeys.Where({ -not $oldKeys.Contains($_) })
    $removedKeys = $oldKeys.Where({ -not $newKeys.Contains($_) })

    $prefixAdds = @()
    $prefixRems = @()

    $common = $newKeys.Where({ $oldKeys.Contains($_) })
    foreach ($k in $common) {
        $oldSet = $old.Map[$k]
        $newSet = $new.Map[$k]

        foreach ($p in $newSet) {
            if (-not $oldSet.Contains($p)) {
                $tag,$region = $k.Split('|',2)
                $prefixAdds += [pscustomobject]@{ ServiceTag=$tag; Region=$region; Prefix=$p }
            }
        }
        foreach ($p in $oldSet) {
            if (-not $newSet.Contains($p)) {
                $tag,$region = $k.Split('|',2)
                $prefixRems += [pscustomobject]@{ ServiceTag=$tag; Region=$region; Prefix=$p }
            }
        }
    }

    $tagAdds = $addedKeys | ForEach-Object {
        $tag,$region = $_.Split('|',2)
        [pscustomobject]@{ ServiceTag=$tag; Region=$region; Change="TagAdded" }
    }
    $tagRems = $removedKeys | ForEach-Object {
        $tag,$region = $_.Split('|',2)
        [pscustomobject]@{ ServiceTag=$tag; Region=$region; Change="TagRemoved" }
    }

    [pscustomobject]@{
        TagAdds     = $tagAdds
        TagRems     = $tagRems
        PrefixAdds  = $prefixAdds
        PrefixRems  = $prefixRems
    }
}

# ---------------- Email + Retention helpers ----------------

function Send-ReportEmail {
    param(
        [Parameter(Mandatory)][string]$SmtpServer,
        [int]$SmtpPort = 587,
        [Parameter(Mandatory)][string]$From,
        [Parameter(Mandatory)][string[]]$To,
        [string[]]$Cc,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyText,
        [switch]$BodyIsHtml,             # NEW: send HTML body when present
        [string[]]$Attachments,
        [string]$SmtpUser,
        [securestring]$SmtpPassword,
        [switch]$UseSsl
    )

    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From = $From
    $To | ForEach-Object { [void]$msg.To.Add($_) }
    if ($Cc) { $Cc | ForEach-Object { [void]$msg.CC.Add($_) } }

    $msg.Subject    = $Subject
    $msg.Body       = $BodyText
    $msg.IsBodyHtml = [bool]$BodyIsHtml

    foreach ($a in $Attachments) {
        if ($a -and (Test-Path $a)) { [void]$msg.Attachments.Add($a) }
    }

    $client = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $client.EnableSsl = [bool]$UseSsl

    if ($SmtpUser) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SmtpPassword)
        try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) } }
        $client.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $plain)
    } else {
        $client.UseDefaultCredentials = $false
    }

    try { $client.Send($msg); Write-Host "Email sent to $($To -join ', ')" -ForegroundColor Green }
    catch { Write-Error "Failed to send email: $($_.Exception.Message)"; throw }
    finally { $msg.Dispose(); $client.Dispose() }
}

function Prune-Archive {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Filter,
        [Parameter(Mandatory)][int]$Keep
    )
    if ($Keep -lt 1) { return }

    $files = Get-ChildItem -Path $Path -Filter $Filter -File |
             Sort-Object Name -Descending

    if ($files.Count -le $Keep) {
        if ($VerboseOutput) { Write-Host "Archive at '$Path' has $($files.Count); retention $Keep. Nothing to prune." -ForegroundColor DarkCyan }
        return
    }

    $toRemove = $files | Select-Object -Skip $Keep
    foreach ($f in $toRemove) {
        try {
            if ($VerboseOutput) { Write-Host "Pruning old file: $($f.FullName)" -ForegroundColor DarkYellow }
            Remove-Item -LiteralPath $f.FullName -Force
        } catch {
            Write-Warning "Failed to remove $($f.FullName): $($_.Exception.Message)"
        }
    }
}

# ---------------- M365 Worldwide helpers ----------------

function Get-M365WorldwideUrl {
    param([switch]$AsCsv)
    # Per docs, a clientrequestid GUID is REQUIRED
    $guid = [guid]::NewGuid().Guid
    $base = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=$guid"
    if ($AsCsv) { return "$base&format=CSV" }
    return $base
}

function Save-M365WorldwideJson {
    param(
        [Parameter(Mandatory)][string]$ArchiveRoot,
        [switch]$VerboseOutput
    )

    $m365Archive = Join-Path $ArchiveRoot "M365"
    New-Item -ItemType Directory -Force -Path $m365Archive | Out-Null

    # Use version to build a stable filename
    $verUrl = "https://endpoints.office.com/version?clientrequestid=$([guid]::NewGuid().Guid)"
    $verObj = Invoke-RestMethod -Uri $verUrl -UseBasicParsing
    $version = $verObj.latest  # e.g., 2025073100

    $dest = Join-Path $m365Archive ("M365_Worldwide_{0}.json" -f $version)
    if (-not (Test-Path $dest)) {
        $url = Get-M365WorldwideUrl
        if ($VerboseOutput) { Write-Host "Downloading M365 Worldwide endpoints: $url" -ForegroundColor Cyan }
        $json = Invoke-RestMethod -Uri $url -UseBasicParsing
        $json | ConvertTo-Json -Depth 8 | Out-File -FilePath $dest -Encoding UTF8
        if ($VerboseOutput) { Write-Host "Archived M365 JSON: $dest" -ForegroundColor Green }
    } elseif ($VerboseOutput) {
        Write-Host "M365 archive already has: $dest" -ForegroundColor Yellow
    }
    return $dest
}

function Load-M365Endpoints {
    param([Parameter(Mandatory)][string]$Path)

    $data = Get-Content -Raw -Path $Path | ConvertFrom-Json

    # Build maps keyed by "ServiceArea|Category"
    $urlMap = @{}  # -> HashSet[string]
    $ipMap  = @{}  # -> HashSet[string]

    foreach ($set in $data) {
        $area = if ($set.serviceArea) { [string]$set.serviceArea } else { "Unknown" }
        $cat  = if ($set.category)    { [string]$set.category    } else { "Default" }
        $key  = "$area|$cat"

        if (-not $urlMap.ContainsKey($key)) { $urlMap[$key] = [System.Collections.Generic.HashSet[string]]::new() }
        if (-not $ipMap.ContainsKey($key))  { $ipMap[$key]  = [System.Collections.Generic.HashSet[string]]::new() }

        if ($set.urls) { foreach ($u in $set.urls) { [void]$urlMap[$key].Add([string]$u) } }
        if ($set.ips)  { foreach ($ip in $set.ips) { [void]$ipMap[$key].Add([string]$ip) } }
    }

    [pscustomobject]@{
        Meta = [pscustomobject]@{ file = (Split-Path $Path -Leaf) }
        Urls = $urlMap
        IPs  = $ipMap
    }
}

function Compare-M365Maps {
    param(
        [Parameter(Mandatory)]$old,
        [Parameter(Mandatory)]$new
    )

    $keysOld = [System.Collections.Generic.HashSet[string]]::new()
    $keysNew = [System.Collections.Generic.HashSet[string]]::new()
    ($old.Urls.Keys + $old.IPs.Keys | Select-Object -Unique) | ForEach-Object { [void]$keysOld.Add($_) }
    ($new.Urls.Keys + $new.IPs.Keys | Select-Object -Unique) | ForEach-Object { [void]$keysNew.Add($_) }

    $allKeys = ($keysOld + $keysNew) | Select-Object -Unique

    $urlAdds = @(); $urlRems = @()
    $ipAdds  = @(); $ipRems  = @()

    foreach ($k in $allKeys) {
        $oU = $old.Urls[$k]; if (-not $oU) { $oU = [System.Collections.Generic.HashSet[string]]::new() }
        $nU = $new.Urls[$k]; if (-not $nU) { $nU = [System.Collections.Generic.HashSet[string]]::new() }
        foreach ($u in $nU) { if (-not $oU.Contains($u)) { $area,$cat=$k.Split('|',2); $urlAdds += [pscustomobject]@{ServiceArea=$area;Category=$cat;Url=$u} } }
        foreach ($u in $oU) { if (-not $nU.Contains($u)) { $area,$cat=$k.Split('|',2); $urlRems += [pscustomobject]@{ServiceArea=$area;Category=$cat;Url=$u} } }

        $oI = $old.IPs[$k]; if (-not $oI) { $oI = [System.Collections.Generic.HashSet[string]]::new() }
        $nI = $new.IPs[$k]; if (-not $nI) { $nI = [System.Collections.Generic.HashSet[string]]::new() }
        foreach ($ip in $nI) { if (-not $oI.Contains($ip)) { $area,$cat=$k.Split('|',2); $ipAdds += [pscustomobject]@{ServiceArea=$area;Category=$cat;IP=$ip} } }
        foreach ($ip in $oI) { if (-not $nI.Contains($ip)) { $area,$cat=$k.Split('|',2); $ipRems += [pscustomobject]@{ServiceArea=$area;Category=$cat;IP=$ip} } }
    }

    [pscustomobject]@{
        UrlAdds = $urlAdds
        UrlRems = $urlRems
        IpAdds  = $ipAdds
        IpRems  = $ipRems
    }
}

function Write-M365Reports {
    param(
        [Parameter(Mandatory)]$Diff,
        [Parameter(Mandatory)][string]$ReportsRoot
    )
    $m365Reports = Join-Path $ReportsRoot "M365"
    New-Item -ItemType Directory -Force -Path $m365Reports | Out-Null

    $urlsAddedCsv   = Join-Path $m365Reports "m365_urls_added.csv"
    $urlsRemovedCsv = Join-Path $m365Reports "m365_urls_removed.csv"
    $ipsAddedCsv    = Join-Path $m365Reports "m365_ips_added.csv"
    $ipsRemovedCsv  = Join-Path $m365Reports "m365_ips_removed.csv"

    $Diff.UrlAdds | Sort-Object ServiceArea,Category,Url | Export-Csv -NoTypeInformation -Path $urlsAddedCsv
    $Diff.UrlRems | Sort-Object ServiceArea,Category,Url | Export-Csv -NoTypeInformation -Path $urlsRemovedCsv
    $Diff.IpAdds  | Sort-Object ServiceArea,Category,IP  | Export-Csv -NoTypeInformation -Path $ipsAddedCsv
    $Diff.IpRems  | Sort-Object ServiceArea,Category,IP  | Export-Csv -NoTypeInformation -Path $ipsRemovedCsv

    [pscustomobject]@{
        UrlsAdded   = $urlsAddedCsv
        UrlsRemoved = $urlsRemovedCsv
        IpsAdded    = $ipsAddedCsv
        IpsRemoved  = $ipsRemovedCsv
    }
}

# ---------------- HTML email builder ----------------

function New-EndpointSummaryHtml {
    param(
        # Azure section (nullable)
        [string]$AzOldFile, [string]$AzNewFile,
        [int]$AzPrefixAdds = 0, [int]$AzPrefixRems = 0,
        [int]$AzTagAdds = 0,    [int]$AzTagRems = 0,
        [object[]]$AzPreviewAdd = @(), [object[]]$AzPreviewRem = @(),

        # M365 section (nullable)
        [string]$M365OldFile, [string]$M365NewFile,
        [int]$M365IpAdds = 0,  [int]$M365IpRems = 0,
        [int]$M365UrlAdds = 0, [int]$M365UrlRems = 0,
        [object[]]$M365PrevUrlAdds = @(), [object[]]$M365PrevUrlRems = @(),
        [object[]]$M365PrevIpAdds  = @(), [object[]]$M365PrevIpRems  = @()
    )
	
    # Cross-platform HTML encoding
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $style = @"
<style type="text/css">
  body { font-family:Segoe UI, Arial, sans-serif; font-size:12.5px; color:#222; }
  h1   { font-size:18px; margin:0 0 8px 0; }
  h2   { font-size:16px; margin:18px 0 6px 0; }
  .muted { color:#666; }
  table { border-collapse:collapse; margin:6px 0 12px 0; }
  th, td { border:1px solid #ddd; padding:6px 8px; }
  th { background:#f3f3f3; text-align:left; }
  ul { margin:6px 0 12px 18px; }
  .kpi { font-weight:600; }
  .add { color:#1a7f37; }     /* green */
  .rem { color:#b00020; }     /* red   */
  .sec { border-left:4px solid #ddd; padding-left:10px; }
  .mono { font-family:Consolas, 'Courier New', monospace; }
</style>
"@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8' />$style</head><body>")
    [void]$sb.AppendLine("<h1>Cloud Endpoints Diff Summary</h1>")
    [void]$sb.AppendLine("<p class='muted'>This message summarizes weekly Azure Service Tags and monthly Microsoft 365 Worldwide endpoint changes.</p>")

    if ($AzNewFile) {
        [void]$sb.AppendLine("<div class='sec'>")
        [void]$sb.AppendLine("<h2>Azure Service Tags (Public)</h2>")
        if ($AzOldFile) { [void]$sb.AppendLine("<div><span class='mono'>Old:</span> $(& $enc $AzOldFile)</div>") }
        [void]$sb.AppendLine("<div><span class='mono'>New:</span> $(& $enc $AzNewFile)</div>")
        [void]$sb.AppendLine("<table><thead><tr><th>Metric</th><th>Count</th></tr></thead><tbody>")
        [void]$sb.AppendLine("<tr><td>Prefix additions</td><td class='kpi add'>$AzPrefixAdds</td></tr>")
        [void]$sb.AppendLine("<tr><td>Prefix removals</td><td class='kpi rem'>$AzPrefixRems</td></tr>")
        [void]$sb.AppendLine("<tr><td>Tag additions</td><td class='kpi add'>$AzTagAdds</td></tr>")
        [void]$sb.AppendLine("<tr><td>Tag removals</td><td class='kpi rem'>$AzTagRems</td></tr>")
        [void]$sb.AppendLine("</tbody></table>")

        if ($AzPreviewAdd -and $AzPreviewAdd.Count -gt 0) {
            [void]$sb.AppendLine("<div><strong>Top 10 new prefixes</strong><ul>")
            foreach ($i in $AzPreviewAdd) {
                $li = ("{0} — {1}: <span class='mono'>{2}</span>" -f (& $enc $i.ServiceTag), (& $enc $i.Region), (& $enc $i.Prefix))
                [void]$sb.AppendLine("<li>$li</li>")
            }
            [void]$sb.AppendLine("</ul></div>")
        }
        if ($AzPreviewRem -and $AzPreviewRem.Count -gt 0) {
            [void]$sb.AppendLine("<div><strong>Top 10 removed prefixes</strong><ul>")
            foreach ($i in $AzPreviewRem) {
                $li = ("{0} — {1}: <span class='mono'>{2}</span>" -f (& $enc $i.ServiceTag), (& $enc $i.Region), (& $enc $i.Prefix))
                [void]$sb.AppendLine("<li>$li</li>")
            }
            [void]$sb.AppendLine("</ul></div>")
        }
        [void]$sb.AppendLine("</div>")
    }

    if ($M365NewFile) {
        [void]$sb.AppendLine("<div class='sec'>")
        [void]$sb.AppendLine("<h2>Microsoft 365 (Worldwide)</h2>")
        if ($M365OldFile) { [void]$sb.AppendLine("<div><span class='mono'>Old:</span> $(& $enc $M365OldFile)</div>") }
        [void]$sb.AppendLine("<div><span class='mono'>New:</span> $(& $enc $M365NewFile)</div>")
        [void]$sb.AppendLine("<table><thead><tr><th>Metric</th><th>Count</th></tr></thead><tbody>")
        [void]$sb.AppendLine("<tr><td>IP additions</td><td class='kpi add'>$M365IpAdds</td></tr>")
        [void]$sb.AppendLine("<tr><td>IP removals</td><td class='kpi rem'>$M365IpRems</td></tr>")
        [void]$sb.AppendLine("<tr><td>URL additions</td><td class='kpi add'>$M365UrlAdds</td></tr>")
        [void]$sb.AppendLine("<tr><td>URL removals</td><td class='kpi rem'>$M365UrlRems</td></tr>")
        [void]$sb.AppendLine("</tbody></table>")

        if ($M365PrevUrlAdds -and $M365PrevUrlAdds.Count -gt 0) {
            [void]$sb.AppendLine("<div><strong>Top 10 new URLs</strong><ul>")
            foreach ($u in $M365PrevUrlAdds) {
                $li = ("{0} [{1}]: <span class='mono'>{2}</span>" -f (& $enc $u.ServiceArea), (& $enc $u.Category), (& $enc $u.Url))
                [void]$sb.AppendLine("<li>$li</li>")
            }
            [void]$sb.AppendLine("</ul></div>")
        }
        if ($M365PrevUrlRems -and $M365PrevUrlRems.Count -gt 0) {
            [void]$sb.AppendLine("<div><strong>Top 10 removed URLs</strong><ul>")
            foreach ($u in $M365PrevUrlRems) {
                $li = ("{0} [{1}]: <span class='mono'>{2}</span>" -f (& $enc $u.ServiceArea), (& $enc $u.Category), (& $enc $u.Url))
                [void]$sb.AppendLine("<li>$li</li>")
            }
            [void]$sb.AppendLine("</ul></div>")
        }
        if ($M365PrevIpAdds -and $M365PrevIpAdds.Count -gt 0) {
            [void]$sb.AppendLine("<div><strong>Top 10 new IPs</strong><ul>")
            foreach ($ip in $M365PrevIpAdds) {
                $li = ("{0} [{1}]: <span class='mono'>{2}</span>" -f (& $enc $ip.ServiceArea), (& $enc $ip.Category), (& $enc $ip.IP))
                [void]$sb.AppendLine("<li>$li</li>")
            }
            [void]$sb.AppendLine("</ul></div>")
        }
        if ($M365PrevIpRems -and $M365PrevIpRems.Count -gt 0) {
            [void]$sb.AppendLine("<div><strong>Top 10 removed IPs</strong><ul>")
            foreach ($ip in $M365PrevIpRems) {
                $li = ("{0} [{1}]: <span class='mono'>{2}</span>" -f (& $enc $ip.ServiceArea), (& $enc $ip.Category), (& $enc $ip.IP))
                [void]$sb.AppendLine("<li>$li</li>")
            }
            [void]$sb.AppendLine("</ul></div>")
        }
        [void]$sb.AppendLine("</div>")
    }

    [void]$sb.AppendLine("<p class='muted'>Attachments: CSV files with full details.</p>")
    [void]$sb.AppendLine("</body></html>")
    return $sb.ToString()
}

# ---------------- TEAMS: builders and sender (optional) ----------------

function New-TeamsMessageCard {
    param(
        [string]$Title,
        [string]$Summary,
        [hashtable[]]$Facts,
        [string[]]$Bullets = @(),
        [string]$ThemeColor = "0078D7"
    )
    $text = if ($Bullets -and $Bullets.Count -gt 0) { "- " + ($Bullets -join "<br>- ") } else { "" }
    return @{
        "@type"    = "MessageCard"
        "@context" = "https://schema.org/extensions"
        "themeColor" = $ThemeColor
        "summary"  = $Summary
        "title"    = $Title
        "sections" = @(@{ "facts" = $Facts; "text" = $text })
    }
}

function New-TeamsAdaptiveCard {
    param(
        [string]$Title,
        [hashtable[]]$KeyValues,   # @{ label=""; value="" }
        [string[]]$Bullets = @()
    )
    $body = @(@{ "type"="TextBlock"; "size"="Medium"; "weight"="Bolder"; "text"=$Title; "wrap"=$true })
    if ($KeyValues -and $KeyValues.Count -gt 0) {
        $facts = @(); foreach ($kv in $KeyValues) { $facts += @{ "title"=$kv.label; "value"=$kv.value } }
        $body += @{ "type"="FactSet"; "facts"=$facts }
    }
    if ($Bullets -and $Bullets.Count -gt 0) {
        $items = @(); foreach ($b in $Bullets) { $items += @{ "type"="TextBlock"; "text"="• $b"; "wrap"=$true } }
        $body += @{ "type"="Container"; "items"=$items }
    }
    $adaptive = @{ "$schema"="http://adaptivecards.io/schemas/adaptive-card.json"; "type"="AdaptiveCard"; "version"="1.4"; "body"=$body }
    return @{ "type"="message"; "attachments"=@(@{ "contentType"="application/vnd.microsoft.card.adaptive"; "content"=$adaptive }) }
}

function Send-TeamsWebhook {
    param(
        [Parameter(Mandatory)][string]$WebhookUrl,
        [Parameter(Mandatory)][object]$Payload
    )
    $json = $Payload | ConvertTo-Json -Depth 10
    if (($json.Length) -gt 28000) { throw "Teams payload exceeds ~28 KB limit. Reduce preview items or content." }
    $attempt = 0; $maxAttempts = 3; $delay = 1
    while ($true) {
        try {
            Invoke-RestMethod -Method Post -Uri $WebhookUrl -ContentType 'application/json' -Body $json -UseBasicParsing | Out-Null
            Write-Host "Posted Teams webhook" -ForegroundColor Green
            break
        } catch {
            $attempt++
            $status = $_.Exception.Response.StatusCode.Value__ 2>$null
            if ($attempt -lt $maxAttempts -and ($status -eq 429 -or $status -ge 500)) {
                Start-Sleep -Seconds $delay
                $delay = [Math]::Min(30, $delay * 2)
                continue
            }
            throw
        }
    }
}

# ---------------- Utilities ----------------

function Test-RegionAllowed {
    param([string]$Region, [string[]]$AllowedPatterns)
    if (-not $Region) { $Region = 'Global' }   # treat missing as Global
    foreach ($pat in $AllowedPatterns) {
        if ($Region -like $pat) { return $true } # wildcard match, case-insensitive by default
    }
    return $false
}

# ------------------------------ MAIN ------------------------------

try {
    # Track outputs to email/teams
    $attachments = @()
    $combinedHasChanges = $false

    # ------------- Azure Service Tags -------------
    $azureAdds = 0; $azureRems = 0; $azureTagAdds = 0; $azureTagRems = 0
    $addCsv = $null; $remCsv = $null; $tagCsv = $null
    $azureNewFile = $null; $azureOldFile = $null

    # Filter helper for Azure rows
    $regionFilterScriptBlock = {
        param($row)
        Test-RegionAllowed -Region $row.Region -AllowedPatterns $AzureRegions
    }

    $azurePreviewAdds = @(); $azurePreviewRems = @()

    if ($MonitorAzureServiceTags) {
        $latestPath = Save-LatestJson
        # Prune only Azure files in root archive
        Prune-Archive -Path $ArchiveRoot -Filter "ServiceTags_Public_*.json" -Keep $ArchiveRetention

        $files = Get-ChildItem -Path $ArchiveRoot -Filter "ServiceTags_Public_*.json" -File | Sort-Object Name
        if ($files.Count -lt 2) {
            Write-Host "Azure: only one JSON in archive. Nothing to compare yet." -ForegroundColor Yellow
        } else {
            $oldPath = $files[$files.Count - 2].FullName
            $newPath = $files[$files.Count - 1].FullName
            $azureOldFile = (Split-Path $oldPath -Leaf)
            $azureNewFile = (Split-Path $newPath -Leaf)

            $old = Load-ServiceTags -Path $oldPath
            $new = Load-ServiceTags -Path $newPath

            $diff = Compare-Maps -old $old -new $new

            # Apply the region filter before exporting and counting.
            $azAddsFiltered    = $diff.PrefixAdds | Where-Object $regionFilterScriptBlock
            $azRemsFiltered    = $diff.PrefixRems | Where-Object $regionFilterScriptBlock
            $azTagAddsFiltered = $diff.TagAdds    | Where-Object $regionFilterScriptBlock
            $azTagRemsFiltered = $diff.TagRems    | Where-Object $regionFilterScriptBlock

            # Export CSVs (filtered)
            $addCsv = Join-Path $ReportsRoot "diff_added_prefixes.csv"
            $remCsv = Join-Path $ReportsRoot "diff_removed_prefixes.csv"
            $tagCsv = Join-Path $ReportsRoot "diff_tag_changes.csv"

            $azAddsFiltered | Sort-Object ServiceTag,Region,Prefix | Export-Csv -NoTypeInformation -Path $addCsv
            $azRemsFiltered | Sort-Object ServiceTag,Region,Prefix | Export-Csv -NoTypeInformation -Path $remCsv
            @($azTagAddsFiltered + $azTagRemsFiltered) | Sort-Object Change,ServiceTag,Region | Export-Csv -NoTypeInformation -Path $tagCsv

            Write-Host "Azure CSV reports written (filtered by regions: $($AzureRegions -join ', ')):" -ForegroundColor Green
            Write-Host " - $addCsv"
            Write-Host " - $remCsv"
            Write-Host " - $tagCsv"

            # Counts for email/Teams use filtered numbers
            $azureAdds    = $azAddsFiltered.Count
            $azureRems    = $azRemsFiltered.Count
            $azureTagAdds = $azTagAddsFiltered.Count
            $azureTagRems = $azTagRemsFiltered.Count

            # Previews (filtered)
            $azurePreviewAdds = $azAddsFiltered | Select-Object -First 10
            $azurePreviewRems = $azRemsFiltered | Select-Object -First 10

            if ($AttachAllReports) {
                $attachments += @($addCsv, $remCsv, $tagCsv)
            } else {
                if ($azureAdds -gt 0) { $attachments += $addCsv }
                if ($azureRems -gt 0) { $attachments += $remCsv }
                if ($azureTagAdds -gt 0 -or $azureTagRems -gt 0) { $attachments += $tagCsv }
            }

            $combinedHasChanges = $combinedHasChanges -or (($azureAdds + $azureRems + $azureTagAdds + $azureTagRems) -gt 0)
        }
    }

    # ------------- M365 Worldwide -------------
    $m365UrlsAdd = 0; $m365UrlsRem = 0; $m365IpsAdd = 0; $m365IpsRem = 0
    $m365NewFile = $null; $m365OldFile = $null
    $m365PreviewUrlAdds = @(); $m365PreviewUrlRems = @()
    $m365PreviewIpAdds = @();  $m365PreviewIpRems  = @()

    if ($MonitorM365Worldwide) {
        $m365Latest = Save-M365WorldwideJson -ArchiveRoot $ArchiveRoot -VerboseOutput:$VerboseOutput
        Prune-Archive -Path (Join-Path $ArchiveRoot "M365") -Filter "M365_Worldwide_*.json" -Keep $ArchiveRetention

        $m365Files = Get-ChildItem -Path (Join-Path $ArchiveRoot "M365") -Filter "M365_Worldwide_*.json" -File | Sort-Object Name
        if ($m365Files.Count -lt 2) {
            Write-Host "M365: only one JSON in archive. Nothing to compare yet." -ForegroundColor Yellow
        } else {
            $m365Old = $m365Files[$m365Files.Count - 2].FullName
            $m365New = $m365Files[$m365Files.Count - 1].FullName
            $m365OldFile = (Split-Path $m365Old -Leaf)
            $m365NewFile = (Split-Path $m365New -Leaf)

            $mOld = Load-M365Endpoints -Path $m365Old
            $mNew = Load-M365Endpoints -Path $m365New
            $mDiff = Compare-M365Maps -old $mOld -new $mNew

            $m365Reports = Write-M365Reports -Diff $mDiff -ReportsRoot $ReportsRoot

            Write-Host "M365 CSV reports written:" -ForegroundColor Green
            $m365Reports.PSObject.Properties.Value | ForEach-Object { Write-Host " - $_" }

            $m365UrlsAdd = $mDiff.UrlAdds.Count
            $m365UrlsRem = $mDiff.UrlRems.Count
            $m365IpsAdd  = $mDiff.IpAdds.Count
            $m365IpsRem  = $mDiff.IpRems.Count

            $m365PreviewUrlAdds = $mDiff.UrlAdds | Select-Object -First 10
            $m365PreviewUrlRems = $mDiff.UrlRems | Select-Object -First 10
            $m365PreviewIpAdds  = $mDiff.IpAdds  | Select-Object -First 10
            $m365PreviewIpRems  = $mDiff.IpRems  | Select-Object -First 10

            if ($AttachAllReports) {
                $attachments += @($m365Reports.UrlsAdded, $m365Reports.UrlsRemoved, $m365Reports.IpsAdded, $m365Reports.IpsRemoved)
            } else {
                if ($m365UrlsAdd -gt 0) { $attachments += $m365Reports.UrlsAdded }
                if ($m365UrlsRem -gt 0) { $attachments += $m365Reports.UrlsRemoved }
                if ($m365IpsAdd  -gt 0) { $attachments += $m365Reports.IpsAdded }
                if ($m365IpsRem  -gt 0) { $attachments += $m365Reports.IpsRemoved }
            }

            $combinedHasChanges = $combinedHasChanges -or (($m365UrlsAdd + $m365UrlsRem + $m365IpsAdd + $m365IpsRem) -gt 0)
        }
    }

    # ------------- Email -------------
    if ($SendEmail) {
        # Construct subject
        $subjectParts = @()
        if ($MonitorAzureServiceTags -and $azureNewFile) {
            $subjectParts += ("AZURE: +{0}/-{1} prefixes; +{2}/-{3} tags " -f $azureAdds,$azureRems,$azureTagAdds,$azureTagRems)
        }
        if ($MonitorM365Worldwide -and $m365NewFile) {
            $subjectParts += ("M365: +{0}/-{1} IPs; +{2}/-{3} URLs " -f $m365IpsAdd,$m365IpsRem,$m365UrlsAdd,$m365UrlsRem)
        }
        if (-not $subjectParts) { $subjectParts = @("No changes") }
        $subject = "$SubjectPrefix " + ($subjectParts -join " | ")

        # Build HTML body
        $bodyHtml = New-EndpointSummaryHtml `
            -AzOldFile $azureOldFile -AzNewFile $azureNewFile `
            -AzPrefixAdds $azureAdds -AzPrefixRems $azureRems `
            -AzTagAdds $azureTagAdds -AzTagRems $azureTagRems `
            -AzPreviewAdd $azurePreviewAdds -AzPreviewRem $azurePreviewRems `
            -M365OldFile $m365OldFile -M365NewFile $m365NewFile `
            -M365IpAdds $m365IpsAdd -M365IpRems $m365IpsRem `
            -M365UrlAdds $m365UrlsAdd -M365UrlRems $m365UrlsRem `
            -M365PrevUrlAdds $m365PreviewUrlAdds -M365PrevUrlRems $m365PreviewUrlRems `
            -M365PrevIpAdds $m365PreviewIpAdds -M365PrevIpRems $m365PreviewIpRems

        if ($combinedHasChanges -or $SendOnNoChange) {
            Send-ReportEmail `
                -SmtpServer $SmtpServer `
                -SmtpPort $SmtpPort `
                -From $MailFrom `
                -To $MailTo `
                -Cc $MailCc `
                -Subject $subject `
                -BodyText $bodyHtml `
                -BodyIsHtml `
                -Attachments $attachments `
                -SmtpUser $SmtpUser `
                -SmtpPassword $SmtpPassword `
                -UseSsl:$UseSsl
        } elseif ($VerboseOutput) {
            Write-Host "No changes detected across feeds; email suppressed (use -SendOnNoChange to force)." -ForegroundColor Yellow
        }
    }

    # ------------- TEAMS (optional) -------------
    if ($PostToTeams -and $combinedHasChanges) {
        if (-not $TeamsWebhookUrl) { throw "Teams posting enabled but -TeamsWebhookUrl not provided." }

        $bul = @()
        if ($MonitorAzureServiceTags -and $azureNewFile) {
            $bul += "Azure (filtered): +$azureAdds/ -$azureRems prefixes; +$azureTagAdds/ -$azureTagRems tags ($azureNewFile)"
        }
        if ($MonitorM365Worldwide -and $m365NewFile) {
            $bul += "M365: +$m365IpsAdd/ -$m365IpsRem IPs; +$m365UrlsAdd/ -$m365UrlsRem URLs ($m365NewFile)"
        }

        $N = [Math]::Max(0, [Math]::Min($TeamsPreviewLimit, 10))
        if ($N -gt 0) {
            if ($azurePreviewAdds) { $bul += "Top new Azure prefixes: " + (($azurePreviewAdds | Select-Object -First $N | ForEach-Object { "$($_.ServiceTag)/$($_.Region): $($_.Prefix)" }) -join ", ") }
            if ($azurePreviewRems) { $bul += "Top removed Azure prefixes: " + (($azurePreviewRems | Select-Object -First $N | ForEach-Object { "$($_.ServiceTag)/$($_.Region): $($_.Prefix)" }) -join ", ") }
            if ($m365PreviewIpAdds)  { $bul += "Top new M365 IPs: " + (($m365PreviewIpAdds  | Select-Object -First $N | ForEach-Object { "$($_.ServiceArea)[$($_.Category)]: $($_.IP)" }) -join ", ") }
            if ($m365PreviewIpRems)  { $bul += "Top removed M365 IPs: " + (($m365PreviewIpRems| Select-Object -First $N | ForEach-Object { "$($_.ServiceArea)[$($_.Category)]: $($_.IP)" }) -join ", ") }
            if ($m365PreviewUrlAdds) { $bul += "Top new M365 URLs: " + (($m365PreviewUrlAdds | Select-Object -First $N | ForEach-Object { "$($_.ServiceArea)[$($_.Category)]: $($_.Url)" }) -join ", ") }
            if ($m365PreviewUrlRems) { $bul += "Top removed M365 URLs: " + (($m365PreviewUrlRems| Select-Object -First $N | ForEach-Object { "$($_.ServiceArea)[$($_.Category)]: $($_.Url)" }) -join ", ") }
        }

        $title = "$TeamsTitlePrefix Changes"
        $summary = "Cloud endpoint changes detected"

        switch ($TeamsPayload) {
            'MessageCard' {
                $facts = @()
                if ($MonitorAzureServiceTags -and $azureNewFile) {
                    $facts += @{ name="Azure new"; value=$azureNewFile }
                    if ($azureOldFile) { $facts += @{ name="Azure old"; value=$azureOldFile } }
                    $facts += @{ name="Azure prefixes + / -"; value="$azureAdds / $azureRems" }
                    $facts += @{ name="Azure tags + / -";     value="$azureTagAdds / $azureTagRems" }
                }
                if ($MonitorM365Worldwide -and $m365NewFile) {
                    $facts += @{ name="M365 new"; value=$m365NewFile }
                    if ($m365OldFile) { $facts += @{ name="M365 old"; value=$m365OldFile } }
                    $facts += @{ name="M365 IPs + / -";  value="$m365IpsAdd / $m365IpsRem" }
                    $facts += @{ name="M365 URLs + / -"; value="$m365UrlsAdd / $m365UrlsRem" }
                }
                $payload = New-TeamsMessageCard -Title $title -Summary $summary -Facts $facts -Bullets $bul
            }
            'AdaptiveCard' {
                $kv = @()
                if ($MonitorAzureServiceTags -and $azureNewFile) {
                    $kv += @{ label="Azure new"; value=$azureNewFile }
                    if ($azureOldFile) { $kv += @{ label="Azure old"; value=$azureOldFile } }
                    $kv += @{ label="Azure prefixes + / -"; value="$azureAdds / $azureRems" }
                    $kv += @{ label="Azure tags + / -";     value="$azureTagAdds / $azureTagRems" }
                }
                if ($MonitorM365Worldwide -and $m365NewFile) {
                    $kv += @{ label="M365 new"; value=$m365NewFile }
                    if ($m365OldFile) { $kv += @{ label="M365 old"; value=$m365OldFile } }
                    $kv += @{ label="M365 IPs + / -";  value="$m365IpsAdd / $m365IpsRem" }
                    $kv += @{ label="M365 URLs + / -"; value="$m365UrlsAdd / $m365UrlsRem" }
                }
                $payload = New-TeamsAdaptiveCard -Title $title -KeyValues $kv -Bullets $bul
            }
        }

        Send-TeamsWebhook -WebhookUrl $TeamsWebhookUrl -Payload $payload
    } elseif ($PostToTeams -and -not $combinedHasChanges -and $VerboseOutput) {
        Write-Host "No changes detected; Teams notification suppressed." -ForegroundColor Yellow
    }

} catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}