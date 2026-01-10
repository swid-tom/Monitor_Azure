# MonitorAzure_IP - Watch-AzureServiceTags

<img src="/Images/item-1.png" alt="Image Description" width="300" height="300">

---------------------------------------------------------

## Info

A PowerShell utility that monitors two official Microsoft sources:

1. Azure IP Ranges & Service Tags (Public Cloud) — weekly JSON from Microsoft Download Center.
2. Microsoft 365 (Worldwide) URLs & IP ranges — REST web service at endpoints.office.com (same data as the “URLs and IP address ranges” page).

It archives each release (with retention), diffs vs. the previous version, writes CSV reports only, and can email a plain‑text summary with attachments.

### References (official docs)

- Azure Service Tags (Public): <https://www.microsoft.com/en-us/download/details.aspx?id=56519>
- Microsoft 365 IP/URL Web Service: <https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide>
- Microsoft 365 URLs & IP address ranges (firewalls/proxies): <https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide>

---------------------------------------------------------

## Contents

- Features
- Prerequisites
- Install
- Configuration (parameters)
- Usage examples
- Filtering Regions
- Output files
- Email configuration notes
- MS Teams configuration notes
- Archive retention
- Scheduling (Task Scheduler)
- Troubleshooting
- Security considerations
- Customization ideas

---------------------------------------------------------

## Features

- Azure Service Tags (Public)
    - Discovers the current JSON link from the confirmation page, downloads & archives it.
    - Diffs service tag and IP prefix changes (IPv4/IPv6) by region.
    - Exports CSVs:
        - azure_added_prefixes.csv
        - azure_removed_prefixes.csv
        - azure_tag_changes.csv
    - Region filtered
```
# --- Azure region filter (supports wildcards; case-insensitive) ---
  [string[]]$AzureRegions = @("Global","us*"),
```

- Microsoft 365 Worldwide endpoints
    - Calls https://endpoints.office.com/version to get the latest version, then fetches https://endpoints.office.com/endpoints/worldwide?clientrequestid=<GUID>.
    - M365 block deliberately does not create a new JSON unless Microsoft publishes a new version. We call https://endpoints.office.com/version to get the current version string; if that version already exists in your archive\M365\ folder, the script skips downloading again. That’s by design and matches Microsoft’s guidance to key off the service version.
    - Diffs URLs and IP prefixes by ServiceArea and Category (Optimize/Allow/Default).
    - Exports CSVs:
        - M365\m365_urls_added.csv
        - M365\m365_urls_removed.csv
        - M365\m365_ips_added.csv
        - M365\m365_ips_removed.csv

- Retention
    - Keeps the newest N JSON files (default 5) per archive path and prunes the rest.
- Email (optional)
    - Sends a plain-text summary and attaches CSVs (either only changed files or all).

## Details

- Azure
  The script scrapes the confirmation page (/download/confirmation.aspx?id=56519) to find the weekly JSON, then archives it. This approach follows Microsoft’s delivery model for the Azure Service Tags JSON (weekly updates).
  Ref: <https://www.microsoft.com/en-us/download/details.aspx?id=56519>

- Microsoft 365 Worldwide
  The script queries /version to get a stable version stamp and names the file as M365_Worldwide_<version>.json. It then downloads /endpoints/worldwide with a fresh clientrequestid GUID, as required by the web service.
  Refs:
    - Web service: <https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide>
    - Firewall/proxy data source: <https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide>

- Diffing
    - Azure: compares two newest JSONs by key <Id>|<Region> and lists added/removed prefixes and tags.
    - M365: compares two newest JSONs by key <ServiceArea>|<Category>, listing added/removed URLs and IP prefixes.

---------------------------------------------------------

## Prerequisites

- PowerShell 5.1 or 7+
- Network access to microsoft.com and endpoints.office.com to fetch the file
- SMTP server (if using -SendEmail)
    - Relay on port 25 without TLS, or
    - Auth on port 587 with TLS (e.g., Microsoft 365)

---------------------------------------------------------

## Install

1. Save the script locally, e.g.:
```
C:\Scripts\Watch-AzureServiceTags.ps1
```

2. Ensure the default directories exist or let the script create them:
    - C:\AzureServiceTags\archive
    - C:\AzureServiceTags\reports

If your execution policy blocks scripts, you can run with:
```
powershell.exe -ExecutionPolicy Bypass -File C:\Scripts\Watch-AzureServiceTags.ps1 ...
```

---------------------------------------------------------

## Configuration (parameters)

```
.\Watch-AzureServiceTags.ps1 `
  [-ArchiveRoot "C:\AzureServiceTags\archive"] `
  [-ReportsRoot "C:\AzureServiceTags\reports"] `
  [-VerboseOutput] `
  [-ArchiveRetention 5] `
  [-SendEmail] `
  [-SmtpServer <string>] [-SmtpPort <int>] [-UseSsl] `
  [-MailFrom <string>] [-MailTo <string[]>] [-MailCc <string[]>] `
  [-SubjectPrefix "[Azure Service Tags]"] `
  [-SmtpUser <string>] [-SmtpPassword <securestring>] `
  [-AttachAllReports] `
  [-SendOnNoChange]
```
Can adjust settings without switches here within the script:
```
param(
  [string]$ArchiveRoot = "C:\Tasks\IS\AzureServiceTags\archive",     <---Change this to your desired archive path--->
  [string]$ReportsRoot = "C:\Tasks\IS\AzureServiceTags\reports",     <---Change this to your desired reports path--->
  [switch]$VerboseOutput,

  # --- Archive retention ---
  [int]$ArchiveRetention = 5,              # Keep newest N files per archive

  # --- Email options ---
  [switch]$SendEmail = $true,
  [string]$SmtpServer = "mail.com",
  [int]$SmtpPort = 25,
  [string]$MailFrom = "MS-CloudReport@mail.com",
  [string[]]$MailTo = "some-one@yourdomain.com",
  [string[]]$MailCc,
  [string]$SubjectPrefix = "[MS Cloud Endpoints]",
  [string]$SmtpUser,
  [securestring]$SmtpPassword,
  [switch]$UseSsl = $false,
  [switch]$AttachAllReports,
  [switch]$SendOnNoChange,

  # --- Optional source toggles ---
  [bool]$MonitorAzureServiceTags = $true
  [bool]$MonitorM365Worldwide = $true,

  # --- Azure region filter (supports wildcards; case-insensitive) ---
  [string[]]$AzureRegions = @("Global","us*"),

  # --- TEAMS: Webhook options ---
  [switch]$PostToTeams,                      # Enable Teams notification
  [ValidateSet('MessageCard','AdaptiveCard')]
  [string]$TeamsPayload = 'MessageCard',     # MessageCard (Incoming Webhook) or AdaptiveCard (Workflows webhook)
  [string]$TeamsWebhookUrl,                  # Webhook URL
  [string]$TeamsTitlePrefix = "[Cloud Endpoints]",
  [int]$TeamsPreviewLimit = 10               # How many items to preview in card lists
)


```

### Key behaviors
- Port 25 default: If you choose -SmtpPort 25 and you don’t pass -UseSsl, the script defaults to no TLS.
- Retention: -ArchiveRetention controls how many JSON files are kept (default 5).
- Email send threshold: By default, emails send only when changes are detected. Use -SendOnNoChange to always send.

---------------------------------------------------------

## Usage examples

1. Basic run (no email)
```
.\Watch-AzureServiceTags.ps1 `
  -ArchiveRoot "C:\AzureServiceTags\archive" `
  -ReportsRoot "C:\AzureServiceTags\reports" `
  -VerboseOutput
```

2. Send via internal SMTP relay (no TLS, port 25)
```
.\Watch-AzureServiceTags.ps1 `
  -SendEmail `
  -SmtpServer "smtp-relay.yourcorp.local" `
  -SmtpPort 25 `
  -UseSsl:$false `
  -MailFrom "azure-tags@yourdomain.com" `
  -MailTo "netops@yourdomain.com","secops@yourdomain.com" `
  -VerboseOutput
```

3. Send via Microsoft 365 SMTP AUTH (TLS, port 587)
```
$pw = Read-Host "SMTP password" -AsSecureString

.\Watch-AzureServiceTags.ps1 `
  -SendEmail `
  -SmtpServer "smtp.office365.com" `
  -SmtpPort 587 `
  -UseSsl `
  -MailFrom "azure-tags@yourdomain.com" `
  -MailTo "netops@yourdomain.com","secops@yourdomain.com" `
  -SubjectPrefix "[Azure Tags]" `
  -SmtpUser "azure-tags@yourdomain.com" `
  -SmtpPassword $pw `
  -AttachAllReports `
  -SendOnNoChange
```

4. Adjust archive retention
```
.\Watch-AzureServiceTags.ps1 -ArchiveRetention 8
```

5. Run both feeds, no email (just write CSVs)
```
.\Watch-AzureServiceTags.ps1 `
  -MonitorAzureServiceTags:$true `
  -MonitorM365Worldwide:$true `
  -VerboseOutput
```

---------------------------------------------------------

## Filtering Regions

- New param: -AzureRegions (defaults to @("Global","us*")) — supports exact names and wildcards (case‑insensitive).
- The filter is applied to:
  - Azure CSV exports (diff_added_prefixes.csv, diff_removed_prefixes.csv, diff_tag_changes.csv)
  - Azure “Top 10” previews in HTML email
  - (If Teams posting is enabled) Azure previews there too
- Azure change counts in the email/Teams subject/body reflect the filtered results (so your summary aligns with the CSVs you care about).

Everything else remains: M365 monitoring via endpoints.office.com, archive retention, HTML email, and optional Teams webhook notification (MessageCard/AdaptiveCard).


---------------------------------------------------------

## Output files

All CSVs go to ReportsRoot (default C:\AzureServiceTags\reports):
```
azure_added_prefixes.csv       # Azure
azure_removed_prefixes.csv     # Azure
azure_tag_changes.csv          # Azure

M365\m365_urls_added.csv      # Microsoft 365
M365\m365_urls_removed.csv
M365\m365_ips_added.csv
M365\m365_ips_removed.csv
```
ArchiveRoot (default C:\AzureServiceTags\archive):
```
ServiceTags_Public_YYYYMMDD.json
...
M365\M365_Worldwide_YYYYMMDDHH.json
...
```
Older files are pruned to the newest N by -ArchiveRetention (default 5).

---------------------------------------------------------

## Email configuration notes

- TLS on/off
    - Port 25 relay: -SmtpPort 25 -UseSsl:$false
    - Port 587 auth/TLS: -SmtpPort 587 -UseSsl

- Credentials
    - Relay (no auth): omit -SmtpUser and -SmtpPassword
    - Auth: provide -SmtpUser and -SmtpPassword (securestring)
- Attachments
    - Default: attach only CSVs that have changes
    - Force all: add -AttachAllReports

---------------------------------------------------------
## MS TEAMS configuration notes

When you enable it, the script will post a notification to a Teams channel only if changes are detected (Azure and/or Microsoft 365). You can choose one of two payload styles:

  - MessageCard → for the legacy Incoming Webhook connector (if your tenant/channel still allows it).
  - AdaptiveCard → for the newer Teams Workflow “When a Teams webhook request is received” (recommended going forward).

```
-PostToTeams                       # turn Teams posting on
-TeamsPayload MessageCard|AdaptiveCard
-TeamsWebhookUrl "<the webhook URL for your channel/workflow>"
-TeamsTitlePrefix "[Cloud Endpoints]"      # optional label at top of the card
-TeamsPreviewLimit 10                     # how many “Top N” items to list in the card
```
The script builds a concise summary (counts) and a short preview list (top N) and posts it to Teams after the run completes.

### Option A - Use a legacy Incoming Webhook (MessageCard)

Works if your tenant/channel still allows adding the “Incoming Webhook” app (some tenants restrict new connectors). Payload format = MessageCard.

One‑time setup (in Teams / Channel):
  1. In the target channel, click … → Apps / Connectors → Incoming Webhook → Configure
  2. Give it a name (e.g., Cloud Endpoints Alerts), optionally upload an icon, then Create.
  3. Copy the Webhook URL (you only see it once). It looks like:
```
https://outlook.office.com/webhook/…/IncomingWebhook/…
```

Run the script:
```
.\Watch-AzureServiceTags.ps1 `
  -PostToTeams `
  -TeamsPayload MessageCard `
  -TeamsWebhookUrl "https://outlook.office.com/webhook/…/IncomingWebhook/…" `
  -TeamsTitlePrefix "[Cloud Endpoints]" `
  -VerboseOutput
```
What you’ll see in Teams
A MessageCard with:
  - Title: e.g., [Cloud Endpoints] Changes
  - A facts table: “Azure new/old file, +/‑ counts”, “M365 +/‑ counts”
  - Bulleted preview lines (up to -TeamsPreviewLimit, default 10)

### Option B - Use a Teams Workflow Webhook (Adaptive Card) — Recommended

This is the new path: create a channel workflow “Post to a channel when a webhook request is received.” It gives you a webhook URL that expects a message with an Adaptive Card attachment. Our script builds exactly that.

One‑time setup (in Teams / Channel):
  1. In the target channel, choose Automate → Create a workflow (or Workflows from “Apps”).
  2. Pick Post to a channel when a webhook request is received.
  3. Finish the quick wizard (select Team/Channel, name it).
  4. Copy the webhook URL (looks like an Azure Logic Apps URL):
```
https://prod-xx.region.logic.azure.com:443/workflows/…/triggers/manual/paths/invoke?api-version=…&sp=…&sv=…&sig=…
```

Run the script:
```
.\Watch-AzureServiceTags.ps1 `
  -PostToTeams `
  -TeamsPayload AdaptiveCard `
  -TeamsWebhookUrl "https://prod-xx.region.logic.azure.com:443/workflows/.../invoke?..." `
  -TeamsPreviewLimit 5 `
  -VerboseOutput
```

What you’ll see in Teams
An Adaptive Card with:
  - A bold title and FactSet (key/value) section with change counts.
  - Optional bullets with the “Top N” changes.

### What goes into the Teams card?

The script composes a short summary using the run’s results:
  - Azure (filtered by your -AzureRegions allow‑list):
    +X / -Y prefixes; +A / -B tags (filename)
    plus top N adds/removes by ServiceTag/Region: Prefix.
  - M365 Worldwide:
    +X / -Y IPs; +A / -B URLs (filename)
    plus top N new/removed IPs/URLs by ServiceArea [Category]

You can adjust:
- Title prefix via -TeamsTitlePrefix
- Preview list length via -TeamsPreviewLimit
- Azure region scope via -AzureRegions @("Global","us*") (already in your updated script)

### Security & limits (practical notes)

- Message size: Teams (Incoming Webhook) enforces ~28 KB per message. The script checks size and will error if exceeded. If you ever hit it, lower -TeamsPreviewLimit.
- Rate limits: The sender retries with exponential backoff on HTTP 429/5xx.
- Scope: A webhook URL is scoped to the channel you created it in—post only what you intend to that channel.

### Common run examples

Post to Teams only on change (default behavior):
```
.\Watch-AzureServiceTags.ps1 `
  -PostToTeams `
  -TeamsPayload AdaptiveCard `
  -TeamsWebhookUrl "<workflow-url>"
```
Include email + Teams, attach all CSVs:
```
$pw = Read-Host "SMTP password" -AsSecureString
.\Watch-AzureServiceTags.ps1 `
  -SendEmail -SmtpServer "smtp.office365.com" -SmtpPort 587 -UseSsl `
  -MailFrom "cloud-endpoints@yourdomain.com" -MailTo "netops@yourdomain.com" `
  -SmtpUser "cloud-endpoints@yourdomain.com" -SmtpPassword $pw `
  -AttachAllReports `
  -PostToTeams -TeamsPayload AdaptiveCard -TeamsWebhookUrl "<workflow-url>"
```
Filter Azure to Global + US only, then post to Teams:
```
.\Watch-AzureServiceTags.ps1 `
  -AzureRegions @("Global","us*") `
  -PostToTeams -TeamsPayload MessageCard -TeamsWebhookUrl "<incoming-webhook-url>"
```

### Troubleshooting

- Nothing posts: The script posts only when changes are detected. Use -SendOnNoChange if you want a weekly “heartbeat” email; if you want a Teams heartbeat too, I can add -TeamsOnNoChange.
- Adaptive Card workflow errors: Make sure you used the workflow URL (Logic Apps style) and selected -TeamsPayload AdaptiveCard.
- Incoming Webhook not allowed: Your tenant may block new Microsoft 365 connectors. In that case, use the Workflow Webhook route.

---------------------------------------------------------

## Archive retention

- The script keeps the newest N JSON files in the archive (-ArchiveRetention, default 5).
- Older files are automatically pruned each run.
- Retention happens after a new download.

---------------------------------------------------------

## Scheduling (Task Scheduler)

1. Open Task Scheduler → Create Task…
2. General: Run whether user is logged on or not; use a service account
3. Triggers: Weekly (e.g., Monday 08:00)
        Microsoft updates the list weekly; new ranges won’t be used for at least a week.
4. Actions:
    - Program/script:
      powershell.exe
    - Add arguments:
```
-ExecutionPolicy Bypass -File "C:\Scripts\Watch-AzureServiceTags.ps1" -SendEmail -SmtpServer "smtp-relay.yourcorp.local" -SmtpPort 25 -UseSsl:$false -MailFrom "azure-tags@yourdomain.com" -MailTo "netops@yourdomain.com"
```
5. Settings: Stop the task if it runs longer than e.g., 1 hour; allow run on demand.

---------------------------------------------------------

## Troubleshooting

### “The remote certificate is invalid according to the validation procedure.”

- You’re attempting TLS to a server with an untrusted or mismatched cert.
- For a plain relay on port 25: force no TLS → -SmtpPort 25 -UseSsl:$false
- If you require TLS:
    - Ensure -SmtpServer matches the certificate CN/SAN
    - Install the issuing CA/intermediate certificates on the sender
    - Use a valid server name and not an IP

### “Could not find JSON URL on confirmation page.”

- The confirmation page structure may have changed or transiently failed.
- Try again; ensure the host has outbound internet access.
- If it persists, we can add a fallback static URL (less reliable over time).

### No email received

- Verify SMTP connectivity from the host (firewall rules, DNS, port)
- For M365 SMTP AUTH:
    - SMTP AUTH must be enabled on the mailbox/app account
    - Use smtp.office365.com:587 with -UseSsl and valid creds
- Check the script output for exceptions

### No CSV changes even after weekly runs

- Some weeks have no changes.
- Confirm the archive has multiple JSONs with different dates.
- Use -VerboseOutput to see which two files are being compared.

---------------------------------------------------------

## Security considerations

- Credentials: If using -SmtpUser/-SmtpPassword, prefer providing the password as a secure string:
```
$pw = Read-Host "SMTP password" -AsSecureString
```
- Avoid bypassing certificate validation in code. Fix trust/DNS or disable TLS on relay if that’s the intended design.
- Run under a least‑privileged service account for scheduled tasks.

---------------------------------------------------------

## Customization ideas

- ~~Teams/Slack notifications: Add a webhook post when changes are detected.~~ **Completed**
- ~~HTML email: Convert the summary to HTML for richer formatting in Outlook.~~ **Completed**
- ~~filtered to US regions and Global.~~ **Completed**
- Syslog/EDL export: Emit formats tailored for your firewall (e.g., Palo Alto EDL TXT).
- Logging: Write a daily log file with start/end time, counts, and actions (including prune events).




