<#
.SYNOPSIS
    Extract Crewhu notification emails from Outlook and save as JSON.

.DESCRIPTION
    Reads emails from a specified Outlook folder, decodes Microsoft SafeLinks,
    and exports the data to a JSON file on the Desktop.

.PARAMETER FolderName
    The name of the folder to read emails from.
    Use "Inbox" for the main inbox, or specify a subfolder name like "Tickets".
    Default: "Inbox"

.PARAMETER OutputFileName
    The name of the output JSON file (saved to Desktop).
    Default: "crewhu_notifications_clean.json"

.EXAMPLE
    .\emailParse.ps1
    # Reads from Inbox, saves to crewhu_notifications_clean.json

.EXAMPLE
    .\emailParse.ps1 -FolderName "Tickets"
    # Reads from a folder named "Tickets"

.EXAMPLE
    .\emailParse.ps1 -FolderName "Tickets" -OutputFileName "tickets_export.json"
    # Custom folder and output file name
#>

param(
    [string]$FolderName = "Inbox",
    [string]$OutputFileName = "crewhu_notifications_clean.json"
)

# --------------------------------------------
# Configuration
# --------------------------------------------
$TargetEmail  = "michael.crawford@pearlsolves.com"
$TargetSender = "notification.system@notification.crewhu.com"

# Load System.Web for URL decoding
Add-Type -AssemblyName System.Web

# --------------------------------------------
# Function: Decode Microsoft SafeLinks
# --------------------------------------------
function Decode-SafeLinks {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    # Regex to match Microsoft SafeLinks URLs
    $pattern = 'https://[^ \r\n>"]*safelinks\.protection\.outlook\.com/\?[^ \r\n>"]*'

    return [System.Text.RegularExpressions.Regex]::Replace($Text, $pattern, {
        param($m)

        $value = $m.Value

        # Extract url= param
        if ($value -match 'url=([^&]+)') {
            $encoded = $matches[1]
            try {
                $decoded = [System.Web.HttpUtility]::UrlDecode($encoded)

                if (-not [string]::IsNullOrWhiteSpace($decoded)) {
                    return $decoded
                }
            } catch {
                # If decode fails, fall back to original SafeLink
                return $value
            }
        }

        return $value
    })
}

# --------------------------------------------
# Function: Get target folder from Outlook
# --------------------------------------------
function Get-TargetFolder {
    param(
        [object]$Namespace,
        [string]$TargetEmail,
        [string]$FolderName
    )

    # Find the mailbox store
    $Store = $Namespace.Folders | Where-Object { $_.Name -eq $TargetEmail }

    if (-not $Store) {
        Write-Host "Could not find mailbox for $TargetEmail. Using default store."
        $Store = $Namespace.GetDefaultFolder(6).Parent  # 6 = olFolderInbox
    }

    # If requesting Inbox, return it directly
    if ($FolderName -eq "Inbox") {
        $Inbox = $Store.Folders | Where-Object { $_.Name -eq "Inbox" }
        if ($Inbox) {
            return $Inbox
        }
        # Fallback to default inbox
        return $Namespace.GetDefaultFolder(6)
    }

    # Otherwise, look for the specified folder
    # First check root level
    $TargetFolder = $Store.Folders | Where-Object { $_.Name -eq $FolderName }

    if (-not $TargetFolder) {
        # Check inside Inbox
        Write-Host "Folder '$FolderName' not found in root. Checking Inbox..."
        $Inbox = $Store.Folders | Where-Object { $_.Name -eq "Inbox" }

        if ($Inbox) {
            $TargetFolder = $Inbox.Folders | Where-Object { $_.Name -eq $FolderName }
        }
    }

    return $TargetFolder
}

# --------------------------------------------
# Main Script
# --------------------------------------------
Write-Host "Crewhu Email Parser"
Write-Host "==================="
Write-Host "Target folder: $FolderName"
Write-Host "Output file:   $OutputFileName"
Write-Host ""

# Connect to Outlook
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Get the target folder
$Folder = Get-TargetFolder -Namespace $Namespace -TargetEmail $TargetEmail -FolderName $FolderName

if (-not $Folder) {
    Write-Host "ERROR: Could not find folder '$FolderName' in mailbox."
    exit 1
}

Write-Host "Using folder: $($Folder.Name)"

# Get and sort items
$Items = $Folder.Items
$Items.Sort("ReceivedTime", $false)  # Newest first

$Output = @()
$ProcessedCount = 0
$SkippedCount = 0

foreach ($msg in $Items) {
    try {
        # Only process actual mail items
        if ($msg.MessageClass -ne "IPM.Note") {
            continue
        }

        # Filter by sender
        if ($msg.SenderEmailAddress -ne $TargetSender) {
            continue
        }

        # Get original body and decode SafeLinks
        $rawBody = $msg.Body
        $cleanBody = Decode-SafeLinks -Text $rawBody

        # Create a preview safely
        if ([string]::IsNullOrEmpty($cleanBody)) {
            $preview = ""
        } else {
            $len     = [Math]::Min(300, $cleanBody.Length)
            $preview = $cleanBody.Substring(0, $len)
        }

        $item = [PSCustomObject]@{
            Subject      = $msg.Subject
            Sender       = $msg.SenderEmailAddress
            ReceivedTime = $msg.ReceivedTime
            BodyPreview  = $preview
            FullBody     = $cleanBody
        }

        $Output += $item
        $ProcessedCount++
    }
    catch {
        Write-Host "Skipped one message (error accessing item)"
        $SkippedCount++
    }
}

# --------------------------------------------
# Save JSON to Desktop
# --------------------------------------------
$Desktop  = [Environment]::GetFolderPath("Desktop")
$JsonPath = Join-Path $Desktop $OutputFileName

$Output | ConvertTo-Json -Depth 10 | Out-File -Encoding UTF8 $JsonPath

Write-Host ""
Write-Host "Done!"
Write-Host "JSON saved to: $JsonPath"
Write-Host "Emails matched: $($Output.Count)"
if ($SkippedCount -gt 0) {
    Write-Host "Emails skipped: $SkippedCount"
}
