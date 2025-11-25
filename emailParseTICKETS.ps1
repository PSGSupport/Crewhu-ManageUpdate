# --------------------------------------------
# Read Outlook emails from a folder named "Tickets"
# Decode Microsoft SafeLinks to original URLs
# Save JSON to Desktop
# --------------------------------------------

$TargetEmail  = "michael.crawford@pearlsolves.com"
$TargetSender = "notification.system@notification.crewhu.com"

# Make sure we can use HttpUtility.UrlDecode
Add-Type -AssemblyName System.Web

function Decode-SafeLinks {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    # Match Microsoft SafeLinks
    $pattern = 'https://[^ \r\n>"]*safelinks\.protection\.outlook\.com/\?[^ \r\n>"]*'

    return [System.Text.RegularExpressions.Regex]::Replace($Text, $pattern, {
        param($m)

        $value = $m.Value

        if ($value -match 'url=([^&]+)') {
            $encoded = $matches[1]
            try {
                $decoded = [System.Web.HttpUtility]::UrlDecode($encoded)

                if (-not [string]::IsNullOrWhiteSpace($decoded)) {
                    return $decoded
                }
            } catch {}

            return $value
        }

        return $value
    })
}

# --------------------------------------------
# Connect to Outlook
# --------------------------------------------
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Store     = $Namespace.Folders | Where-Object { $_.Name -eq $TargetEmail }

if (-not $Store) {
    Write-Host " Could not find mailbox for $TargetEmail. Using default store."
    $Store = $Namespace.GetDefaultFolder(6).Parent
}

# --------------------------------------------
# Locate the "Tickets" folder
# --------------------------------------------
$TicketsFolder = $Store.Folders | Where-Object { $_.Name -eq "Tickets" }

if (-not $TicketsFolder) {
    Write-Host " Folder 'Tickets' not found in root. Checking Inbox..."
    $Inbox = $Store.Folders | Where-Object { $_.Name -eq "Inbox" }

    if ($Inbox) {
        $TicketsFolder = $Inbox.Folders | Where-Object { $_.Name -eq "Tickets" }
    }
}

if (-not $TicketsFolder) {
    Write-Host " Could not find a folder named 'Tickets' in mailbox."
    exit
}

Write-Host " Using folder: $($TicketsFolder.Name)"

$Items = $TicketsFolder.Items
$Items.Sort("ReceivedTime", $false)  # Newest first

$Output = @()

foreach ($msg in $Items) {
    try {
        if ($msg.MessageClass -ne "IPM.Note") { continue }

        if ($msg.SenderEmailAddress -ne $TargetSender) { continue }

        $rawBody = $msg.Body
        $cleanBody = Decode-SafeLinks -Text $rawBody

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
    }
    catch {
        Write-Host " Skipped one message due to error."
    }
}

# --------------------------------------------
# Save JSON to Desktop
# --------------------------------------------
$Desktop  = [Environment]::GetFolderPath("Desktop")
$JsonPath = Join-Path $Desktop "crewhu_notifications_clean.json"

$Output | ConvertTo-Json -Depth 10 | Out-File -Encoding UTF8 $JsonPath

Write-Host ""
Write-Host " Done!"
Write-Host " JSON saved to:"
Write-Host "   $JsonPath"
Write-Host " Emails matched: $($Output.Count)"
