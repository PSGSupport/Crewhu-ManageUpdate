# --------------------------------------------
# Read Outlook emails for a specific sender
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
                # If decode fails, just fall back to original SafeLink
                return $value
            }
        }

        return $value
    })
}

# --------------------------------------------
# Connect to Outlook
# --------------------------------------------
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Try to find the account (optional safety check)
$Account = $Namespace.Accounts | Where-Object { $_.SmtpAddress -eq $TargetEmail }

if (-not $Account) {
    Write-Host "⚠ Could not find Outlook account for $TargetEmail. Using default Inbox."
}

# 6 = olFolderInbox
$Inbox  = $Namespace.GetDefaultFolder(6)
$Items  = $Inbox.Items
$Items.Sort("ReceivedTime", $false)  # Newest first

$Output = @()

foreach ($msg in $Items) {
    try {
        # Only process actual mail items
        if ($msg.MessageClass -ne "IPM.Note") { continue }

        # Filter by sender
        if ($msg.SenderEmailAddress -ne $TargetSender) { continue }

        # Get original body
        $rawBody = $msg.Body

        # Decode any SafeLinks inside the body
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
    }
    catch {
        Write-Host " Skipped one message (error accessing item)"
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
