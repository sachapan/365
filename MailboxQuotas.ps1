# This script locates mailboxes nearing their quota limit
# Sacha Panasuik
# https://github.com/sachapan/365

param(
    [switch]$ReportOnly,           # If specified, only output results to console, no email
    [string]$OutputFile,           # Optional: Path to save CSV report (used with -ReportOnly)
    [int]$ThresholdPercent = 90,   # Alert threshold
    [string]$AdminEmail = "admin@yourdomain.com",
    [string]$SmtpServer = "your-smtp-relay.yourdomain.com",
    [string]$FromEmail = "alerts@yourdomain.com"
)

# Connect to Exchange Online
Connect-ExchangeOnline

# Get all user mailboxes
$Mailboxes = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

$AlertMailboxes = @()

foreach ($Mailbox in $Mailboxes) {
    $Stats = Get-ExoMailboxStatistics -Identity $Mailbox.UserPrincipalName
    
    # Handle unlimited quotas or missing data
    $ItemSizeBytes = if ($Stats.TotalItemSize) {
        ($Stats.TotalItemSize.ToString() -split '\(')[1] -replace '[^\d]', '' -as [long]
    }
    else { 0 }

    $QuotaBytes = if ($Mailbox.ProhibitSendReceiveQuota -eq "Unlimited") { 
        [long]::MaxValue 
    }
    else {
        ($Mailbox.ProhibitSendReceiveQuota -split '\(')[1] -replace '[^\d]', '' -as [long]
    }
    
    if ($QuotaBytes -eq 0) { continue }
    
    $PercentUsed = [math]::Round(($ItemSizeBytes / $QuotaBytes) * 100, 2)
    
    if ($PercentUsed -ge $ThresholdPercent) {
        $AlertMailboxes += [PSCustomObject]@{
            DisplayName        = $Mailbox.DisplayName
            PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
            SizeGB             = [math]::Round($ItemSizeBytes / 1GB, 2)
            QuotaGB            = if ($QuotaBytes -eq [long]::MaxValue) { "Unlimited" } else { [math]::Round($QuotaBytes / 1GB, 2) }
            PercentUsed        = $PercentUsed
        }
    }
}

# Output results
if ($AlertMailboxes.Count -gt 0) {
    if ($ReportOnly) {
        # Output to console
        $AlertMailboxes | Format-Table -AutoSize
        
        # Optional: Export to CSV if -OutputFile is provided
        if ($OutputFile) {
            $AlertMailboxes | Export-Csv -Path $OutputFile -NoTypeInformation
            Write-Host "Report saved to $OutputFile"
        }
    }
    else {
        # Normal mode: Send email
        $Body = $AlertMailboxes | Format-Table -AutoSize | Out-String
        $Body = "The following mailboxes are at or above $ThresholdPercent% of their quota:`n`n$Body"
        
        Send-MailMessage -From $FromEmail -To $AdminEmail -Subject "Mailbox Quota Alert: Near Limit" `
            -Body $Body -SmtpServer $SmtpServer
        Write-Host "Alert email sent to $AdminEmail"
    }
}
else {
    Write-Host "No mailboxes are near their quota limit ($ThresholdPercent%)."
}

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
