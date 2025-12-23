# PowerShell script to list which shared mailboxes, MS365 groups and distribution lists an account has access to.  
# The command line '-Mailbox user@domain.com' setting is required.
# Sacha Panasuik
# https://github.com/sachapan/365

param(
    [Parameter(Mandatory=$true)]
    [string]$Mailbox
)

# Shared mailboxes (FullAccess, SendAs)
Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox |
ForEach-Object {
    $mbx = $_
    Get-MailboxPermission -Identity $mbx.Identity |
    Where-Object {
        $_.User -like $Mailbox -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited
    } |
    Select-Object @{n="Type";e={"SharedMailbox"}},
                  @{n="Object";e={$mbx.PrimarySmtpAddress}},
                  @{n="Access";e={"FullAccess"}}

    Get-RecipientPermission -Identity $mbx.Identity |
    Where-Object {
        $_.Trustee -like $Mailbox -and $_.AccessRights -contains "SendAs"
    } |
    Select-Object @{n="Type";e={"SharedMailbox"}},
                  @{n="Object";e={$mbx.PrimarySmtpAddress}},
                  @{n="Access";e={"SendAs"}}
}

# Distribution Lists
Get-DistributionGroup -ResultSize Unlimited |
Where-Object {
    (Get-DistributionGroupMember $_.Identity -ResultSize Unlimited |
     Select-Object -ExpandProperty PrimarySmtpAddress) -contains $Mailbox
} |
Select-Object @{n="Type";e={"DistributionList"}},
              @{n="Object";e={$_.PrimarySmtpAddress}},
              @{n="Access";e={"Member"}}

# Microsoft 365 Groups
Get-UnifiedGroup -ResultSize Unlimited |
Where-Object {
    (Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members |
     Select-Object -ExpandProperty PrimarySmtpAddress) -contains $Mailbox
} |
Select-Object @{n="Type";e={"M365Group"}},
              @{n="Object";e={$_.PrimarySmtpAddress}},
              @{n="Access";e={"Member"}}
