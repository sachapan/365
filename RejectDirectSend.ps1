Connect-ExchangeOnline
# Disable Direct Send
Set-OrganizationConfig -RejectDirectSend $true
# display current setting
Get-OrganizationConfig | Select-Object Identity, RejectDirectSend
