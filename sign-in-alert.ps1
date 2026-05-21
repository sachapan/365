New-ProtectionAlert -Name "[SECURITY ALERT] Break-Glass Account Sign-In Detected" `
    -Category Others `
    -NotifyUser SOCalerts@contoso.com `
    -ThreatType Activity `
    -Operation UserLoggedIn `
    -Description "Alert when admin logs in" `
    -AggregationType None `
    -Filter "Activity.UserId -eq 'Admin.GA@contoso.com'"   
