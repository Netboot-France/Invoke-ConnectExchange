#----------------------------------------------
# With SecureString file + Key
#----------------------------------------------
$ExchangeConfig =@{
    Identity       = "unicorn@microsoft.fr"
    PasswordFile   = "C:\Securestring.txt"
    KeyFile        = "C:\MyCertificat.key"
    Authentication = "Basic"
    ConnectionUri  = "https://outlook.office365.com/powershell-liveid/"
    Cmdlet         = @('Set-Mailbox')
    SessionName    = "Exchange"
}

if(Invoke-ConnectExchange -Config $ExchangeConfig -eq $true) {
    write-host "Connected"
} else {
    Write-Host "Error"
}