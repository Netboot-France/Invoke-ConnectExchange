#----------------------------------------------
#  Connect to exchange with Plain Password
#----------------------------------------------
$ExchangeConfig =@{
    Identity       = "unicorn@microsoft.fr"
    Password       = "BeatifullUnicorne!"
    Authentication = "Basic"
    ConnectionUri  = "https://outlook.office365.com/powershell-liveid/"
    Cmdlet         = ''
    SessionName    = "Exchange"
}

if(Invoke-ConnectExchange -Config $ExchangeConfig -eq $true) {

    $Bals =  get-mailbox -ResultSize unlimited
    foreach($Bal in $Bals) {

        # Test Exchange Session
        Test-ExchangeSession -Config $ExchangeConfig -Reconnect 500

        # Get information
        $Size = (Get-MailboxStatistics -Identity $Bal).TotalItemSize.Value
        Write-Host "$($Bal.UserPrincipalName) - $Size"
    }
    write-host "Connected"
} else {
    Write-Host "Error to connect !!!"
}