 Function Connect-Exchange {
    <#  
        .SYNOPSIS  
            Simple function to connect into Exchange

        .DESCRIPTION
            This function will allow you to connect with Exchange using PowerShell.

        .NOTES  
            File Name    : Connect-Exchange.ps1
            Author       : Thomas ILLIET, contact@thomas-illiet.fr
            Date	     : 2017-08-01
            Last Update  : 2017-08-01
    	    Test Date    : 2017-10-17
            Version	     : 2.0.0

        .REQUIRE
            Function :
                + New-sleep
                    - https://github.com/thomas-illiet/Powershell/tree/master/3-Tools/2-New-Sleep
                + Test-ExchangeSession
                    - https://github.com/thomas-illiet/Powershell/tree/master/2-Office365/2-ExchangeOnline/1-Connect

        .PARAMETER
            Config : Configuration Array

        .EXAMPLE 
            #----------------------------------------------
            #  With Plain Password
            #----------------------------------------------
            $ExchangeConfig =@{
                Identity       = "unicorn@microsoft.fr"
                Password       = "BeatifullUnicorne!"
                Authentication = "Basic"
                ConnectionUri  = "https://outlook.office365.com/powershell-liveid/"
                Cmdlet         = @('Set-Mailbox')
                SessionName    = "Exchange"
            }
            Connect-Exchange -Config $ExchangeConfig

        .EXAMPLE
            #----------------------------------------------
            # With SecureString file
            #----------------------------------------------
            $ExchangeConfig =@{
                Identity       = "unicorn@microsoft.fr"
                PasswordFile   = "c:\Securestring.txt"
                Authentication = "Basic"
                ConnectionUri  = "https://outlook.office365.com/powershell-liveid/"
                Cmdlet         = @('Set-Mailbox')
                SessionName    = "Exchange"
            }
            Connect-Exchange -Config $ExchangeConfig

        .EXAMPLE
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
            Connect-Exchange -Config $ExchangeConfig

    #>
    Param (
        [parameter(Mandatory=$true)]
        [Array]$Config
    )

    write-debug "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    write-debug "+ Connect to Exchange"
    write-debug "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    
    # Load Credential
    if(-not([string]::IsNullOrEmpty($Config.PasswordFile)))
    {
        if(-not([string]::IsNullOrEmpty($Config.KeyFile)))
        {
            $Methode = "SecureString file + Key"
            $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Config.Identity, (Get-Content $Config.PasswordFile | ConvertTo-SecureString -Key (Get-Content $Config.KeyFile))
        } # END Credential with Key File
        else
        {
            $Methode = "SecureString file"
            $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Config.Identity, (Get-Content $Config.PasswordFile | ConvertTo-SecureString)
        } # END Credential without Key File
    }
    else
    {
        $Methode = "Plain password"
        $Secpasswd = ConvertTo-SecureString $Config.Password -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($Config.Identity, $secpasswd)
    } # END Credential with plain password

    # Connect to Exchange
	$error.clear();
    Write-Debug "| + Attempting Connection to Exchange Online with $Methode"
    $SessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck 
	$Session = New-PSSession -Name $Config.SessionName -ConfigurationName Microsoft.Exchange -ConnectionUri $Config.ConnectionUri -Credential $Credential -Authentication $Config.Authentication -SessionOption $SessionOptions -AllowRedirection -ErrorAction SilentlyContinue
    
    
    # Import Session
    if(-not([string]::IsNullOrEmpty($Config.Cmdlet)))
    {
        Import-PSSession $Session -AllowClobber -CommandName $Config.Cmdlet | Out-Null
    } else {
        Import-PSSession $Session -AllowClobber| Out-Null
    } # END Cmdlet selection

    # Error Management
	If ($error)
	{
		Write-warning "| + Unable to import Exchange PS session : $Error"
		return $false
	}#END Error
	else
	{
        Write-Debug "| + Connected to Exchange"

        # Set the Start time for the current session
        Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
    
        return $true
    }#END Success
}