Function Invoke-ConnectExchange {
    <#  
        .SYNOPSIS  
            Simple function to connect into Exchange

        .DESCRIPTION
            This function will allow you to connect with Exchange using PowerShell.

        .NOTES
            Author       : Thomas ILLIET, contact@thomas-illiet.fr
            Date         : 2017-08-01
            Last Update  : 2017-08-01
            Test Date    : 2017-10-17
            Version      : 2.1.0
        
        .PARAMETER config
            Configuration Array

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

    Try {

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
        Write-Debug "Attempting Connection to Exchange Online with $Methode"
        $SessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck 
        $Session = New-PSSession -Name $Config.SessionName -ConfigurationName Microsoft.Exchange -ConnectionUri $Config.ConnectionUri -Credential $Credential -Authentication $Config.Authentication -SessionOption $SessionOptions -AllowRedirection
        
        # Import Session
        Write-Debug "Importing Exchange Session"
        if(-not([string]::IsNullOrEmpty($Config.Cmdlet)))
        {
            Import-Module (Import-PSSession $Session -AllowClobber -CommandName $Config.Cmdlet -DisableNameChecking ) -Global -DisableNameChecking
        } else {
            Import-Module (Import-PSSession $Session -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        } # END Cmdlet selection

        # Error Management
        If ($error)
        {
            Write-warning "Unable to import Exchange PS session : $Error"
            return $false
        }#END Error
        else
        {
            Write-Debug "Connected to Exchange"

            # Set the Start time for the current session 
            Set-Variable -Scope 'Global' -Name 'ExchangeSessionStartTime' -Value (Get-Date)

            return $true
        }#END Success
    } Catch {
        Write-Error "Unable to connect to Exchange $_"
        return $false
    }
}

Function Test-ExchangeSession {
    <#  
    .SYNOPSIS  
        Simple function to test exchange session

    .NOTES
        Author      : Thomas ILLIET, contact@thomas-illiet.fr
        Date        : 2017-08-01
        Last Update : 2017-08-01
        Test Date   : 2018-01-10
        Version     : 1.1.0 
        
    .PARAMETER Config
        Exchange connexion config

    .PARAMETER ManualThrottle
        Manual throttle value then sleep for that many milliseconds

    .PARAMETER ActiveThrottle
        Amount of time gt our reset seconds then tear the session down and recreate it

    .PARAMETER Reconnect

    .EXAMPLE
        $ExchangeConfig =@{
            Identity       = "unicorn@microsoft.fr"
            Password       = "BeatifullUnicorne!"
            Authentication = "Basic"
            ConnectionUri  = "https://outlook.office365.com/powershell-liveid/"
            Cmdlet         = @('Set-Mailbox')
            SessionName    = "Exchange"
        }
        Test-ExchangeSession -Config $ExchangeConfig -Reconnect 500

    #>
    Param (
       [parameter(Mandatory=$true)]
        [Array]$Config,
        [parameter(Mandatory=$false)]
        [int]$ManualThrottle = 0,
        [parameter(Mandatory=$false)]
        [double]$ActiveThrottle = .25,
        [parameter(Mandatory=$false)]
        [int]$Reconnect = 870
    )
    
    # Get the time that we are working on this object to use later in testing
    $ObjectTime = Get-Date
    
    # Reset and regather our session information
    $SessionInfo = $null
    $SessionInfo = Get-PSSession -Name $Config.SessionName -ErrorAction SilentlyContinue
    
    # Make sure we found a session
    if ($SessionInfo -eq $null) {
        write-debug "| ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        write-debug "| + Test Exchange Session"
        write-debug "| ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        Write-Debug "| + No Session Found"
        Write-Debug "| + Recreating Session"
        Invoke-ConnectExchange -Config $ExchangeConfig
    }	

    # Make sure it is in an opened state if not log and recreate
    elseif ($SessionInfo.State -ne "Opened"){
        write-debug "| ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        write-debug "| + Test Exchange Session"
        write-debug "| ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        Write-Debug "| + Session not in Open State"
        Write-Debug "| + Recreating Session"
        Get-PSSession -Name $Config.SessionName | Remove-PSSession -Confirm:$false
        New-Sleep 10 "Waitung for reconnect..."
        Invoke-ConnectExchange -Config $ExchangeConfig

    }

    # If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
    elseif (($ObjectTime - $ExchangeSessionStartTime    ).totalseconds -gt $Reconnect){
        write-debug "| ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        write-debug "| + Test Exchange Session"
        write-debug "| ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        Write-Debug "| + Session has been active for greater than given number of seconds"
        Write-Debug "| + Rebuilding Connection"
        
        # Estimate the throttle delay needed since the last session rebuild
        # Amount of time the session was allowed to run * our activethrottle value
        # Divide by 2 to account for network time, script delays, and a fudge factor
        # Subtract 15s from the results for the amount of time that we spend setting up the session anyway
        [int]$DelayinSeconds = ((($Reconnect * $ActiveThrottle) / 2) - 15)
        
        # If the delay is >15s then sleep that amount for throttle to recover
        if ($DelayinSeconds -gt 15){
            Write-Debug "| + Sleeping some addtional seconds to allow throttle recovery"
            New-Sleep $DelayinSeconds "Sleeping some addtional seconds to allow throttle recovery"
        }
                
        # new O365 session and reset our object processed count
        Get-PSSession -Name $Config.SessionName | Remove-PSSession -Confirm:$false
        New-Sleep 15 "Waitung for reconnect..."
        Invoke-ConnectExchange -Config $ExchangeConfig

    }
    else {
        # If session is active and it hasn't been open too long then do nothing and keep going
    }
    
    # If we have a manual throttle value then sleep for that many milliseconds
    if ($ManualThrottle -gt 0){
        Write-Debug "| + Sleeping $ManualThrottle milliseconds"
        Start-Sleep -Milliseconds $ManualThrottle
    }
}

function New-Sleep {
    <#  
        .SYNOPSIS  
            Suspends the activity in a script or session for the specified period of time.
        .DESCRIPTION
            The New-Sleep cmdlet suspends the activity in a script or session for the specified period of time.
            You can use it for many tasks, such as waiting for an operation to complete or pausing before repeating an operation.
        .NOTES  
            File Name   : New-Sleep.ps1
            Author      : Thomas ILLIET, contact@thomas-illiet.fr
            Date        : 2017-05-10
            Last Update : 2018-01-08
            Version     : 1.0.2
        .PARAMETER S
            Time to wait
        .PARAMETER Message
            Message you want to display
        .EXAMPLE  
            New-Sleep -S 60 -Message "wait and see"
        .EXAMPLE
            New-Sleep -S 60
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)]
        [int]$S,
        [parameter(Mandatory=$false)]
        [string]$Message="Wait"
    )
    for ($i = 1; $i -lt $s; $i++) 
    {
        if ($host.ui.RawUi.KeyAvailable) { # Cancel waiting if CTRL+Q is pressed
            $key = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyUp") 
            if (($key.VirtualKeyCode -eq 81) -AND ($key.ControlKeyState -match "LeftCtrlPressed"))
            {
                break
            }
        }
        [int]$TimeLeft = $s - $i
        Write-Progress -Activity $message -PercentComplete (100 / $s * $i) -CurrentOperation "$TimeLeft seconds left" -Status "Please wait (Cancel with CTRL + Q)"
        Start-Sleep -s 1
    }
    Write-Progress -Completed $true -Status "Please wait"
}