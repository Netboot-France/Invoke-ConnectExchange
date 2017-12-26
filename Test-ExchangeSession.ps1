Function Test-ExchangeSession {
	<#  
	.SYNOPSIS  
		Simple function to test exchange session

	.NOTES  
		File Name  : Test-ExchangeSession.ps1
		Author     : Thomas ILLIET, contact@thomas-illiet.fr
		Date	   : 2017-08-01
		Last Update: 2017-08-01
		Test Date  :
		Version	   : 1.0.0 
		
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
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $Reconnect){
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