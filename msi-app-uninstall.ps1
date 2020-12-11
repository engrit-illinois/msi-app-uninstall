param(
	[string]$LogDir = "c:\engrit\logs",
	[string]$ScriptLogFile = "msi-app-uninstall-script.log",
	[string]$MsiLogFile = "msi-app-uninstall-msi.log",
	
	# Comma-delimted string of services to stop before attempting anything.
	# e.g. "bes client"
	[string]$StopServices = "false",
	
	# Set to "true" to exit without further action if specified services were not successfully stopped.
	# Cannot use a [switch] parameter because SCCM Console does not recognize those.
	# Has no effect if $StopServices is "false"
	[string]$RequireServicesToStop = "false",
	
	# Comma-delimited string of directory paths to remove after uninstallation.
	# e.g. "c:\dir1,c:\dir2\test"
	[string]$RemoveDirs = "false",
	
	# Common-delimited string of registry keys and/or values to remove after uninstallation.
	# For keys, just give the path. All subkeys and values will be removed. For individual values, give "path;value".
	# e.g. "hklm:\software\bigfix\enterpriseclient,hklm:\software\wow6432node\bigfix\enterpriseclient;version"
	[string]$RemoveRegKeysValues = "false",
	
	# String to search for in display names of installed apps.
	# Asterisks can be used, e.g. "*bigfix*".
	# Make sure this only matches one or fewer apps
	# Originally this parameter was mandatory, but with the inclusion of $KnownApp below, I wanted one OR the other to be mandatory.
	# It turns out powershell doesn't have a good way to do this. The best solution I've found is to make all parameters optional and validate within the code, which is what I've done.
	[string]$AppTitleQuery,
	
	# Turns out SCCM (1802) is bugged to hell with issues revolving around parameter values being passed. Seems like A) spaces are not passed properly, B) commas are not passed properly, C) quotes are handled weirdly and not properly escaped, and D) most importantly there appears to be a limit to the number of characters that can be passed as arguments, which is critically restrictive. When any of these problems are present in the arguments specified in a script deployment, the script simply fails to run on the endpoints, giving virtually no error feedback at all other than a script output of "8", whatever the hell that means.
	# As a workaround, I will include the following parameter that will allow me to define app-specific parameters within the script and allow the script to be called using only this parameter.
	[string]$KnownApp
)

# Uninstall script by mseng3
# This script searches through the registry for an app with a display name that matches the given query string.
# If found, it pulls the app's MSI product code and uses msiexec to uninstall it.
# Optionally, directories can be specified for deletion as well, for cases where a product's uninstaller isn't thorough.
# Originally built to cleanly remove the BigFix Enterprise Client application

# Scripts are called by SCCM using the following syntax:
# "C:\Windows\system32\WindowsPowerShell\v1.0\PowerShell.exe" -NonInteractive -NoProfile -ExecutionPolicy RemoteSigned -Command "& { . 'C:\Windows\CCM\ScriptStore\<script guid>.ps1' -Param1 "value" -Param2 "value" | ConvertTo-Json -Compress } "
# Logs are located here: c:\windows\ccm\logs\scripts.log
# Script cache is here: c:\windows\ccm\scriptstore\<script guid>.ps1
# Source: https://www.systemcenterdudes.com/sccm-deploy-powershell-script/

# If you copy the above line to a PowerShell prompt, you'll notice that the quotation marks in the string of passed parameters and values are not escaped and thus break the command?
# Also, does this mean it really uses PowerShell v1.0?

# https://docs.microsoft.com/en-us/sccm/apps/deploy-use/create-deploy-scripts
# Parameters cannot contain apostrophes.
# Parameters with spaces do not work (known issue). Everything before the space is passed while everything after the space is not.
# To get around this, ":_:" can be passed in place of spaces and will be replaced with spaces.
# e.g. if you want to pass the string "c:\program files", you can instead optionally pass "c:\program:_:files"

# Exit codes
# A combination of powershell and SCCM quirks prevent actual exit codes from being passed back to the SCCM GUI.
# However the script output is passed back. So It's important to keep the script from producing any extraneous output you don't want passed back.
# All of my output is being sent to logs, except where I exit I both output the exit code and exit with the exit code.
# https://www.reddit.com/r/SCCM/comments/8kjnce/exit_codes_in_new_scripts_feature_of_sccm_cb/
$EXIT_UnknownApp = -2
$EXIT_InvalidParameters = -1
$EXIT_Success = 0
$EXIT_AppNotFound = 1
$EXIT_AppUninstallFailed = 2
$EXIT_AppUninstalledDirsRemovedRegsFailed = 3
$EXIT_AppUninstalledDirsFailedRegsRemoved = 4
$EXIT_AppUninstalledDirsFailedRegsFailed = 5
$EXIT_ServicesFailed = 6
$EXIT_AppTitleBlank2 = 7
$EXIT_AppTitleBlank3 = 8

# Validate parameters
$useKnownApp = $false
$savedLogs = ""
$invalidParams = $false
if($AppTitleQuery -eq "") {
	if($KnownApp -eq "") {
		$savedLogs = "$savedLogs;Either `$AppTitleQuery or `$KnownApp must be specified. Exiting with exit code $EXIT_InvalidParameters."
		$invalidParams = $true
	}
	else {
		$savedLogs = "$savedLogs;`$AppTitleQuery was not specified, but `$KnownApp was specified (`"$KnownApp`"). Continuing to use `$KnownApp."
		$useKnownApp = $true
	}
}
else {
	if($KnownApp -eq "") {
		$savedLogs = "$savedLogs;`$AppTitleQuery was specified (`"$AppTitleQuery`"), but `$KnownApp was not specified. Continuing to use `$AppTitleQuery."
	}
	else {
		$savedLogs = "$savedLogs;`$AppTitleQuery (`"$AppTitleQuery`") and `$KnownApp (`"$KnownApp`") were both specified. `$AppTitleQuery will be ignored."
		$useKnownApp = $true
	}
}

# If $KnownApp is specified, fill out the relevant parameters with appropriate info.
# Because apparently the SCCM scripts function doesn't handle passing this many parameters/characters without choking to death.
$unknownApp = $false
if($useKnownApp) {
	switch($KnownApp) {
		'bigfix' {
			$savedLogs = "$savedLogs;`$KnownApp (`"bigfix`") recognized. Setting appropriate parameters."
			$ScriptLogFile = "msi-app-uninstall-bigfix-script.log"
			$MsiLogFile = "msi-app-uninstall-bigfix-msi.log"
			$StopServices = "bes client"
			$RemoveDirs = "c:\program files\ibm\tivoli,c:\program files\tivoli,c:\program files (x86)\bigfix enterprise"
			$RemoveRegKeysValues = "hklm:\software\bigfix,hklm:\software\wow6432node\bigfix"
			$AppTitleQuery = "*bigfix*"
			break
		}
		default {
			$savedLogs = "$savedLogs;`$KnownApp was specified, but the string provided was not recognized as a known app. Exiting with exit code $EXIT_UnknownApp."
			$unknownApp = $true
		}
	}
}

# Log stuff
if(-not (Test-Path $LogDir -PathType Container)) {
	md $LogDir
}
$ScriptLog = "$LogDir\$ScriptLogFile"
$MsiLog = "$LogDir\$MsiLogFile"

$timestamp = Get-Date -UFormat "%Y-%m-%d %T"
echo "[$timestamp] Starting..." > $ScriptLog

function log ([string]$text = '') {
	$timestamp = Get-Date -UFormat "%Y-%m-%d %T"
	echo "[$timestamp] $text" >> $ScriptLog 2>&1
}

log "--------------"
log ""

# Output saved logs
$savedLogsArray = $savedLogs -split ';'
foreach($savedLog in $savedLogsArray) {
	log $savedLog
}
		
# Exit if invalid params or unknown app
if($invalidParams) {
	$EXIT_InvalidParameters
	exit $EXIT_InvalidParameters
}

if($unknownApp) {
	$EXIT_UnknownApp
	exit $EXIT_UnknownApp
}

# Reconstitute spaces in passed in parameter values
# Doubles as logging of passed parameters
# https://stackoverflow.com/questions/21559724/getting-all-named-parameters-from-powershell-including-empty-and-set-ones
log "Reconstituting spaces in passed-in parameter values..."
log ""
# https://www.petri.com/unraveling-mystery-myinvocation
foreach ($key in $MyInvocation.BoundParameters.keys) {
	$var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
	if($var) {
		log "Parameter: $($var.name)"
		log "Original value: $($var.value)"
		$var.value = $var.value -replace ':_:',' '
		log "New value: $($var.value)"
	}
	else {
		log "Parameter $key is null."
	}
	log ""
}

# Finding an MSI product code from an application name string
# https://stackoverflow.com/questions/5063129/how-to-find-the-upgradecode-and-productcode-of-an-installed-application-in-windo﻿﻿

# Set reg keys to look at
$Reg1 = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
$Reg2 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
log "Searching reg keys:"
log $Reg1
log $Reg2
log ""

$Reg = @($Reg1, $Reg2)

# Stop specified services
function StopTheServices {
	$success = $true
	log "Stopping specified services..."
	log ""
	$services = $StopServices -split ','
	foreach($service in $services) {
		log "Stopping `"$service`"..."
		Stop-Service -Name $service -Force >> $ScriptLog 2>&1 3>&1
		$serviceObj = Get-Service -Name $service
		if($serviceObj.status -eq 'Stopped') {
			log "Service stopped."
		}
		else {
			log "Failed to stop service."
			$success = $false
		}
		log ""
	}
	log "Done stopping services."
	
	return $success
}
		

# Modularized searching for the app, so it can be done more than once to check for successfull uninstallation
function CheckForApp {

	# Double check to make sure $AppTitle Query is populated, otherwise it might end up picking a random app
	if($AppTitleQuery -ne "") {
		
		# Get list of installed apps and their properties
		$InstalledApps = Get-ItemProperty $Reg

		log "All installed apps:"
		#log $InstalledApps
		log "[omitted due to verbosity]"
		log ""

		# Filter to desired app
		log "Looking for `"$AppTitleQuery`"..."
		log ""
		$TargetApp = $InstalledApps | Where { $_.DisplayName -like $AppTitleQuery }
		
		# TODO: add check to make sure $TargetApp matched only one app, otherwise abort with new exit code.

		if($TargetApp -eq $null) {
			$TargetApp = "Not found"
		}
		log "Target app:"
		log $TargetApp
		log ""

		return $TargetApp
	}
	else {
		log "`$AppTitleQuery not populated on 2nd check. Aborting and exiting with exit code $EXIT_AppTitleBlank2."
		$EXIT_AppTitleBlank2
		exit $EXIT_AppTitleBlank2
	}
}

function UninstallTheApp {

	param(
		[string]$ProductKey
	)
	
	# Uninstall
	log "Uninstalling..."
	log ""
	
	# Uninstall normally
	#$UninstallString

	# Uninstall silently with logging
	log "Uninstall output (also check $MsiLog):"
	#$UninstallString /qn /norestart /l*v $MsiLog >> $ScriptLog 2>&1
	
	# https://stackoverflow.com/questions/1741490/how-to-tell-powershell-to-wait-for-each-command-to-end-before-starting-the-next
	# https://www.zerrouki.com/wait-for-an-executable-to-finish/
	$process = "msiexec.exe"
	$args = "/x $ProductKey /qn /norestart /l*v $MsiLog"
	
	# Final check to make sure $AppTitle Query is populated, otherwise it might end up picking a random app
	if($AppTitleQuery -ne "") {
		Start-Process $process -ArgumentList $args -Wait
		#$job = Start-Job -filepath $process $args
		#Wait-Job $job
		#Receive-Job $job
	}	
	else {
		log "`$AppTitleQuery not populated on 2nd check. Aborting and exiting with exit code $EXIT_AppTitleBlank3."
		$EXIT_AppTitleBlank3
		exit $EXIT_AppTitleBlank3
	}
}

# Remove optionally specified directories
function RemoveTheDirs {
	$RemoveDirsSuccessful = $true
	if($RemoveDirs -ne "false") {
		log "Removing directories:"
		log ""
		$dirs = $RemoveDirs -split ','
		foreach($dir in $dirs) {
			log $dir
			if(Test-Path $dir -PathType Container) {
				log "Directory exists."
				# Using -Force fails on some files and not others. Sounds like a bug per below links.
				# To work around this, I will try deleting without -Force first and then with -Force
				# https://social.technet.microsoft.com/Forums/windows/en-US/98b70109-7f41-4f66-8fe8-d7e241ad5453/removeitem-force-fails-with-access-denied-not-without-force?forum=winserverpowershell
				# https://stackoverflow.com/questions/25606481/remove-item-doesnt-work-delete-does
				# https://windowsserver.uservoice.com/forums/301869-powershell/suggestions/11088240-remove-item-force-fails-when-remove-item-succeeds
				Remove-Item -Recurse $dir >> $ScriptLog 2>&1
				if(Test-Path $dir -PathType Container) {
					log "Directory failed to be removed. Retrying with -Force."
					Remove-Item -Recurse -Force $dir >> $ScriptLog 2>&1
				}
				if(Test-Path $dir -PathType Container) {
					log "Directory failed to be removed."
					$RemoveDirsSuccessful = $false
				}
				else {
					log "Directory removed."
				}
			}
			else {
				log "Directory doesn't exist."
			}
			log ""
		}
		log "Done removing directories."
		log ""
	}
	return $RemoveDirsSuccessful
}

# https://stackoverflow.com/questions/5648931/test-if-registry-value-exists
Function Test-RegistryValue {
    param(
        [Alias("PSPath")]
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Path
        ,
        [Parameter(Position = 1, Mandatory = $true)]
        [String]$Name
        ,
        [Switch]$PassThru
    ) 

    process {
        if (Test-Path $Path) {
            $Key = Get-Item -LiteralPath $Path
            if ($Key.GetValue($Name, $null) -ne $null) {
                if ($PassThru) {
                    Get-ItemProperty $Path $Name
                } else {
                    $true
                }
            } else {
                $false
            }
        } else {
            $false
        }
    }
}


# Remove optionally specified registry keys/values
function RemoveTheRegs {
	$RemoveRegsSuccessful = $true
	if($RemoveRegKeysValues -ne "false") {
		log "Removing registry keys/values:"
		log ""
		$regs = $RemoveRegKeysValues -split ','
		foreach($reg in $regs) {
			log $reg
			$key = $reg.split(';')[0]
			$value = $reg.split(';')[1]
			$keyIncludesValue = $false
			if($value -eq $null) {
				log "Key does not include value."
			}
			else {
				$keyIncludesValue = $true
				log "Key includes value:"
				log "Key: $key"
				log "Value: $value"
			}				
			
			if(Test-Path $key) {
				log "Key exists."
				if($keyIncludesValue) {
					if(Test-RegistryValue -Path $key -Name $value) {
						log "Value exists."
						log "Value is:"
						$valueDataString = Get-ItemProperty -Path $key -Name $value
						$valueData = $valueDataString.$value
						log $valueData
						# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/remove-itemproperty?view=powershell-6
						Remove-ItemProperty -Path $key -Name $value -Force -Verbose >> $ScriptLog 2>&1 4>&1
						if(Test-RegistryValue -Path $key -Name $value) {
							log "Value failed to be removed."
							$RemoveRegsSuccessful = $false
						
							log "Value exists."
							log "Value is:"
							$valueDataString = Get-ItemProperty -Path $key -Name $value
							$valueData = $valueDataString.$value
							log $valueData
						}
						else {
							log "Value removed."
						}
					}
					else {
						log "Value doesn't exist."
					}
				}
				else {
					# https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/working-with-registry-keys?view=powershell-6
					Remove-Item -Path $key -Recurse -Force -Verbose >> $ScriptLog 2>&1 4>&1
					if(Test-Path $key) {
						log "Key failed to be removed."
						$RemoveRegsSuccessful = $false
					}
					else {
						log "Key removed."
					}
				}
			}
			else {
				log "Key doesn't exist."
			}
			log ""
		}
		log "Done removing registry keys/values."
		log ""
	}
	return $RemoveRegsSuccessful
}

# Check for the existence of the app
$TargetApp = CheckForApp

# If not found, exit
if($TargetApp -eq "Not found") {
	log "App not found. Exiting with exit code $EXIT_AppNotFound."
	$EXIT_AppNotFound
	exit $EXIT_AppNotFound
}
else {
	# If found, pull product code (and default uninstall string because why not)
	# This may fail if more than one app matched the given query string
	$ProductKey = $TargetApp.PSChildName
	log "Product key (`"PSChildName`"):"
	log $ProductKey
	log ""

	$UninstallString = $TargetApp.UninstallString
	log "Uninstall string:"
	log $UninstallString
	log ""
	
	# Stop specified services
	if($StopServices -ne "false") {
		$StoppedServices = StopTheServices
		
		if($RequireServicesToStop -eq "true") {
			if($StoppedServices -eq $false) {
				log "-RequireServicesToStop was specified and not all specified services were stopped. Exiting with exit code $EXIT_ServicesFailed."
				$EXIT_ServicesFailed
				exit $EXIT_ServicesFailed
			}
		}
	}
	
	$UninstallApp = UninstallTheApp -ProductKey $ProductKey
	
	# Check for success
	$TargetApp = CheckForApp
	
	# If app is still found, then something failed
	# or perhaps the script is not waiting for msiexec to finish for some reason
	if($TargetApp -ne "Not found") {
		log "App not successfully uninstalled. Exiting with exit code $EXIT_AppUninstallFailed."
		$EXIT_AppUninstallFailed
		exit $EXIT_AppUninstallFailed
	}
	# If app is not found, then success
	else {
		log "App successfully uninstalled."
		log ""
		
		# Remove optionally specified directories
		$RemoveDirsSuccessful = RemoveTheDirs
		
		# Remove optionally specified registry keys/values
		$RemoveRegsSuccessful = RemoveTheRegs
		
		# Exit with relevant exit code
		if($RemoveDirsSuccessful) {
			if($RemoveRegsSuccessful) {
				log "App successfully uninstalled, and all directories and registry keys/values specified for removal successfully removed or didn't exist. Exiting with exit code $EXIT_Success."
				$EXIT_Success
				exit $EXIT_Success
			}
			else {
				log "App successfully uninstalled, all directories specified for removal successfully removed or didn't exist, but one or more registry keys/values specified for removal were not successfully removed. Exiting with exit code $EXIT_AppUninstalledDirsRemovedRegsFailed."
				$EXIT_AppUninstalledDirsRemovedRegsFailed
				exit $EXIT_AppUninstalledDirsRemovedRegsFailed
			}
		}
		else {
			if($RemoveRegsSuccessful) {
				log "App successfully uninstalled, all registry keys/values specified for removal successfully removed or didn't exist, but one or more directories specified for removal were not successfully removed. Exiting with exit code $EXIT_AppUninstalledDirsFailedRegsRemoved."
				$EXIT_AppUninstalledDirsFailedRegsRemoved
				exit $EXIT_AppUninstalledDirsFailedRegsRemoved
			}
			else {
				log "App successfully uninstalled, but one or more directories specified for removal were not successfully removed and one or more registry keys/values specified for removal were not successfully removed. Exiting with exit code $EXIT_AppUninstalledDirsFailedRegsFailed."
				$EXIT_AppUninstalledDirsFailedRegsFailed
				exit $EXIT_AppUninstalledDirsFailedRegsFailed
			}
		}
	}
}