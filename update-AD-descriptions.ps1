<#
This script updates description fields in Active Directory from data in a CSV file

How to use:

1. Make sure the CSV has the following columns:
    a. "Machine" (The machine name)
    b. "Location"
    c. "User"
    d. "Model"
    e. "OS"
2. In the terminal, enter: .\script_name.ps1 -csvPath "C:\path\to\the\data.csv"

Optional:
- Overwrite flag
    - Attach "-overwrite $true" to the terminal command if you wish to overwrite fields with an exisitng description.
    - The program by default ignores machines that already have a description

#>


param(
    [string]$csvPath = ".\ad_cleanup-workingdoc.csv",
    [string]$logFile = ".\AD_Update_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt",
    [bool]$overwrite = $false # overwrite machines that already have descriptions?
    #[string]$OU = "OU=SampleOU,OU=Computers,DC=example,DC=com"
)


$machines = import-csv -Path $csvPath # array of row objects

# Desired result:

# Location - User - Model - OS
# Jupiter - John Smith - Lenovo Thinkpad - Win 11


# add all computers
#foreach ($machine in $machines) {
#    New-ADComputer -Name $machine.Machine -Path "OU=SampleOU,OU=Computers,DC=example,DC=com"
#}

$successCount = 0
$failCount = 0
foreach ($machine in $machines) {
    # check if the machine exists in AD, if not, continue to next machine
    try {
	    $cur_machine = Get-ADComputer -Identity $machine.Machine -Properties Description
    } catch {
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | ERROR | $($machine.Machine) not found in AD"
        $logEntry | Add-Content -Path $logFile
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | ERROR | $($machine.Machine) not found in AD"
        $failCount++
        continue
    }
    
    # overwrite condition:
	# if the user does not want to overwrite and there is something already in the description, skip this machine
	if (-not [string]::IsNullOrWhiteSpace($cur_machine.Description) -and $overwrite -ne $true) { 
        continue 
    }

	$location = $machine.Location
	$user = $machine.User
	$model = $machine.Model
	$os = $machine.OS
	#$complete_flag = $machine.'Dont email/complete?'
	#$sent_flag = $machine.'Sent?'

	# if its a blank cell, make sure it will resolve to a default value, otherwise fill it in with info
    $location = if ($null -eq $location -or $location -eq "") { '' } else { "$location - " }
    $user = if ($null -eq $user -or $user -eq "") { '' } else { "$user - " }
    $model = if ($null -eq $model -or $model -eq "") { '' } else { "$model - " }
    $os = if ($null -eq $os -or $os -eq "") { '' } else { $os }

	$description = "$location"+"$user"+"$model"+"$os"
    # if we try to add nothing set to ' ' as setting it to '' freaks out the script
    if ([string]::IsNullOrWhiteSpace($description)) {
         $description = ' '
    }
    # if we're trying to update it with the same thing, no need to update, go to next machine
    if ($description -eq $cur_machine.Description) {
        continue
    }

	try {
		Set-ADComputer -Identity $machine.Machine -Description $description
		$logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | SUCCESS | $($machine.Machine) | Old: $($cur_machine.Description) | New: '$description'"
        $successCount++
	} catch {
		$logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | ERROR | $($machine.Machine) | Failed to update description: $_"
        $failCount++
	}

	$logEntry | Add-Content -Path $logFile
    Write-Host $logEntry 
}

Write-Host "Update complete. Successes: $successCount, Failures: $failCount"

#$machine.Machine | select -First 1