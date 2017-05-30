<# 	Name: Install-Win-LIAgent.ps1
	Author: Jim McMahon II
	Date: 5/22/2017
	Synopsis: Quickly deploy Windows Log Insight agents
	Description: This script can be used to automate the installation of the Log Insight Windows agent without the
	use of Active Directory.
	
	Below example calls hostnames from a CSV and installs the agent using default values.
	Example: .\Install-Win-LIAgent.ps1 -Name (Get-Content "C:\Scripts\hosts.csv") -SourceInstall "C:\Software\LIInstall\liagent.msi" 
	-LIFQDN "log-01a.corp.local"
	
	Below example calls hostnames from a CSV and installs the agent with SSL set to false. Port is 9000.
	Example: .\Install-Win-LIAgent.ps1 -Name (Get-Content "C:\Scripts\hosts.csv") -SourceInstall "\\share\software\liinstall\liagent.msi"" 
	-LIFQDN "log-01a.corp.local" -LISSL $false	

    Logs: Logs are stored on the local C:\drive and the remote server. Install logs are stored on the remote server and
    a CSV file containing the job results are in .\liagent-install-results.csv. You can use the CSV to parse for errors
    using ReturnValue. If the value is not 0, there was an error with the install.
#>

[CmdletBinding()]
Param(
	[parameter(valuefrompipeline = $true, mandatory = $true,
		HelpMessage = "The FQDN or shortname of the computer. This parameter also accepts pipeline content.")]
		[PSObject]$ComputerName,
	[parameter(mandatory = $true,
		HelpMessage = "Enter the path to the Log Insight agent installation file. This can be a local or UNC path.")]
		[String]$SourceInstallFile,
	[parameter(mandatory = $true,
		HelpMessage = "This is the FQDN of the Log Insight appliance/cluster the agent will connect to.")]
		[string]$LIFQDN,
	#[string]$LIProtocol,
	#[string]$LIPort,
	[parameter(mandatory = $false,
		HelpMessage = "Enable or Disable SSL. Acceptable values are True or False.")]
		[Switch]$LISSL
	)
	
<#Example usage PS C:\scripts> .\test.ps1 -Name (Get-Content "C:\Scripts\hosts.csv") -SourceInstall "C:\scripts" -LIFQDN "log-01a.corp.l
ocal" -LISSL $true#>

# *************************************************************************
# Install the new Log Insight Agent and set its configuration.
function Install-LIagent {
    param(
    $ssl,
    $msiFile,
    $fqdn,
    $log)
    
    #$process = Start-Process -FilePath msiexec.exe -ArgumentList "/i $msiFile /q SERVERHOST=$fqdn /l*v $log" -Wait -Verbose -PassThru
    
    # Set the installation arguments list based on the decision to use SSL or not. The port will automatically be set
    # based on the SSL selection, i.e. SSL=No; Port=9000 SSL=Yes; Port=9543. CFAPI will be the set protocol.

    # SSL set to true, use the default values.
    if($ssl){
        $process = Start-Process -FilePath msiexec.exe -ArgumentList "/i $msiFile /q SERVERHOST=$fqdn /l*v $log" -Wait -Verbose -PassThru
        }
     
     #SSL set to false, set ssl=no.
     else{
        $process = Start-Process -FilePath msiexec.exe -ArgumentList "/i $msiFile /q SERVERHOST=$fqdn LIAGENT_SSL=no /l*v $log" -Wait -PassThru
        }
<#	if ($process.ExitCode -eq 0){
		Write-Host "$msiFile successfully installed on $comp" -ForegroundColor green
	}
	else {
		Write-Host "Could not install $msiFile on $comp. See the installer log file for more details" -ForegroundColor red
	}#>
}

# *************************************************************************
# Get any installed Log Insight agent on the server and uninstall it.
Function Uninstall-LIagent {
	#$app = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Log Insight Agent"}
	#$app.Uninstall()

    # Check if the agent is already installed. If so, uninstall it incase its old.
    if((Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Log Insight Agent"})) {
	    $app = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Log Insight Agent"}
	    $app.Uninstall()
	}
}

# *************************************************************************
# Get the status of running job(s)
function Get-JobStatus {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.Job[]]
        $Jobs
    )

    switch($jobs[0].Name){
        "Install Agent" {$jobType = "Installing agents..."}
        "Uninstall Agent" {$jobType = "Uninstalling agents..."}
    }

    $Count = 0
    Do {

        $TotProg = 0
        $total = $Jobs.Count

        ForEach ($Job in $Jobs){
            Try {
                $Prog = ($Job | Get-Job).ChildJobs[0].Progress.StatusDescription[-1]
                If ($Prog -is [char]){
                    $Prog = 0
                }
                
                $TotProg += $Prog
            }

            Catch {
                Start-Sleep -Milliseconds 500
                Break
            }
        }

        Write-Progress -Id 1 -Activity $jobType -Status "Waiting for background jobs to complete: $TotProg of $Total" -PercentComplete (($TotProg / $Total) * 100)
        Start-Sleep -Seconds 3
    } Until (($Jobs | Where State -eq "Running").Count -eq 0)
}

<# *************************************************************************
    Main script body
   *************************************************************************#>

# Set some additional paramaters for the script.
$localLogLoc  = "liagent_install.log" # Location of the installation log file on the remote server.
$locDir = "C:\LIInstall"              # Name of the temporary directory the msi is stored on the remote server.
$skipComputers = @()                  # Servers which failed during the file copy.
$goodComputers = @()                  # Successful file copy. We only start sessions with these.
$c1 = 0                               # Counter for the progress bar.

# Get the file name from the source path.
$a = dir $SourceInstallFile
$fileName = $a.Name

# Clean up jobs
Get-Job | Stop-Job
Get-Job | Remove-Job

# Copy install file locally.
$ComputerName | ForEach-Object {
    $dstFolder = "\\$_\C$\LIInstall"
    $c1++
    
    Write-Progress -Id 0 -Activity 'Copying Log Insight Agent' -CurrentOperation $_ -PercentComplete (($c1 / $ComputerName.count) * 100)    

    #This section will copy the $sourcefile to the $destinationfolder. If the Folder does not exist it will create it.
    if (!(Test-Path -path $dstFolder))
    {
        New-Item $dstFolder -Type Directory
    }
    Copy-Item -Path $SourceInstallFile -Destination $dstFolder -Verbose -Force
    
    # It is possible to install a single agent at a time if prefered by uncommenting the Invoke-Command below and commenting out the $s command
    # in the "Start the install..." section.
    if(Test-Path -path $dstFolder\$fileName){
        Write-Host "$fileName copied to $_ successfully. Proceeding with install." -ForegroundColor Green
        #Invoke-Command -Session $s -ScriptBlock ${function:Install-LIagent} -ArgumentList $LISSL,$_,$locDir\$fileName,$LIFQDN,$localLogLoc
        $goodComputers += $_
        }
    else {
        Write-Host "File copy to $_ was unsuccessful. This install will be skipped"
        $skipComputers += $_
    }
}
Write-Progress -Id 0 -Activity 'Copying Log Insight Agent' -PercentComplete (($c1 / $ComputerName.count) * 100) -Completed

# Start the install on all of the servers at once.
Write-Host "*******************************************************`nBeginning installation of agents. `n*******************************************************`n" -ForegroundColor Magenta
Write-Host "Uninstalling existing agents." -ForegroundColor Yellow

$s = New-PSSession -Computer $goodComputers
Invoke-Command -Session $s -ScriptBlock ${function:Uninstall-LIagent} -AsJob -JobName "Uninstall Agent"
Get-Job | Where-Object {$_.Name.Contains("Uninstall Agent")} | Get-JobStatus | Wait-Job

# Do not start the install until the uninstall job is complete.
Write-Host "*******************************************************`nInstalling new agents.`n*******************************************************`n" -ForegroundColor Yellow
Invoke-Command -Session $s -ScriptBlock ${function:Install-LIagent} -ArgumentList $LISSL,$locDir\$fileName,$LIFQDN,$localLogLoc -AsJob -JobName "Install Agent"
Get-Job | Where-Object {$_.Name.Contains("Install Agent")} | Get-JobStatus | Wait-Job

# Cleanup sessions.
# Remove the installation folder we copied.
$goodComputers | ForEach-Object {
    $dstFolder = "\\$_\C$\LIInstall"
    Remove-Item -Path $dstFolder -recurse
    }

# Remove all Powershell Sessions that were created.
Get-PSSession | Remove-PSSession

# Write results and dump the job status to a CSV.
Write-Host "Completed Log Insight Agent installations. Check the job for any child job errors using Get-Job -Id <#> | Select-Object -Property *"
Get-Job | Wait-Job | Out-Null
$result = Get-Job | Receive-Job | Select PSComputerName,RunspaceId,ReturnValue
$result | Export-Csv "C:\liagent-install-results.csv"

If($skipComputers){
    Write-Host "The following servers were skipped because the file could not be copied:" -ForegroundColor Yellow
    $skipComputers
    }
else{
    Write-Host "No servers were skipped!" -ForegroundColor Green
    }

<# *************************************************************************
    End script body
   *************************************************************************#>