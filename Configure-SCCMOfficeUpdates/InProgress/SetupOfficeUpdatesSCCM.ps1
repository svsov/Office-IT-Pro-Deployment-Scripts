<#

#>

[CmdletBinding(SupportsShouldProcess=$true)]
Param
(
	[Parameter(Mandatory=$True)]
	[String]$version,

	[Parameter(Mandatory=$True)]
	[String]$path,

	[Parameter()]
	[String]$bitness = '64',

	[Parameter(Mandatory=$True)]
	[String]$siteId,

	<#[Parameter(Mandatory=$True)]
	[PSCredential]$credential,	#>
	
	[Parameter()]
	[String]$UpdateSourceConfigFileName = 'Configuration_UpdateSource.xml',

	[Parameter()]
	[String]$UpdateTestGroupConfigFileName = 'Configuration_UpdateTestGroup.xml'
)
Begin
{
    #if(-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    #{
        #$arguments = "& '" + $myinvocation.mycommand.definition + "'"; 
        #Start-Process $($PSHOME)\powershell.exe -Verb runAs -ArgumentList $arguments; Break 
    #}

    $currentExecutionPolicy = Get-ExecutionPolicy
	Set-ExecutionPolicy Unrestricted -Scope Process -Force  
    $startLocation = Get-Location
}
Process
{
	# Get Credentials required to connect to the Share Path
	#$secpasswd = ConvertTo-SecureString $plainTextPassword -AsPlainText -Force

	#$credentials = Get-Credential #New-Object System.Management.Automation.PSCredential ($username, $secpasswd) # Alternative - $credentials = Get-Credential
    
    Write-Host 'Updating Config Files'

	$c2rFileName = 'setup.exe'

	#1 - Set the correct version number to update Source location
	$sourceFilePath = "$path\$UpdateSourceConfigFileName"
    $localSourceFilePath = ".\$UpdateSourceConfigFileName"
	$sourceContent = [Xml] (Get-Content $localSourceFilePath)
	$addNode = $sourceContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
	$addNode.Version = $version

	$sourceContent.Save($sourceFilePath)

	$testGroupFilePath = "$path\$UpdateTestGroupConfigFileName"
    $localtestGroupFilePath = ".\$UpdateTestGroupConfigFileName"
	$testGroupConfigContent = [Xml] (Get-Content $localtestGroupFilePath)
	$addNode = $testGroupConfigContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
    $addNode.SourcePath = $path	

	$updatesNode = $testGroupConfigContent.Configuration.Updates
	$updatesNode.UpdatePath = $path
	$updatesNode.TargetVersion = $version

	$testGroupConfigContent.Save($testGroupFilePath)
    
    Write-Host 'Setting up Click2Run to download specified version'

	$setupExePath = "$path\$c2rFileName"

	#2 - Run Setup.exe to download bits for specified version

	#Connect PowerShell to Share location	
	Set-Location $path
	# set up the executable with appropriate arguments
	$app = ".\$c2rFileName" 
	$arguments = "/download", "$UpdateSourceConfigFileName"
    
    Write-Host 'Download Start'

	#run the executable, this will trigger the download of bits to \\ShareName\Office\Data\
	& $app @arguments

    Write-Host 'Download Complete'

	Set-Location $startLocation

    Set-Location $PSScriptRoot

    Write-Host 'Loading SCCM Module'

    Import-Module "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"

    # The SCCM Script runs here
	. .\SetupOfficeUpdates.ps1

	#connect to sccm PowerShell
	Set-Location "$siteId`:"	
	
    Write-Host 'Starting SCCM Script'

	SetupOfficeUpdates -path '\\officeautosc-dc\shares\Office' -collectionToUse 'TestCollection' -distributionPointGroupName 'TestDPGroup' -configFileName $UpdateTestGroupConfigFileName 
}
End
{
    Set-ExecutionPolicy $currentExecutionPolicy -Scope Process -Force
    Set-Location $startLocation    
}







