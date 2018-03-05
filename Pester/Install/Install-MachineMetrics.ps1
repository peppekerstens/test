#Requires -Version 5

<#
.SYNOPSIS
Installeert of vernieuwd de PowerShell module MachineMetrics.

.DESCRIPTION
Installeert of vernieuwd de PowerShell module MachineMetrics inclusief alle voorwaarden om deze uit te kunnen voeren.
Tevens wordt een dagelijks schema gepland voor de uitvoer ervan.

Bij uitvoer wordt:
1. Gecontroleert of de module MachineMetrics aanwezig is.
2. Als deze aanwezig is en parameter force is niet gekozen dan wordt een update van de module gestart.
3. Als deze niet aanwezig is of als parameter force gekozen is dan;
a. Wordt juiste NuGet module geinstalleerd, als niet aanwezig
b. Wordt laatste versie van AzureRM module geinstalleerd en wordt getracht eerdere versie(s) te verwijderen
c. Wordt laatste versie van AzureRMStorageTable module geinstalleerd
d. Wordt het script 'Update-Package' geinstalleerd, wat zorgt voor de controle op en installatie van nieuwere versies van PowerShell modules en/of packages als het uitgevoerd wordt
e. Wordt de laatste versie van MachineMetrics gedownload via eerder script 'Update-Package'
f.

De functionaliteit werkt alleen op het Windows 2012 platform en hoger. Een server, gebaseerd op Windows 2008(R2) code,
wordt alleen voorzien van de packagemanegement module voor dat platform. Er vindt verder geen installatie van de PowerShell
 module MachineMetrics plaats

.PARAMETER Repository
Parameter description

.EXAMPLE
An example

.NOTES
Versie: 2.0
Datum: 1-3-2018
- Volledig herschreven
- De Guid van de klant
- Maakt gebruik van GenericFunctions module
- Console output alleen als -Verbose aangegeven is, op Warnings en Errors na
- Als parameter 'Log' niet opegegeven wordt, dan zal geen logging plaatsvinden
- Nieuwe parameter Eventlog. Als hier de naam van de eventlog opgegeven wordt, dan zal alle logging hierin plaatsvinden
- Nieuwe parameter EventID. Defaul


Copyright 2018 ITON Services B.V.
#>

[CmdletBinding()]
param(
    #KvK number of the customer.
    [Parameter(Mandatory = $true, ValueFromPipeline = $True)]
    [String]$KvK,
    #Path to log actions of this script to.
    [ValidateScript( {Test-Path $_ -IsValid})]
    [String]$Log,
    #Maximum size of the logfile in MB
    [Int]$LogSize = 1,
    [ValidateScript( {If (Get-Eventlog -LogName $_ -Newest 1) {$true} `
                Elsif(Write-EventLog -LogName $_ -Source $MyInvocation.MyCommand.Name -EventID 0 -EntryType Information) {$true} `
                Else {$false} `
        } )]
    $EventLogName,
    #Normal function of this script is: install or update when already installed. The force parameters forces a re-install/update of submodules.
    [Switch]$Force,
    [Switch]$MoreVerbose
)


#Script needs to run in elevated modus, self-elevate if possible...
#Source: https://blogs.msdn.microsoft.com/virtual_pc_guy/2010/09/23/a-self-elevating-powershell-script/
# Get the ID and security principal of the current user account
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

# Get the security principal for the Administrator role
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole)) {
    # We are running "as Administrator" - so change the title and background color to indicate this

    #$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
    #$Host.UI.RawUI.BackgroundColor = "DarkBlue"
    #clear-host
}
else {
    # We are not running "as Administrator" - so relaunch as administrator

    # Create a new process object that starts PowerShell
    $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";

    # Specify the current script path and name as a parameter
    $newProcess.Arguments = $myInvocation.MyCommand.Definition;

    # Indicate that the process should be elevated
    $newProcess.Verb = "runas";

    # Start the new process
    [System.Diagnostics.Process]::Start($newProcess);

    # Exit from the current, unelevated, process
    exit
}

#Start script here
$ScriptName = $MyInvocation.MyCommand.Name
$PSGallery = @{
    Name = 'PSGallery'
}
$PSGalleryITON = @{
    Name            = 'PSGalleryIton'
    SourceLocation  = 'https://psgallery.office-connect.nl/nuget'
    PublishLocation = 'https://psgallery.office-connect.nl/api/v2/package'
}
#Set write-warning to better stand-out from verbose and debug info.
$a = (Get-Host).PrivateData
If ($a) {
    #Not every PS host has this capability
    $PreviousWarningBackgroundColor = $a.WarningBackgroundColor
    $PreviousWarningForegroundColor = $a.WarningForegroundColor
    $PreviousVerboseForegroundColor = $a.VerboseForegroundColor
    $a.WarningBackgroundColor = "red"
    $a.WarningForegroundColor = "white"
    $a.VerboseForegroundColor = 'cyan'
}

#Create or shrink logfile to desired size
If ($Log) {
    If (!(Test-Path -Path (Split-Path -Path $Log -Parent -Verbose:$MoreVerbose) -Verbose:$MoreVerbose)) {New-Item -Path (Split-Path -Path $Log -Parent -Verbose:$MoreVerbose) -ItemType Directory -Verbose:$MoreVerbose | Out-Null}
    If ((Test-Path -Path $Log -Verbose:$MoreVerbose) -and ($LogSize)) {
        While ((Get-Item $Log -Verbose:$MoreVerbose).length -gt 1mb / $LogSize) {
            #Bron: https://stackoverflow.com/questions/2074271/remove-top-line-of-text-file-with-powershell
            ${$Log} = ${$Log} | Select-Object -Skip 10 -Verbose:$MoreVerbose
        }
    }
}

#Phase 1: Install minimum needed functionality for packagemanagement and ITON specific functions
If (!(Get-Package -Name 'GenericFunctions' -Verbose:$MoreVerbose -ErrorAction 'SilentlyContinue') -or $Force) {
    . $PSScriptRoot\Invoke-PackageBootstrap.ps1 
    $Splatparam = {}
    ($log) -and ($Splatparam.Log = $Log)
    ($LogSize) -and ($Splatparam.LogSize = $LogSize)
    ($EventLogName) -and ($Splatparam.EventLogName = $EventLogName)
    ($MoreVerbose) -and ($Splatparam.MoreVerbose = $MoreVerbose)
    Invoke-PackageBootstrap @Splatparam
}

#Phase 2: Start testing the machine/environment for the Desired State

#Start all pester based tests
Invoke-Pester -OutputFile $PSScriptRoot\PesterResult.xml -Verbose:$MoreVerbose

#Check the test results
#$PesterTests = Get-ChildItem -Path $PSScriptRoot -Recurse | Where-Object Name -like "*.Tests.ps1"
#Foreach ($Test in $PesterTests) {
#[xml]$PesterResult = Get-Content -Path ($Test -replace '.ps1', '.xml')
[xml]$PesterResult = Get-Content -Path $PSScriptRoot\PesterResult.xml -Verbose:$MoreVerbose
$PesterExceptions = $PesterResult.DocumentElement.faillures + $PesterResult.DocumentElement.inconclusive + $PesterResult.DocumentElement.skipped + $PesterResult.DocumentElement.invalid
If ($PesterExceptions -ne 0) {
    Write-Warning -Message "Some pre-installation tests have failed. Do you want to remediate the system/environment? This may imply changes in behavior and reboots."
    $Title = [String]::Empty
    $Info = "Do you want to continue?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continues with currect selection" -Verbose:$MoreVerbose
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Stops currect selection" -Verbose:$MoreVerbose
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    [int]$defaultchoice = 0
    $opt = $host.UI.PromptForChoice($Title , $Info , $Options, $defaultchoice)
    switch ($opt) {
        1 { exit }
        Default { }
    }

    #There are errors, iterate through seperate tests
    Foreach ($Test in ($PesterResult.DocumentElement.ChildNodes.results.ChildNodes.results.ChildNodes.results.ChildNodes.results.ChildNodes.Where{$_.result -ne 'Succes'})) {

    }
}
#}
#>
($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($ScriptName) -EventID $EventID -EntryType Information -Message "MyApp added a user-requested feature to the display.") | Out-Null

If (Get-Package -Name 'GenericFunctions' -Verbose:$MoreVerbose -ErrorAction 'SilentlyContinue') {
    $Conditions = @(
        @{
            Label  = 'latest AzureRM module installed'
            Test   = {
                (Update-Package -Name 'AzureRM' -Source $PSGallery.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled
            }
            Action = {
                #AzureRm is a 'special' module which needs other update method
                Update-AzureRM -Source $PSGallery.Name -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'latest AzureRmStorageTable installed'
            Test   = {
                (Update-Package -Name 'AzureRmStorageTable' -Source $PSGallery.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled
            }
            Action = {
                Update-Package -Name 'AzureRmStorageTable' -Source $PSGallery.Name -AllVersions -Force -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'latest MachineMetrics installed'
            Test   = {
                (Update-Package -Name 'MachineMetrics' -Source $PSGalleryIton.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled
            }
            Action = {
                Update-Package -Name 'MachineMetrics' -Source $PSGalleryIton.Name -AllVersions -Force -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'set "Update MachineMetrics" ScheduledTask'
            Test   = {
                (Get-ScheduledTask -TaskName "Update MachineMetrics" -ErrorAction 'SilentlyContinue' -Verbose:$MoreVerbose) -and $true
            }
            Action = {
                $SplatSettings = @{
                    TaskName    = "Update MachineMetrics"
                    Description = "Updates MachineMetrics module, if there is any"
                    Execute     = "C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe"
                    Argument    = "-NonInteractive -NoProfile -ExecutionPolicy Unrestricted -Command Import-Module GenericFunctions;Update-Package -Name MachineMetrics -Force"
                    User        = "LOCALSYSTEM"
                }
                Register-Schedule @SplatSettings -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'set "MachineMetrics" ScheduledTask'
            Test   = {
                (Get-ScheduledTask -TaskName "MachineMetrics" -ErrorAction 'SilentlyContinue' -Verbose:$MoreVerbose) -and $true
            }
            Action = {
                $SplatSettings = @{
                    TaskName    = "MachineMetrics"
                    Description = "ITON daily machine metrics"
                    Execute     = "C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe"
                    Argument    = "-NonInteractive -NoProfile -ExecutionPolicy Unrestricted -Command Import-Module MachineMetrics;Publish-Metrics"
                    Type        = 'Daily'
                    At          = '00:00:01'
                    RandomDelay = '00:05:00'
                    User        = "LOCALSYSTEM"
                }
                Register-Schedule @SplatSettings -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'set customer info'
            Test   = {
                [String]$Key = 'HKLM:\SYSTEM\ITON'
                (Get-ItemProperty -Path $Key -Name 'KvK' -ErrorAction 'SilentlyContinue' -Verbose:$MoreVerbose) -and $true
            }
            Action = {
                [String]$Key = 'HKLM:\SYSTEM\ITON'
                If (!(Test-Path $Key)) {
                    New-Item -Path $Key -Force -Verbose:$MoreVerbose | Out-Null
                }
                Try {
                    Get-ItemProperty -Path $Key -Name 'KvK' -ErrorAction Stop -Verbose:$MoreVerbose
                    Set-ItemProperty -Path $Key -Name 'KvK' -Value $KvK -Force -Verbose:$MoreVerbose | Out-Null
                }
                Catch {
                    New-ItemProperty -Path $Key -Name 'KvK' -Value $KvK `
                        -PropertyType STRING -Force -Verbose:$MoreVerbose | Out-Null
                }
            }
        }
    )

    Import-Module GenericFunctions -Verbose:$MoreVerbose
    #Check if another process with this scriptname is running, and kill it
    Stop-PSCommandLineProcess -CommandLine $ScriptName -Verbose:$MoreVerbose
    $WriteLogPresent = (Get-Command Write-Log -ErrorAction SilentlyContinue -Verbose:$MoreVerbose) -and $true

    Write-Verbose "Installing $ScriptName functionality"
    ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Installing $ScriptName functionality" -Path $Log -Verbose:$MoreVerbose) | Out-Null

    @($Conditions).foreach( {
            Write-Verbose "Testing condition [$($_.Label)]"
            ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Testing condition [$($_.Label)]" -Path $Log -Verbose:$MoreVerbose)| Out-Null
            if (-not (& $_.Test)) {
                Write-Warning "Condition [$($_.Label)] failed. Remediating..."
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Condition [$($_.Label)] failed. Remediating..." -Path $Log -Level 'Warning' -Verbose:$MoreVerbose)| Out-Null
                & $_.Action
            }
            else {
                Write-Verbose 'Passed.'
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Condition [$($_.Label)] passed" -Path $Log -Verbose:$MoreVerbose)| Out-Null
            }
        }) | Out-Null

    #Set some final settings and run first Publish-Metrics
    Disable-AzureRMDataCollection
    Set-Diskmap
    Publish-Metrics    
}


If ($a) {
    $a.WarningBackgroundColor = $PreviousWarningBackgroundColor
    $a.WarningForegroundColor = $PreviousWarningForegroundColor
    $a.VerboseForegroundColor = $PreviousVerboseForegroundColor
}

#Write-Host -NoNewLine "Press any key to continue..."
#$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")