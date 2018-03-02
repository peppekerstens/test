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
    [Switch]$Force
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


Function Update-Package {
    [CmdletBinding(
        SupportsShouldProcess = $true,
        ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [String[]]$Name,
        [String]$Source,
        [Switch]$AllVersions,
        [Switch]$Force
    )

    Begin {}

    Process {
        Foreach ($N in $Name) {
            $ReturnInfo = [PSCustomObject]@{
                Name            = $N
                IsInstalled     = $false
                Version         = [System.Version]'0.0'
                Latest          = [System.Version]'0.0'
                UpdateNeeded    = $false
                UpdateSucceeded = $false
                Source          = $Source
                Success         = $false
            }
            If (Get-PackageProvider -Verbose:$false| Where-Object -Property Name -eq NuGet) {
                #Check if source is available otherwise stop
                $SplatParam = @{
                    Name = $N
                }
                ($Source) -and ($SplatParam.Source = $Source) | out-null
                Try {
                    $Latest = Find-Package @SplatParam -ErrorAction Stop -Verbose:$false
                    $ReturnInfo.Latest = $Latest.Version
                }
                Catch {
                    #Just leave latest as it is..
                }
                Try {
                    $SplatParam = @{
                        Name = $N
                    }
                    $Installed = Get-Package @SplatParam -ErrorAction Stop -Verbose:$false
                    $ReturnInfo.IsInstalled = $true
                    $ReturnInfo.Version = $Installed.Version
                    If ($Installed.Version -ge $Latest.Version) {$Update = $false}
                }
                Catch {
                    $Update = $true
                }
                $ReturnInfo.UpdateNeeded = $Update
                $ReturnInfo.UpdateSucceeded = $false
                If ($Update) {
                    ($Source) -and ($SplatParam.Source = $Source) | out-null
                    ($Force) -and ($SplatParam.Force = $Force) | out-null
                    Switch ($Force) {
                        $false {$Answer = $PSCmdlet.ShouldProcess($Name)}
                        $true {$Answer = $true}
                    }
                    If ($Answer) {
                        Try {
                            $SplatAllVersions = @{}
                            ($AllVersions) -and ($SplatAllVersions.AllVersions = $AllVersions) | out-null
                            Uninstall-Package @SplatParam @SplatAllVersions -ErrorAction Stop -Verbose:$false
                            $Info.Source = (Install-Package @SplatParam -ErrorAction Stop -Verbose:$false).Source
                            $Info.UpdateSucceeded = $true
                        }
                        Catch {
                            Try {
                                #Force installation
                                $Info.Source = (Install-Package @SplatParam -ErrorAction Stop -Verbose:$false).Source
                                $Info.UpdateSucceeded = $true
                            }
                            Catch {
                            }
                        }
                    }
                }
                $ReturnInfo.Success = $true
                #Just post current status back...
                $ReturnInfo
            }
        }
    }

    End {}
}

#Start script here
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
    If (!(Test-Path $Log)) {New-Item -Path $Log -ItemType File}
    While ((Get-Item $Log).length -gt $LogSize / 1mb) {
        #Bron: https://stackoverflow.com/questions/2074271/remove-top-line-of-text-file-with-powershell
        ${$Log} = ${$Log} | Select-Object -Skip 10
    }
}

#Phase 1: Minimum pre-requisites. Prepare machine for PackageMangement, Pester testing and GenericFunctions
$GenericFunctionsPresent = Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -WhatIf
If (!($GenericFunctionsPresent.Success -and $GenericFunctionsPresent.IsInstalled)) {
    #First, setup/test basic pre-requisites for installation
    $prereqConditions = @(
        @{
            Label  = 'minimum OS level 2012R2'
            Test   = {
                #Detect OS version/type
                $OSVersion = [System.Environment]::OSVersion.Version
                [System.Version]"$($OSVersion.Major).$($OSVersion.Minor)" -ge [System.Version]'6.2'
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "OS does not meet minimum requirements." -Path $Log -Level 'Error')| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($MyInvocation.MyCommand.Name) -EventID $EventID -EntryType Error -Message "OS does not meet minimum requirements")
                Throw "OS does not meet minimum requirements!"
            }
        },
        @{
            Label  = 'Internet connection'
            Test   = {
                Try {
                    Test-Connection -ComputerName 8.8.8.8 -Count 1 -ErrorAction Stop
                    $true
                }
                Catch {
                    $false
                }
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "No Internet connection." -Path $Log -Level 'Error')| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($MyInvocation.MyCommand.Name) -EventID $EventID -EntryType Error -Message "No Internet connection")
                Throw "No Internet connection!"
            }
        },
        @{
            Label  = '64bit PSSession'
            Test   = {
                [Environment]::Is64BitProcess
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "No 64bit PSSession." -Path $Log -Level 'Error')| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($MyInvocation.MyCommand.Name) -EventID $EventID -EntryType Error -Message "No 64bit PSSession")
                Throw "No 64bit PSSession!"
            }
        },
        @{
            Label  = 'latest NuGet provider installed'
            Test   = {
                Try {
                    $Latest = Find-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop -Verbose:$false
                    $Installed = Get-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop -Verbose:$false
                    $Installed.Version -ge $Latest.Version
                }
                Catch {
                    $false
                }
            }
            Action = {
                Install-PackageProvider -Name 'NuGet' -Force -Verbose:$false #if not on system, just force install
            }
        },
        @{
            Label  = 'latest PowerShellGet and PackageManagement installed'
            Test   = {
                (Update-Package -Name 'PowerShellGet' -Source $PSGallery.Name -WhatIf).IsInstalled
            }
            Action = {
                Update-Package -Name 'PowerShellGet' -Source $PSGallery.Name -AllVersions -Force
            }
        },
        @{
            Label  = 'latest Pester module installed'
            Test   = {
                (Update-Package -Name 'Pester' -Source $PSGallery.Name -WhatIf).IsInstalled
            }
            Action = {
                Install-Package -Name 'Pester' -Source $PSGallery.Name -Force -SkipPublisherCheck -Verbose:$false
            }
        }
        @{
            Label  = 'Set PSGalleryITON as trusted PSRepository'
            Test   = {
                (Get-PSRepository -Verbose:$false).Name -contains $PSGalleryIton.Name
            }
            Action = {
                Register-PSRepository -Name $PSGalleryIton.Name -SourceLocation $PSGalleryIton.SourceLocation -PublishLocation $PSGalleryIton.PublishLocation -Verbose:$false
                Set-PackageSource -Name $PSGalleryIton.Name -Trusted -Verbose:$false
            }
        },
        @{
            Label  = 'latest GenericFunctions module installed'
            Test   = {
                (Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -WhatIf).IsInstalled
            }
            Action = {
                Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -AllVersions -Force
            }
        }
    )

    $WriteLogPresent = (Get-Command Write-Log -ErrorAction SilentlyContinue) -and $true

    Write-Verbose "Preparing minimum requirements for installation"
    ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Preparing minimum requirements for installation" -Path $Log) | Out-Null

    @($prereqConditions).foreach( {
            Write-Verbose "Testing condition [$($_.Label)]"
            ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Testing condition [$($_.Label)]" -Path $Log -Level 'Info')| Out-Null
            if (-not (& $_.Test)) {
                Write-Warning "Condition [$($_.Label)] failed. Remediating..."
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Condition [$($_.Label)] failed. Remediating..." -Path $Log -Level 'Warn')| Out-Null
                & $_.Action
            }
            else {
                Write-Verbose 'Passed.'
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Condition [$($_.Label)] passed" -Path $Log -Level 'Info')| Out-Null
            }
        }) | Out-Null
}


#Phase 2: Start testing the machine/environment for the Desired State

#Start all pester based tests
Invoke-Pester -OutputFile $PSScriptRoot\PesterResult.xml -Verbose:$false

#Check the test results
#$PesterTests = Get-ChildItem -Path $PSScriptRoot -Recurse | Where-Object Name -like "*.Tests.ps1"
#Foreach ($Test in $PesterTests) {
#[xml]$PesterResult = Get-Content -Path ($Test -replace '.ps1', '.xml')
[xml]$PesterResult = Get-Content -Path $PSScriptRoot\PesterResult.xml
$PesterExceptions = $PesterResult.DocumentElement.faillures + $PesterResult.DocumentElement.inconclusive + $PesterResult.DocumentElement.skipped + $PesterResult.DocumentElement.invalid
If ($PesterExceptions -ne 0) {
    Write-Warning -Message "Some pre-installation tests have failed. Do you want to remediate the system/environment? This may imply changes in behavior and reboots."
    $Title = [String]::Empty
    $Info = "Do you want to continue?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continues with currect selection"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Stops currect selection"
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
($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($MyInvocation.MyCommand.Name) -EventID $EventID -EntryType Information -Message "MyApp added a user-requested feature to the display.")

If ($GenericFunctionsPresent.Success -and $GenericFunctionsPresent.IsInstalled) {
    $Conditions = @(
        @{
            Label  = 'latest AzureRM module installed'
            Test   = {
                (Update-Package -Name 'AzureRM' -Source $PSGallery.Name -WhatIf).IsInstalled
            }
            Action = {
                #AzureRm is a 'special' module which needs other update method
                Update-AzureRM -Source $PSGallery.Name
            }
        },
        @{
            Label  = 'latest AzureRmStorageTable installed'
            Test   = {
                (Update-Package -Name 'AzureRmStorageTable' -Source $PSGallery.Name -WhatIf).IsInstalled
            }
            Action = {
                Update-Package -Name 'AzureRmStorageTable' -Source $PSGallery.Name -AllVersions -Force
            }
        },
        @{
            Label  = '64bit PSSession'
            Test   = {
                [Environment]::Is64BitProcess
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "No 64bit PSSession." -Path $Log -Level 'Error')| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($MyInvocation.MyCommand.Name) -EventID $EventID -EntryType Error -Message "No 64bit PSSession")
                Throw "No 64bit PSSession!"
            }
        },
        @{
            Label  = 'latest NuGet provider installed'
            Test   = {
                Try {
                    $Latest = Find-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop -Verbose:$false
                    $Installed = Get-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop -Verbose:$false
                    $Installed.Version -ge $Latest.Version
                }
                Catch {
                    $false
                }
            }
            Action = {
                Install-PackageProvider -Name 'NuGet' -Force -Verbose:$false #if not on system, just force install
            }
        },
        @{
            Label  = 'latest PowerShellGet and PackageManagement installed'
            Test   = {
                (Update-Package -Name 'PowerShellGet' -Source $PSGallery.Name -WhatIf).IsInstalled
            }
            Action = {
                Update-Package -Name 'PowerShellGet' -Source $PSGallery.Name -AllVersions -Force
            }
        },
        @{
            Label  = 'latest Pester module installed'
            Test   = {
                (Update-Package -Name 'Pester' -Source $PSGallery.Name -WhatIf).IsInstalled
            }
            Action = {
                Install-Package -Name 'Pester' -Source $PSGallery.Name -Force -SkipPublisherCheck -Verbose:$false
            }
        }
        @{
            Label  = 'Set PSGalleryITON as trusted PSRepository'
            Test   = {
                (Get-PSRepository -Verbose:$false).Name -contains $PSGalleryIton.Name
            }
            Action = {
                Register-PSRepository -Name $PSGalleryIton.Name -SourceLocation $PSGalleryIton.SourceLocation -PublishLocation $PSGalleryIton.PublishLocation -Verbose:$false
                Set-PackageSource -Name $PSGalleryIton.Name -Trusted -Verbose:$false
            }
        },
        @{
            Label  = 'latest GenericFunctions module installed'
            Test   = {
                (Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -WhatIf).IsInstalled
            }
            Action = {
                Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -AllVersions -Force
            }
        }
    )

    Import-Module GenericFunctions
    $WriteLogPresent = (Get-Command Write-Log -ErrorAction SilentlyContinue) -and $true

    Write-Verbose "Preparing minimum requirements for installation"
    ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Preparing minimum requirements for installation" -Path $Log) | Out-Null

    @($prereqConditions).foreach( {
            Write-Verbose "Testing condition [$($_.Label)]"
            ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Testing condition [$($_.Label)]" -Path $Log -Level 'Info')| Out-Null
            if (-not (& $_.Test)) {
                Write-Warning "Condition [$($_.Label)] failed. Remediating..."
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Condition [$($_.Label)] failed. Remediating..." -Path $Log -Level 'Warn')| Out-Null
                & $_.Action
            }
            else {
                Write-Verbose 'Passed.'
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Condition [$($_.Label)] passed" -Path $Log -Level 'Info')| Out-Null
            }
        }) | Out-Null

}


If ($a) {
    $a.WarningBackgroundColor = $PreviousWarningBackgroundColor
    $a.WarningForegroundColor = $PreviousWarningForegroundColor
    $a.VerboseForegroundColor = $PreviousVerboseForegroundColor
}

Write-Host -NoNewLine "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")