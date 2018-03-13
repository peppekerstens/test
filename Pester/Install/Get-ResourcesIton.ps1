#First some really basic checking before run
#Requires -Version 5
[CmdletBinding()]
param(
    [String]$PSGallery = "PSGallery",
    [String]$PSGalleryIton = "PSGalleryIton",
    [String]$ChocolateyIton = "ChocolateyIton"
)

#Script needs to run in elevated modus, self-elevate...
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

#Just call all functions as resources, if exists
Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Functions' ) -Recurse | Where-Object { $_.Name -notlike '_*' -and $_.Name -notlike '*.tests.ps1' -and $_.Name -like '*.ps1'} | ForEach-Object { . $_.FullName }

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

$Conditions = @(
    @{
        Label  = 'minimum OS level 2012R2'
        Test   = {
            #Detect OS version/type
            $OSVersion = [System.Environment]::OSVersion.Version
            [System.Version]"$($OSVersion.Major).$($OSVersion.Minor)" -ge [System.Version]'6.2'
        }
        Action = {
            Throw "OS does not meet minimum requirements. Quitting"
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
            Throw "No Internet connecting. Quitting"
        }
    },
    @{
        Label  = '64bit PSSession'
        Test   = {
            [Environment]::Is64BitProcess
        }
        Action = {
            Throw "No 64bit PSSession. Quitting"
        }
    },
    @{
        Label  = 'latest NuGet provider installed'
        Test   = {
            Try {
                $Latest = Find-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop
                $Installed = Get-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop
                $Installed.Version -ge $Latest.Version
            }
            Catch {
                $false
            }
        }
        Action = {
            Install-PackageProvider -Name 'NuGet' -Force #if not on system, just force install
        }
    },
    @{
        Label  = 'Latest PowerShellGet and PackageManagement installed'
        Test   = {
            (Update-Package -Name 'PowerShellGet' -Source $PSGallery -WhatIf).IsInstalled
        }
        Action = {
            #Update-Package -Name 'PowerShellGet' -Source $PublicSource
            Install-Module -Name 'PowerShellGet' -Force
        }
    },
    @{
        Label  = 'Latest Pester module installed'
        Test   = {
            (Update-Package -Name 'Pester' -Source $PSGallery -WhatIf).IsInstalled
        }
        Action = {
            #Update-Package -Name 'Pester' -Source $PublicSource
            Install-Module -Name 'Pester' -Force -SkipPublisherCheck
        }
    },
    @{
        Label  = 'latest Chocolatey provider installed'
        Test   = {
            Try {
                $Latest = Find-PackageProvider -Name 'chocolatey' -ForceBootstrap -ErrorAction Stop
                $Installed = Get-PackageProvider -Name 'chocolatey' -ForceBootstrap -ErrorAction Stop
                $Installed.Version -ge $Latest.Version
            }
            Catch {
                $false
            }
        }
        Action = {
            Install-PackageProvider -Name 'chocolatey' -Force #if not on system, just force install
        }
    },
    @{
        Label  = "Set $($PSGalleryIton) as trusted PSRepository"
        Test   = {
            (Get-PSRepository).Name -contains $PSGalleryIton
        }
        Action = {
            Register-PSRepository -Name $PSGalleryIton -SourceLocation 'https://psgallery.office-connect.nl/nuget' -PublishLocation 'https://psgallery.office-connect.nl/api/v2/package'
            Set-PackageSource -Name $PSGalleryIton -Trusted
        }
    },
    @{
        Label  = "Set $($ChocolateyIton) as trusted PSRepository"
        Test   = {
            (Get-PSRepository).Name -contains $ChocolateyIton
        }
        Action = {
            Register-PackageSource -Name $ChocolateyIton -Location 'https://chocolatey.office-connect.nl/nuget' -Provider chocolatey
            Set-PackageSource -Name $ChocolateyIton -Trusted
        }
    },
    @{
        Label  = 'GenericFunctions module'
        Test   = {
            (Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton  -WhatIf).IsInstalled
        }
        Action = {
            Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton -Force -AllVersions
        }
    },

    @{
        Label  = 'kvk00000000_Sonicwall-DPI-Certificate'
        Test   = {
            (Update-Package -Name 'kvk00000000_Sonicwall-DPI-Certificate' -Source $ChocolateyIton  -WhatIf).IsInstalled
        }
        Action = {
            Update-Package -Name 'kvk00000000_Sonicwall-DPI-Certificate' -Source $ChocolateyIton -Force -AllVersions
        }
    }
)

@($Conditions).foreach( {
        Write-Verbose "Testing condition [$($_.Label)]" -Verbose
        if (-not (& $_.Test)) {
            Write-Warning "Failed. Remediating..."
            & $_.Action
        }
        else {
            Write-Verbose 'Passed.' -Verbose
        }
    }) | Out-Null

If ($a) {
    $a.WarningBackgroundColor = $PreviousWarningBackgroundColor
    $a.WarningForegroundColor = $PreviousWarningForegroundColor
    $a.VerboseForegroundColor = $PreviousVerboseForegroundColor
}

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