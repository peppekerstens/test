#First some really basic checking before run
#Requires -Version 5
[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High')]
param(
    [String]$PSGallery = "psgallery"
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
            #Check if source is available otherwise stop
            $ReturnInfo = [PSCustomObject]@{
                Name            = $N
                IsInstalled     = $false
                Version         = 0
                Latest          = 0
                UpdateNeeded    = $false
                UpdateSucceeded = $false
                Source          = $Source
            }
            $SplatParam = @{
                Name = $N
            }
            ($Source) -and ($SplatParam.Source = $Source) | out-null
            $Latest = Find-Package @SplatParam -ErrorAction Stop
            $ReturnInfo.Latest = $Latest.Version
            Try {
                $SplatParam = @{
                    Name = $N
                }
                $Installed = Get-Package @SplatParam -ErrorAction Stop
                $ReturnInfo.IsInstalled = $true
                $ReturnInfo.Version = $Installed.Version
                If ($Installed.Version -ge $Latest.Version) {$Update = $false}
            }
            Catch {
                $Update = $true
            }
            <#
            #Sometimes, the package registration fails but lingering older versions prevent right installation/registration
            #Below code deletes on name, but is risky. Disabled for now; needs better solution.
            $PSModulePath = $env:PSModulePath -split ';'
            Foreach ($PSModPath in $PSModulePath) {
                If (Test-Path -Path $PSModPath) {
                    If ((Get-ChildItem -Path $PSModPath).Name -contains $N) {
                        Remove-Item -Path (Join-Path -Path $PSModPath -ChildPath $N) -Force -Recurse -Confirm:$false
                    }
                }
            }
            $CurrentPackage = @{
                Version = '0'
            }
            #>
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
                        Uninstall-Package @SplatParam @SplatAllVersions -ErrorAction Stop
                        $Info.Source = (Install-Package @SplatParam -ErrorAction Stop).Source
                        $Info.UpdateSucceeded = $true
                    }
                    Catch {
                        Try {
                            #Force installation
                            $Info.Source = (Install-Package @SplatParam  -ErrorAction Stop).Source
                            $Info.UpdateSucceeded = $true
                        }
                        Catch {
                        }
                    }
                }
            }
            #Just post current status back...
            $ReturnInfo
        }
    }

    End {}
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
        Label  = 'latest PowerShellGet and PackageManagement installed'
        Test   = {
            (Update-Package -Name 'PowerShellGet' -Source $PSGallery -WhatIf).IsInstalled
        }
        Action = {
            #Update-Package -Name 'PowerShellGet' -Source $PublicSource
            Install-Module -Name 'PowerShellGet' -Force
        }
    },
    @{
        Label  = 'latest Pester module installed'
        Test   = {
            (Update-Package -Name 'Pester' -Source $PSGallery -WhatIf).IsInstalled
        }
        Action = {
            Install-Module -Name 'Pester' -Force -SkipPublisherCheck
        }
    }
    <#
    @{
        Label  = 'Set PSGalleryITON as trusted PSRepository'
        Test   = {
            (Get-PSRepository).Name -contains $ITONSource
        }
        Action = {
            Register-PSRepository -Name $ITONSource -SourceLocation 'https://psgallery.office-connect.nl/nuget' -PublishLocation 'https://psgallery.office-connect.nl/api/v2/package' -Verbose
            Set-PackageSource -Name $ITONSource -Trusted -Verbose
        }
    }

    @{
        Label  = 'Latest AzureRM module installed'
        Test   = {
            Test-UpdatePackage -Name 'AzureRM' -Source $PublicSource
        }
        Action = {
            Update-AzureRM
        }
    },
    @{
        Label  = 'Latest AzureStorageTable module installed'
        Test   = {
            Test-UpdatePackage -Name 'AzureRmStorageTable' -Source $PublicSource
        }
        Action = {
            Update-Package -Name 'AzureRmStorageTable' -Source $PublicSource
            #Install-AzureRmStorageTable
        }
    },

    @{
        Label  = 'Get latest IaaSMetrics module'
        Test   = {
            Test-UpdatePackage -Name 'IaaSMetrics' -Source $PublicSource
        }
        Action = {
            Update-Package -Name 'IaaSMetrics' -Source $ITONSource
        }
    },
    @{
        Label  = 'SQL 2012 Express LocalDB edition installed'
        Test   = {
            Try {
                Get-Package -Name "Microsoft SQL Server 2012 Express LocalDB*" -ProviderName 'msi' -ErrorAction 'Stop'
                $true
            }
            Catch {
                $false
            }
        }
        Action = {
            #source: https://stackoverflow.com/questions/42951632/sql-server-express-localdb-msi-offline-installer
            #2012 download site: https://www.microsoft.com/en-us/download/details.aspx?id=29062
            #2017 version direct link: https://download.microsoft.com/download/E/F/2/EF23C21D-7860-4F05-88CE-39AA114B014B/SqlLocalDB.msi
            #2017 installed MSI naam: "Microsoft SQL Server 2017 localDB*"
            #2016 version direct link: https://download.microsoft.com/download/9/0/7/907AD35F-9F9C-43A5-9789-52470555DB90/ENU/SqlLocalDB.msi
            #2016 version direct link: https://download.microsoft.com/download/8/D/D/8DD7BDBA-CEF7-4D8E-8C16-D9F69527F909/ENU/x64/SqlLocalDB.MSI
            Get-UrlFile -Url 'https://download.microsoft.com/download/8/D/D/8DD7BDBA-CEF7-4D8E-8C16-D9F69527F909/ENU/x64/SqlLocalDB.MSI' -Output "$($env:TEMP)\SqlLocalDB.msi"
            $SplatParams = @{
                Executable = "msiexec.exe"
                Arguments  = "/i `"$($env:TEMP)\SqlLocalDB.msi`" IACCEPTSQLLOCALDBLICENSETERMS=YES /qn"
            }
            Start-Executable @SplatParams
        }
    },
    @{
        Label  = 'Azure Storage Emulator installed'
        Test   = {
            Try {
                Get-Package -Name "Microsoft Azure Storage Emu*" -ProviderName Programs -ErrorAction 'Stop'
                $true
            }
            Catch {
                $false
            }
        }
        Action = {
            <#
            After installing both SQLLocalDB and this, normal init as per https://docs.microsoft.com/nl-nl/azure/storage/common/storage-use-emulator fails with message:
            "...
            Looking for a LocalDB Installation.
            Probing SQL Instance: '(localdb)\MSSQLLocalDB'.
            Caught exception while probing for SQL endpoint. A network-related or instance-specific error occurred
            ..."
            A fix for this issue is this: https://stackoverflow.com/questions/12103410/issue-with-azure-emulator-and-sqllocaldb
            #>
    <#
            Get-UrlFile -Url 'https://go.microsoft.com/fwlink/?linkid=717179&clcid=0x409' -Output "$($env:TEMP)\MicrosoftAzureStorageEmulator.msi"
            $SplatParams = @{
                Executable = "msiexec.exe"
                Arguments  = "/i `"$($env:TEMP)\MicrosoftAzureStorageEmulator.msi`" /qn"
            }
            Start-Executable @SplatParams
        }
    }
    #>
)

Write-Verbose "Preparing minimum requirements for automated installation"

@($prereqConditions).foreach( {
        Write-Verbose "Testing condition [$($_.Label)]"
        if (-not (& $_.Test)) {
            Write-Warning "Condition [$($_.Label)] failed. Remediating..."
            & $_.Action
        }
        else {
            Write-Verbose 'Passed.'
        }
    }) | Out-Null



#Start all pester based tests
Invoke-Pester -OutputFile $PSScriptRoot\PesterResult.xml

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

If ($a) {
    $a.WarningBackgroundColor = $PreviousWarningBackgroundColor
    $a.WarningForegroundColor = $PreviousWarningForegroundColor
    $a.VerboseForegroundColor = $PreviousVerboseForegroundColor
}

Write-Host -NoNewLine "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
