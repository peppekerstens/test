#Requires -Version 5
#Requires -RunAsAdministrator

Invoke-PackageBootstrap {
    [CmdletBinding()]
    param(
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
        [Switch]$MoreVerbose
    )

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

    #Install minimum needed functionality for packagemanagement and ITON specific functions
    $prereqConditions = @(
        @{
            Label  = 'minimum OS level 2012'
            Test   = {
                #Detect OS version/type
                $OSVersion = [System.Environment]::OSVersion.Version
                [System.Version]"$($OSVersion.Major).$($OSVersion.Minor)" -ge [System.Version]'6.2'
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "OS does not meet minimum requirements." -Path $Log -Level 'Error' -Verbose:$MoreVerbose)| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($ScriptName) -EventID $EventID -EntryType Error -Message "OS does not meet minimum requirements" -Verbose:$MoreVerbose)
                Throw "OS does not meet minimum requirements!"
            }
        },
        @{
            Label  = 'Internet connection'
            Test   = {
                Try {
                    Test-Connection -ComputerName 8.8.8.8 -Count 1 -ErrorAction Stop -Verbose:$MoreVerbose
                    $true
                }
                Catch {
                    $false
                }
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "No Internet connection." -Path $Log -Level 'Error' -Verbose:$MoreVerbose)| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($ScriptName) -EventID $EventID -EntryType Error -Message "No Internet connection" -Verbose:$MoreVerbose)
                Throw "No Internet connection!"
            }
        },
        @{
            Label  = '64bit PSSession'
            Test   = {
                [Environment]::Is64BitProcess
            }
            Action = {
                ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "No 64bit PSSession." -Path $Log -Level 'Error' -Verbose:$MoreVerbose)| Out-Null
                ($EventLogName) -and (Write-EventLog -LogName $EventLogName -Source $($ScriptName) -EventID $EventID -EntryType Error -Message "No 64bit PSSession")
                Throw "No 64bit PSSession!"
            }
        },
        @{
            Label  = 'latest NuGet provider installed'
            Test   = {
                Try {
                    $Latest = Find-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop -Verbose:$MoreVerbose
                    $Installed = Get-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop -Verbose:$MoreVerbose
                    $Installed.Version -ge $Latest.Version
                }
                Catch {
                    $false
                }
            }
            Action = {
                Install-PackageProvider -Name 'NuGet' -Force -Verbose:$MoreVerbose #if not on system, just force install
            }
        },
        @{
            Label  = 'latest PowerShellGet and PackageManagement installed'
            Test   = {
                (Update-Package -Name 'PowerShellGet' -Source $PSGallery.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled
            }
            Action = {
                Update-Package -Name 'PowerShellGet' -Source $PSGallery.Name -AllVersions -Force -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'latest Pester module installed'
            Test   = {
                (Update-Package -Name 'Pester' -Source $PSGallery.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled
            }
            Action = {
                #If this line not added, then a verbose load of the module happens...
                Import-Module PackageManagement -Verbose:$MoreVerbose
                Install-Package -Name 'Pester' -Source $PSGallery.Name -Force -SkipPublisherCheck -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'Set PSGalleryITON as trusted PSRepository'
            Test   = {
                #If this line not added, then a verbose load of the module happens...
                Import-Module PackageManagement -Verbose:$MoreVerbose
                (Get-PSRepository -Verbose:$MoreVerbose).Name -contains $PSGalleryIton.Name
            }
            Action = {
                Register-PSRepository -Name $PSGalleryIton.Name -SourceLocation $PSGalleryIton.SourceLocation -PublishLocation $PSGalleryIton.PublishLocation -Verbose:$MoreVerbose
                #If this line not added, then a verbose load of the module happens...
                Import-Module PackageManagement -Verbose:$MoreVerbose
                Set-PackageSource -Name $PSGalleryIton.Name -Trusted -Verbose:$MoreVerbose
            }
        },
        @{
            Label  = 'latest GenericFunctions module installed'
            Test   = {
                (Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled
            }
            Action = {
                Update-Package -Name 'GenericFunctions' -Source $PSGalleryIton.Name -AllVersions -Force -Verbose:$MoreVerbose
            }
        }
    )

    $WriteLogPresent = (Get-Command Write-Log -ErrorAction SilentlyContinue -Verbose:$MoreVerbose) -and $true

    Write-Verbose "Preparing minimum requirements for installation"
    ($Log) -and ($WriteLogPresent) -and (Write-Log -Message "Preparing minimum requirements for installation" -Path $Log -Verbose:$MoreVerbose) | Out-Null

    @($prereqConditions).foreach( {
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

    If ($a) {
        $a.WarningBackgroundColor = $PreviousWarningBackgroundColor
        $a.WarningForegroundColor = $PreviousWarningForegroundColor
        $a.VerboseForegroundColor = $PreviousVerboseForegroundColor
    }

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


function Write-Log {
    <#
    .Synopsis
    Write-Log writes a message to a specified log file with the current time stamp.
    .DESCRIPTION
    The Write-Log function is designed to add logging capability to other scripts.
    In addition to writing output and/or verbose you can write to a log file for
    later debugging.
    .NOTES
    Created by: Jason Wasser @wasserja
    Modified: 11/24/2015 09:30:19 AM  

    Changelog:
        * Code simplification and clarification - thanks to @juneb_get_help
        * Added documentation.
        * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks
        * Revised the Force switch to work as it should - thanks to @JeffHicks

    To Do:
        * Add error handling if trying to create a log file in a inaccessible location.
        * Add ability to write $Message to $Verbose or $Error pipelines to eliminate
        duplicates.
    .PARAMETER Message
    Message is the content that you wish to add to the log file. 
    .PARAMETER Path
    The path to the log file to which you would like to write. By default the function will 
    create the path and file if it does not exist. 
    .PARAMETER Level
    Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational)
    .PARAMETER NoClobber
    Use NoClobber if you do not wish to overwrite an existing file.
    .EXAMPLE
    Write-Log -Message 'Log message' 
    Writes the message to c:\Logs\PowerShellLog.log.
    .EXAMPLE
    Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log
    Writes the content to the specified log file and creates the path and file specified. 
    .EXAMPLE
    Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error
    Writes the message to the specified log file as an error message, and writes the message to the error pipeline.
    .LINK
    https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [Alias('LogPath')]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warning", "Info")]
        [string]$Level = "Info",
        
        [Parameter(Mandatory = $false)]
        [switch]$NoClobber
    )

    Begin {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
    }

    Process {
        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
        }
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            New-Item $Path -Force -ItemType File | Out-Null
        }
        else {
            # Nothing to see here yet.
        }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write log entry to $Path
        "$FormattedDate $($Level.ToUpper()): $Message" | Out-File -FilePath $Path -Append
    }
    End { }
}

