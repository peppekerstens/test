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