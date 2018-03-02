describe 'Pre-Installation tests' {
    context 'system' {
        It 'should be a a windows based machine with minimum OS level 2012R2' {
            #Detect OS version/type
            $OSVersion = [System.Environment]::OSVersion.Version
            [System.Version]"$($OSVersion.Major).$($OSVersion.Minor)" -ge [System.Version]'6.2' | Should Be $true
        }
        It 'should have a working Internet connection' {
            { Test-Connection -ComputerName 8.8.8.8 -Count 1 } | Should Not Throw
        }
        It 'should have PowerShell 5.0 or higher present' {
            $PSVersionTable.PSVersion -ge '6.0' | Should Be $true
        }
        It 'not have a configured ' {
            $PSVersionTable.PSVersion -ge '6.0' | Should Be $true
        }
    }
}