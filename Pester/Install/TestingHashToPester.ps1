$ScriptName = $MyInvocation.MyCommand.Name
$Conditions = @(
    @{
        Describe = "Dit is een test"
        It       = 'should have latest AzureRM module installed'
        Test     = "{
            (Update-Package -Name 'AzureRM' -Source $PSGallery.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled | Should Be $true
        }"
        Action   = "{
            #AzureRm is a 'special' module which needs other update method
            Update-AzureRM -Source $PSGallery.Name -Verbose:$MoreVerbose
        }"
    },
    @{
        Describe = "Dit is een test"
        It       = 'should have latest AzureRmStorageTable installed'
        Test     = "{
            (Update-Package -Name 'AzureRmStorageTable' -Source $PSGallery.Name -WhatIf -Verbose:$MoreVerbose).IsInstalled | Should Be $true
        }"
        Action   = "{
            Update-Package -Name 'AzureRmStorageTable' -Source $PSGallery.Name -AllVersions -Force -Verbose:$MoreVerbose
        }"
    }
)

<#
$ht = @{'One'=1;'Two'=2}
$results = @()
$keys = $ht.keys
foreach ($key in $keys) {
$results += New-Object psobject -Property @{'Number'=$key;'Value'=$ht[$key]}
}
$results
#>

function ConvertHashtableTo-Object {
    [CmdletBinding()]
    Param([Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $True)]
        [hashtable]$ht
    )
    PROCESS {
        $results = @()

        $ht | % {
            $result = New-Object psobject;
            foreach ($key in $_.keys) {
                $result | Add-Member -MemberType NoteProperty -Name $key -Value $_[$key]
            }
            $results += $result;
        }
        return $results
    }
}

function ConvertTo-PsCustomObjectFromHashtable { 
    param ( 
        [Parameter(  
            Position = 0,   
            Mandatory = $true,   
            ValueFromPipeline = $true,  
            ValueFromPipelineByPropertyName = $true  
        )] [object[]]$hashtable 
    ); 
 
    begin { $i = 0; } 
 
    process { 
        foreach ($myHashtable in $hashtable) { 
            if ($myHashtable.GetType().Name -eq 'hashtable') { 
                $output = New-Object -TypeName PsObject; 
                Add-Member -InputObject $output -MemberType ScriptMethod -Name AddNote -Value {  
                    Add-Member -InputObject $this -MemberType NoteProperty -Name $args[0] -Value $args[1]; 
                }; 
                $myHashtable.Keys | Sort-Object | % {  
                    $output.AddNote($_, $myHashtable.$_);  
                } 
                $output; 
            }
            else { 
                Write-Warning "Index $i is not of type [hashtable]"; 
            } 
            $i += 1;  
        } 
    } 
}

#$Conditions | ConvertHashtableTo-Object
#$Conditions | ConvertTo-PsCustomObjectFromHashtable

#Convert collection/array of hashes to collection/arry of objects
$ObjConditions = @()

#@($Conditions).foreach({$ObjConditions += [PSCustomObject]$_})
@($Conditions).foreach( {$ObjConditions += [PSCustomObject] $_})



#Write-Host "Testing condition [$($_.Describe)]"
#        Write-Host "Testing condition [$($_.It)]"
<#
        ($Log) -and (Write-Log -Message "Testing condition [$($_.Label)]" -Path $Log -Verbose:$MoreVerbose)| Out-Null
        if (-not (& $_.Test)) {
            Write-Warning "Condition [$($_.Label)] failed. Remediating..."
            ($Log) -and (Write-Log -Message "Condition [$($_.Label)] failed. Remediating..." -Path $Log -Level 'Warning' -Verbose:$MoreVerbose)| Out-Null
            & $_.Action
        }
        else {
            Write-Verbose 'Passed.'
            ($Log) -and (Write-Log -Message "Condition [$($_.Label)] passed" -Path $Log -Verbose:$MoreVerbose)| Out-Null
        }#><#
    }) | Out-Null

$ScriptName = $MyInvocation.MyCommand.Name
([PSCustomObject]$Conditions | Group-Object -Property Describe).Group
([PSCustomObject]$Conditions).Action[0]

})
#>
#$ObjConditions

$PesterTest = [String]::Empty
$GroupedTests = $ObjConditions | Group-Object -Property Describe
ForEach ($Describe in $GroupedTests) {
    $PesterTest += "describe `'$($Describe.Name)`' {`n"
    ForEach ($It in $Describe.Group) {
        $PesterTest += "it `'$($It.It)`' $($It.test) {`n"
    }
    $PesterTest += "}`n"
}
$PesterTest