<#
Parameters:
    Policy Analyzer Exported Excel File.
    TargetDomain

Process:
    Import excel document
    Parse out GPO Names from the "baudette-mn_PolicyRules %newline% GPO" field (column K)
    Get XML Report for all the gpos, once.

    Parse XML for Name, Links, Enforced, GUID
    Export to CSV.

TODO: Prompt for TargetDomain (~ Line 30)
#>

#requires -module ImportExcel

function Get-FileName($initialDirectory) {
    $initialDirectory = $PSscriptroot
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Open Policy Analyzer Exported XLSX"
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "XLSX (*.XLSX)| *.XLSX"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

function Get-GPOLinks() {
    param (
        # Target Domain to get GPO Reports from
        [Parameter(Mandatory)]
        [string]
        $TargetDomain,

        [Parameter(Mandatory)]
        [Alias("Input","File")]
        [string]$InputFile

        # [array]$GPONames
    )

    Process{        
        
        Write-Output "TargetDomain: $TargetDomain"
        # Write-Output "GPOs: $GPONames"
        Write-Output "Input File: $InputFile"
        
        $confirmation = Read-Host -Prompt "Continue? (Y)es / Press any other key to exit"
        if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
            exit
        }
        
        $Excel = Import-Excel $InputFile
        
################################################        
        # Get all Unique GPO Names:
        $allGPONames = $($($excel | Select-Object -ExpandProperty "*PolicyRules*GPO").split("`n")).trim() | Select-Object -Unique

        #Then get their XML Reports & store them:
        $gpoXMLArray = foreach ($gpo in $allGPONames) {
            try{
                [XML]$(Get-GPOReport -Name $gpo -Domain $targetDomain -ReportType XML)
                Write-Host "Got GPO Report for $gpo"
            }
            catch{
                Write-Host "Could not get GPO Report for $gpo"
            }
        }

        # Then loop through each row in excel, so we can generate output in the same order as the input file.







        
###########################
        #Get GPO Names:
        $currRow = 0
        $results = foreach ($row in $Excel) {
            $currRow++
            $gpoNames = $row | Select-Object -ExpandProperty "*PolicyRules*GPO"
            # Write-Host $gpoNames
            $gpoXMLArray = foreach ($gpo in $gpoNames.split("`n")) {
                try{
                    [XML]$(Get-GPOReport -Name $gpo -Domain $targetDomain -ReportType XML)
                    Write-Host "Got GPO Report for $gpo"
                }
                catch{
                    Write-Host "Could not get GPO Report for $gpo"
                }
            }
            
            $gpoXMLArray.GPO |
            Select-Object -Property Name,
                @{Name="Links (Expand row height to see multiline cells) "; Expression = {[string]::join("`n",($_.LinksTo.SOMPath))}},
                @{Name="Enabled"; Expression = {[string]::join("`n",($_.LinksTo.Enabled))}},
                @{Name="Enforced"; Expression = {[string]::join("`n",($_.LinksTo.NoOverride))}},
                @{Name="GUID"; Expression = {$_.Identifier.Identifier."#Text" }},
                @{Name="currRow"; Expression = {$currRow % 2}}
        }
        
        # Export:
        # Color based on $currRow % 2
        $results | Export-Excel .\test.xlsx -show -CellStyleSB {
            param (
                $workSheet,
                $totalRows,
                $LastColumn
            )
        
            foreach ($row in (2..$totalRows)) {
                # Set CellStyle (alternating row colors):
                # if ($row % 2 -eq 0) {
                #     Set-CellStyle $workSheet $row $LastColumn Solid Gray
                # } elseif ($row % 2 -eq 1) {
                #     Set-CellStyle $workSheet $row $LastColumn Solid LightGray
                # }
        
                # Set CellStyle based on column 6
                if ($workSheet.cells[$row,6].Value -eq 0) {
                    Set-CellStyle $workSheet $row $lastColumn Solid Gray
                } elseif ($workSheet.cells[$row,6].Value -eq 1) {
                    Set-CellStyle $workSheet $row $lastColumn Solid LightGray
                }
            }
        }
    }
}

$TargetDomain
$InputFile

if ($TargetDomain -like "") {
    #Prompt for input file:
    $TargetDomain = "baudette-mn.catholichealth.net"
}

$InputFile = Get-FileName($PSScriptRoot)

Get-GPOLinks -TargetDomain $TargetDomain -InputFile $inputFile