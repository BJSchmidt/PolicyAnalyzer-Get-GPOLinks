<#
    .SYNOPSIS
    Gets report of GPO links for comparison against output from Microsoft Policy Analyzer

    .DESCRIPTION
    Gets report of GPO links for comparison against output from Microsoft Policy Analyzer
    Run Policy Analyzer and export results as XLSX.
    Use that for input into this script.
    A report with GPO Links will be written to OutputFile.
    This makes it possible to quickly see where the GPOs from Policy Analyzer are linked
    and to quickly compare their scopes.

    .EXAMPLE
    Get-GPOLinks.ps1 -InputFile "PolicyAnalyzerExport.xlsx" -OutputFile "PolicyAnalyzerGPOLinks.xlsx"

    .TODO
    Parameterize script
    Allow Pipeline input for InputFile.

#>

<#

Process:
    Import excel document
    Parse out GPO Names from the "%domain%_PolicyRules %newline% GPO" field (column K)
    Get XML Report for all the gpos, once.
    Parse XML for Name, Links, Enforced, GUID
    Export to CSV.
#>

#requires -module ImportExcel

param (
        # Target Domain to get GPO Reports from
        [Alias("Domain")][string]$TargetDomain,
        [Parameter(ValueFromPipeline=$true)][Alias("Input")][string]$InputFile,
        [Alias("Output")][string]$OutputFile
)

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
        [Alias("Input")]
        [string]$InputFile,

        [Parameter(Mandatory)]
        [Alias("Output")]
        [string]$OutputFile
    )

    Process {
        
        #Vars:
        $errorCountGPOReport = 0
        $GPOReportErrors = @()

        Write-Output "`nTargetDomain: $TargetDomain"
        Write-Output "Input File: $InputFile"
        Write-Output "Output File: $OutputFile"
       
        $confirmation = Read-Host -Prompt "Continue? (Y)es / Press any other key to exit"
        if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
            exit
        }

        $Excel = Import-Excel $InputFile

        # Get all Unique GPO Names:
        $allGPONames = $($($excel | Select-Object -ExpandProperty "*PolicyRules*GPO").split("`n")).trim() | Select-Object -Unique

        Write-Host "`nGetting GPO Reports for $($allGPONames.count) GPOs:"
        #Then get their XML Reports & store them:

        $count = 0
        $gpoXMLArray = foreach ($gpo in $allGPONames) {
            $count++
            Write-Progress -Id 1 -Activity "Getting GPO Reports" -Status "Getting GPO Report $count of $($allGPONames.count): $gpo" -PercentComplete $([math]::Round(($count/$allGPONames.count)*100)) -CurrentOperation "$([math]::Round(($count/$allGPONames.count)*100))% Complete"
            try {
                [XML]$(Get-GPOReport -Name $gpo -Domain $targetDomain -ReportType XML -ErrorAction Stop)
                # Write-Host "Got GPO Report for $gpo"
            } catch {
                Write-Host "Could not get GPO Report for $gpo. Was it removed since the backup was created?"
                $errorCountGPOReport++
                $GPOReportErrors += $gpo
            }
        }
        Write-Progress -Id 1 -Completed -Activity "Getting GPO Reports"
        if ($errorCountGPOReport) {
            Write-Host "Could not get $errorCountGPOReport / $($allGPONames.count) GPO Reports"
        }

        Write-Host "Processing GPOs"
               
        # Loop through each row in $Excel, so we can generate output in the same order as the input file.
        $currRow = 2
        $results = foreach ($row in $Excel) {
            # Get GPOs in this cell:
            $thisCellsGPOs = $($($row | Select-Object -ExpandProperty "*PolicyRules*GPO").split("`n")).trim()
            # Write-Host "Row: $currRow | GPOs in this cell: $($thisCellsGPOs.count) | $thisCellsGPOs"
            
            foreach ($gpo in $thisCellsGPOs) {
                # Match to an item in $gpoXMLArray, then output that.
                $gpoXMLArray.GPO | Where-Object {$_.Name -eq $gpo} |
                Select-Object -Property @{Name="Analyzer Source Row"; Expression = {$currRow}},
                                                Name,
                                        @{Name="Enabled"; Expression = {[string]::join("`r`n",($_.LinksTo.Enabled))}},
                                        @{Name="Enforced"; Expression = {[string]::join("`r`n",($_.LinksTo.NoOverride))}},
                                        @{Name="GUID"; Expression = {$_.Identifier.Identifier."#Text"}},
                                        @{Name="Policy Setting Name"; Expression = {$row | Select-Object -ExpandProperty "Policy Setting Name"}},
                                        @{Name="Policy Setting"; Expression = {$row | Select-Object -ExpandProperty "Policy Setting"}},
                                        # @{Name="Policy Rules"; Expression = {$row | Select-Object -ExpandProperty "*_PolicyRules"}},
                                        # @{Name="Policy Rules Option"; Expression = {$row | Select-Object -ExpandProperty "*_PolicyRules*Option"}},
                                        @{Name="Links (Expand row height to see multiline cells) "; Expression = {[string]::join("`r`n",($_.LinksTo.SOMPath))}}
                                        # @{Name="Links (Expand row height to see multiline cells)"; Expression = {$_.LinksTo.SOMPath}}
            }
            $currRow++        
        }
       
        # Export $Results:
        # Color based on $currRow % 2
        Write-Host "Writing Output file $OutputFile"

        # Delete Outputfile if it already exists
        if (Test-Path $OutputFile)  {
            Remove-Item $OutputFile
        }

        #Export Results:
        $results | Export-Excel $OutputFile -WorkSheetname "GPO Links" -CellStyleSB {
            param (
                $workSheet,
                $totalRows,
                $LastColumn
            )
        
            foreach ($row in (2..$totalRows)) {
                if ($workSheet.cells[$row,1].Value % 2 -eq 0) {
                    Set-CellStyle $workSheet $row $lastColumn Solid Gray
                } elseif ($workSheet.cells[$row,1].Value % 2 -eq 1) {
                    Set-CellStyle $workSheet $row $lastColumn Solid LightGray
                }
            }
        }

        #Export list of gpos that we could not get a report for.
        if ($GPOReportErrors) { $GPOReportErrors | Export-Excel $OutputFile -WorkSheetname "Missing Reports" -AutoSize -show }
    }
}

if ($TargetDomain -like "") {
    #Prompt for TargetDomain:
    $TargetDomain = Read-Host -Prompt "Enter Target Domain"
}

if ($InputFile -like "") {
    #Prompt for InputFile:
    $InputFile = Get-FileName($PSScriptRoot)
}

if ($OutputFile -like "" -or ((-not($OutputFile -like "*.xlsx")) -and (-not($OutputFile -like "*.XLSX")))) {
    #Prompt for OutputFile:
    While(-not($OutputFile -like "*.xlsx") -and (-not($OutputFile -like "*.XLSX"))){
        $OutputFile = Read-Host -Prompt "OutputFile (Path and XLSX extension) Ex '.\GPOLinksReport.xlsx'" 
    }
}

Get-GPOLinks -TargetDomain $TargetDomain -InputFile $inputFile -outputFile $OutputFile