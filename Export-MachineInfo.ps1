function Export-MachineInfo {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ExcelFilePath,

        [string]$OutputDir = "./Output",

        [int]$VMIndex,

        [switch]$ExcludeNoDataVMs,

        [switch]$ExcludePoweredOff,

        [switch]$ExcludeTemplates,

        [switch]$ExcludeSRM,

        [switch]$VerboseLogging = $false
    )

    Write-Information "Input file: $ExcelFilePath"
    Write-Information "Output directory: $OutputDir"
    Write-Warning "Exclude VMs with no additional data: $($ExcludeNoDataVMs.IsPresent)"
    Write-Warning "Exclude powered-off VMs: $($ExcludePoweredOff.IsPresent)"
    Write-Warning "Exclude templates: $($ExcludeTemplates.IsPresent)"
    Write-Warning "Exclude SRM placeholders: $($ExcludeSRM.IsPresent)"
    Write-Warning "Verbose logging: $($VerboseLogging.IsPresent)"

    if (-not (Test-Path -Path $OutputDir -PathType Container)) {
        New-Item -Path $OutputDir -ItemType Directory
    }

    $ExcelSheets = (Open-ExcelPackage -Path $ExcelFilePath).Workbook.Worksheets.Name

    $AllData = @{}
    foreach ($Sheet in $ExcelSheets) {
        $AllData[$Sheet] = Import-Excel -Path $ExcelFilePath -WorksheetName $Sheet
    }

    $VInfo = $AllData["vInfo"]

    if ($null -ne $VMIndex) {
        $VInfo = $VInfo | Select-Object -Index ($VMIndex - 1)
    }

    foreach ($Machine in $VInfo) {
        $MachineName = $Machine.VM

        if ($VerboseLogging) {
            Write-Host -f Yellow "Processing VM: $MachineName..."
        }

        $ProcessVM = (-not $ExcludePoweredOff -or $Machine.Powerstate -ne "poweredOff") -and
                     (-not $ExcludeTemplates -or $Machine.Template -ne "True") -and
                     (-not $ExcludeSRM -or $Machine."SRM Placeholder" -ne "True")

                     if (-not $ProcessVM) {
                        $skipReasons = @()
                        if ($ExcludePoweredOff -and $Machine.Powerstate -eq "poweredOff") {
                            $skipReasons += "Powered off"
                        }
                        if ($ExcludeTemplates -and $Machine.Template -eq "True") {
                            $skipReasons += "Is a template"
                        }
                        if ($ExcludeSRM -and $Machine."SRM Placeholder" -eq "True") {
                            $skipReasons += "Is an SRM placeholder"
                        }
                        $skipReason = ($skipReasons -join ", ")
                        
                        Write-Host -f Red "Skipping VM: $MachineName due to exclusion conditions: ($skipReason)."
                        continue
                    }
                    

        $MachineInfo = [PSCustomObject]@{}

        foreach ($Property in $Machine.PSObject.Properties) {
            $MachineInfo | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value
        }

        $OtherSheets = $ExcelSheets | Where-Object { $_ -ne "vInfo" }

        $DataFound = $false

        foreach ($Sheet in $OtherSheets) {
            $SheetData = $AllData[$Sheet]
            $MatchingData = $SheetData | Where-Object { $_.VM -eq $MachineName }
            
            if ($null -ne $MatchingData) {
                if ($VerboseLogging) {
                    Write-Host -f Yellow "Adding data from $Sheet for VM: $MachineName..."
                }
                
                foreach ($Property in $MatchingData.PSObject.Properties) {
                    $MachineInfo | Add-Member -MemberType NoteProperty -Name ($Sheet + "_" + $Property.Name) -Value $Property.Value
                }

                $DataFound = $true
            } else {
                Write-Host -f Yellow "No data found in $Sheet for VM: $MachineName"
            }
        }
        
        if (-not $ExcludeNoDataVMs -or $DataFound) {
            $CsvFileName = Join-Path -Path $OutputDir -ChildPath ("$MachineName.csv")
            $MachineInfo | Export-Csv -Path $CsvFileName -NoTypeInformation -Force

            Write-Host -f Green "Exported to CSV:" $CsvFileName
        } else {
            Write-Host -f Red "Skipping $MachineName - No additional data found."
        }
    }

    Write-Host -f Yellow "Function execution completed!"
}
