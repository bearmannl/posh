<#
.SYNOPSIS
	Parses the output from SP health analyzer missing dependencies.
	
.DESCRIPTION
	Parses the output from the SharePoint Health Analyzer missing dependencies message, splitting it into useful input files for other scripts.

.PARAMETER InputFilePath
    Optionally, override the filepath for the input file.

.PARAMETER OutputFilePath
    Optionally, override the filepath for the output files.

.EXAMPLE
    Parse-SPMissingDependencies.ps1
    Parse-SPMissingDependencies.ps1 -InputFilePath "C:\Temp\MyInput.txt"
    Parse-SPMissingDependencies.ps1 -InputFilePath "C:\Temp\MyInput.txt" -OutputFilePath "C:\Temp"

.LINK
    http://bearman.nl
    http://rubicon.nl

.NOTES
    Authors : Mike Beerman
	Company	: Rubicon B.V.
    Date    : 2017-10-24
    Version : 1.0
#>
param(
    [string]$InputFilePath = ".\Input.txt",
    [string]$OutputPath = ".\"
)

if($null -eq ($OutputPath -Match '.+?\\$')) { $OutputPath = $OutputPath + "\" }

$inputContent = Get-Content $InputFilePath

if ($null -ne $inputContent -and ($inputContent | Measure-Object).Count -gt 0) {
    $missingAssemblies = @()
    $missingFeatures = @()
    $missingSetupFiles = @()
    $missingWebParts = @()
    $missingSiteDefs = @()
    $orphanedSites = @()
    $missingUnmatched = @()
    
    foreach ($row in $inputContent) {
        if ($null -ne $row -and $row.Length -gt 0 -and $row -match '^\[.+?') {
            if ([regex]::matches($row, '\[[^\]]+\]').Groups[0].Value -eq "[MissingAssembly]") {
                $missingAssemblies += "$([regex]::matches($row, '\[[^\]]+\]').Groups[2].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+')"
            }
            elseif ([regex]::matches($row, '\[[^\]]+\]').Groups[0].Value -eq "[MissingFeature]") {
                $missingFeatures += "$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[2].Value -Replace '[[\]]+')"
            }
            elseif ([regex]::matches($row, '\[[^\]]+\]').Groups[0].Value -eq "[MissingSetupFile]") {
                $missingSetupFiles += "$([regex]::matches($row, '\[[^\]]+\]').Groups[3].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+')"
            }
            elseif ([regex]::matches($row, '\[[^\]]+\]').Groups[0].Value -eq "[MissingWebPart]") {
                if (([regex]::matches($row, '\[[^\]]+\]').Count -eq 5)) {
                    $missingWebParts += "$([regex]::matches($row, '\[[^\]]+\]').Groups[3].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+')"
                }
                elseif (([regex]::matches($row, '\[[^\]]+\]').Count -eq 7)) {
                    $missingWebParts += "$([regex]::matches($row, '\[[^\]]+\]').Groups[5].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+')"
                }
                else {
                    $missingUnmatched += $row
                }
            }
            elseif ([regex]::matches($row, '\[[^\]]+\]').Groups[0].Value -eq "[SiteOrphan]") {
                $orphanedSites += "$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[2].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[3].Value -Replace '[[\]]+')"
            }
            elseif ([regex]::matches($row, '\[[^\]]+\]').Groups[0].Value -eq "[MissingSiteDefinition]") {
                $missingSiteDefs += "$([regex]::matches($row, '\[[^\]]+\]').Groups[1].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[2].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[3].Value -Replace '[[\]]+');$([regex]::matches($row, '\[[^\]]+\]').Groups[4].Value -Replace '[[\]]+')"
            }
            else {
                $missingUnmatched += $row
            }
        }
    }
    
    if ($missingAssemblies.Length -gt 0) {
        Write-Host "Missing Assemblies: $($missingAssemblies.Count)"
        $missingAssemblies > "$($OutputPath)InputMissingAssembly.txt"
        Write-Host ""
    }
    
    if ($missingFeatures.Length -gt 0) {
        Write-Host "Missing Features: $($missingFeatures.Count)"
        $missingFeatures > "$($OutputPath)InputMissingFeature.txt"
        Write-Host ""
    }
    
    if ($missingSetupFiles.Length -gt 0) {
        Write-Host "Missing Setup Files: $($missingSetupFiles.Count)"
        $missingSetupFiles > "$($OutputPath)InputMissingSetupFile.txt"
        Write-Host ""
    }
    
    if ($missingWebParts.Length -gt 0) {
        Write-Host "Missing Web Parts: $($missingWebParts.Count)"
        $missingWebParts > "$($OutputPath)InputMissingWebPart.txt"
        Write-Host ""
    }
    if ($missingSiteDefs.Length -gt 0) {
        Write-Host "Missing Site Definitions: $($missingSiteDefs.Count)"
        $missingSiteDefs > "$($OutputPath)InputMissingSiteDefinitions.txt"
        Write-Host ""
    }
    if ($orphanedSites.Length -gt 0) {
        Write-Host "Orphaned sites: $($orphanedSites.Count)"
        $orphanedSites > "$($OutputPath)OrphanedSites.txt"
        Write-Host ""
    }
    if ($missingUnmatched.Length -gt 0) {
        Write-Host "Unmatched entries: $($missingUnmatched.Count)"
        $missingUnmatched > "$($OutputPath)UnmatchedMissingDependencies.txt"
    }   
}