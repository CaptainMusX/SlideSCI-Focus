[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$sourcePath = Join-Path $root "SlideSCI\Ribbon1.Titles.cs"
$designerPath = Join-Path $root "SlideSCI\Ribbon1.Designer.cs"
$source = Get-Content -LiteralPath $sourcePath -Raw
$designer = Get-Content -LiteralPath $designerPath -Raw

function Assert-Layout([bool]$condition, [string]$message)
{
    if (-not $condition) { throw $message }
    Write-Output "PASS $message"
}

Assert-Layout (($source -split "CreateTitleColumnSeparator\(").Count - 1 -eq 3) "two column separators plus factory method"
Assert-Layout ($source -match 'titleTextEditBox\.SizeString = "0000000000"') "title input compensates for combo drop-down width"
Assert-Layout ($source -match 'fontNameEditBox\.Label = "字体"; fontNameEditBox\.SizeString = "00000000"') "font input uses the shared fixed width"
Assert-Layout ($source -match 'rowSlot\.Items\.Add\(row\)') "third row has an independent vertical layout slot"
Assert-Layout (($source -split "CreateRibbonButton\(\)").Count - 1 -ge 2) "alignment menu uses ordinary buttons"
Assert-Layout (-not ($source -match 'CreateRibbonToggleButton\(\)')) "alignment menu has no checked toggle items"
Assert-Layout ($source -match 'titleAlignmentMenu\.OfficeImageId') "collapsed menu icon follows current alignment"
Assert-Layout ($designer -match 'CreateRibbonToggleButton\(\)') "grouping control is a toggle button"
Write-Output "TOTAL PASS 8"
