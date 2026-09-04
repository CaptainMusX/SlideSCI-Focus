[CmdletBinding()]
param(
    [string]$BackupDirectory = "",
    [switch]$Preview
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot ".." )).Path
$artifactRoot = Join-Path $repoRoot "artifacts"
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
# Only clean this fork's current and earlier CaptainMusX identities. The
# upstream Achuan-2.SlideSCI/SlideSCI installation is deliberately preserved.
$ownedIdentityPattern = '(?i)captainmusx\.slidesci(?:\.focus)?|slidesci[\s._-]+focus'

function Test-PathUnderRoot
{
    param(
        [Parameter(Mandatory = $true)][string]$Candidate,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $candidateFull = [System.IO.Path]::GetFullPath($Candidate).TrimEnd('\', '/')
    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    return $candidateFull.Equals($rootFull, [System.StringComparison]::OrdinalIgnoreCase) -or
        $candidateFull.StartsWith($rootFull + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Convert-ManifestValueToPath
{
    param([AllowNull()][string]$ManifestValue)

    if ([string]::IsNullOrWhiteSpace($ManifestValue)) { return $null }
    $manifestUri = ($ManifestValue -replace '\|.*$', '')
    if ($manifestUri -notmatch '^file:///') { return $null }

    $rawPath = $manifestUri.Substring(8)
    if ([string]::IsNullOrWhiteSpace($rawPath)) { return $null }
    $rawPath = [Uri]::UnescapeDataString($rawPath).Replace('/', '\')
    if ($rawPath -notmatch '^[A-Za-z]:\\') { return $null }
    return [System.IO.Path]::GetFullPath($rawPath)
}

function Get-RegistryInventory
{
    param([string[]]$Roots)

    $records = @()
    foreach ($root in $Roots)
    {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($key in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue))
        {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            $manifestProperty = $properties.PSObject.Properties['Manifest']
            $manifestValue = if ($manifestProperty) { [string]$manifestProperty.Value } else { '' }
            if ($key.PSChildName -notmatch $ownedIdentityPattern -and $manifestValue -notmatch $ownedIdentityPattern)
            {
                continue
            }

            $values = [ordered]@{}
            foreach ($property in $properties.PSObject.Properties)
            {
                if ($property.Name -notmatch '^PS')
                {
                    $values[$property.Name] = $property.Value
                }
            }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $root
                Name = $key.PSChildName
                Values = $values
                ManifestPath = Convert-ManifestValueToPath $manifestValue
            }
        }
    }
    return @($records | Sort-Object Path -Unique)
}

function Get-ClickOnceBranches
{
    param([string]$ClickOnceRoot)

    if (-not (Test-Path -LiteralPath $ClickOnceRoot -PathType Container)) { return @() }
    $branches = @()
    $candidateFiles = Get-ChildItem -LiteralPath $ClickOnceRoot -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.vsto', '.manifest') }

    foreach ($candidateFile in $candidateFiles)
    {
        try
        {
            $content = Get-Content -LiteralPath $candidateFile.FullName -Raw -ErrorAction Stop
        }
        catch
        {
            continue
        }
        if ($content -notmatch $ownedIdentityPattern) { continue }

        $relativePath = $candidateFile.FullName.Substring($ClickOnceRoot.Length).TrimStart('\', '/')
        $parts = $relativePath -split '[\\/]'
        if ($parts.Count -lt 2) { continue }
        $branch = Join-Path (Join-Path $ClickOnceRoot $parts[0]) $parts[1]
        if (-not (Test-PathUnderRoot -Candidate $branch -Root $ClickOnceRoot)) { continue }
        if ($branch -notin $branches) { $branches += $branch }
    }
    return @($branches)
}

function Get-UninstallInventory
{
    param([string[]]$Roots)

    $records = @()
    foreach ($root in $Roots)
    {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($key in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue))
        {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            $displayProperty = $properties.PSObject.Properties['DisplayName']
            $uninstallProperty = $properties.PSObject.Properties['UninstallString']
            $displayName = if ($displayProperty) { [string]$displayProperty.Value } else { '' }
            $uninstallString = if ($uninstallProperty) { [string]$uninstallProperty.Value } else { '' }
            if ($displayName -notmatch $ownedIdentityPattern -and $uninstallString -notmatch $ownedIdentityPattern)
            {
                continue
            }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $root
                Name = $key.PSChildName
                DisplayName = $displayName
                UninstallString = $uninstallString
            }
        }
    }
    return @($records | Sort-Object Path -Unique)
}

function Get-VstoRegistryInventory
{
    $records = @()
    $inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
    if (Test-Path -LiteralPath $inclusionRoot)
    {
        foreach ($key in (Get-ChildItem -LiteralPath $inclusionRoot -ErrorAction SilentlyContinue))
        {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            $urlProperty = $properties.PSObject.Properties['Url']
            $url = if ($urlProperty) { [string]$urlProperty.Value } else { '' }
            if ($url -notmatch $ownedIdentityPattern) { continue }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $inclusionRoot
                Name = $key.PSChildName
                Kind = 'Inclusion'
                Url = $url
            }
        }
    }

    $metadataRoot = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
    if (Test-Path -LiteralPath $metadataRoot)
    {
        $metadataProperties = Get-ItemProperty -LiteralPath $metadataRoot -ErrorAction SilentlyContinue
        foreach ($property in $metadataProperties.PSObject.Properties)
        {
            if ($property.Name -notmatch '^PS' -and $property.Name -match $ownedIdentityPattern)
            {
                $records += [pscustomobject]@{
                    Path = $metadataRoot
                    Root = $metadataRoot
                    Name = $property.Name
                    Kind = 'SolutionMetadataRootValue'
                    ValueName = $property.Name
                    Url = $property.Name
                    Value = [string]$property.Value
                }
            }
        }

        foreach ($key in (Get-ChildItem -LiteralPath $metadataRoot -ErrorAction SilentlyContinue))
        {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            $matched = $false
            foreach ($propertyName in @('addInName', 'friendlyName', 'description'))
            {
                $property = $properties.PSObject.Properties[$propertyName]
                if ($property -and [string]$property.Value -match $ownedIdentityPattern) { $matched = $true }
            }
            if (-not $matched) { continue }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $metadataRoot
                Name = $key.PSChildName
                Kind = 'SolutionMetadata'
                Url = ''
            }
        }
    }

    $vstaSolutionsRoot = 'HKCU:\Software\Microsoft\VSTA\Solutions'
    if (Test-Path -LiteralPath $vstaSolutionsRoot)
    {
        foreach ($key in (Get-ChildItem -LiteralPath $vstaSolutionsRoot -ErrorAction SilentlyContinue))
        {
            $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            $solutionText = (($properties.PSObject.Properties | Where-Object Name -notmatch '^PS' |
                ForEach-Object { [string]$_.Name + ' ' + [string]$_.Value }) -join ' ')
            if ($solutionText -notmatch $ownedIdentityPattern) { continue }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $vstaSolutionsRoot
                Name = $key.PSChildName
                Kind = 'VstaSolution'
                Url = [string]$properties.Url
            }
        }
    }

    foreach ($addinsDataRoot in @(
        'HKCU:\Software\Microsoft\Office\PowerPoint\AddinsData',
        'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddinsData'
    ))
    {
        if (-not (Test-Path -LiteralPath $addinsDataRoot)) { continue }
        foreach ($key in (Get-ChildItem -LiteralPath $addinsDataRoot -ErrorAction SilentlyContinue))
        {
            if ($key.PSChildName -notmatch $ownedIdentityPattern) { continue }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $addinsDataRoot
                Name = $key.PSChildName
                Kind = 'OfficeAddinsData'
                Url = ''
            }
        }
    }

    foreach ($officeValueRoot in @(
        'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddInLoadTimes',
        'HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache'
    ))
    {
        if (-not (Test-Path -LiteralPath $officeValueRoot)) { continue }
        $properties = Get-ItemProperty -LiteralPath $officeValueRoot -ErrorAction SilentlyContinue
        foreach ($property in $properties.PSObject.Properties)
        {
            if ($property.Name -match '^PS' -or $property.Name -notmatch $ownedIdentityPattern) { continue }
            $records += [pscustomobject]@{
                Path = $officeValueRoot
                Root = $officeValueRoot
                Name = $property.Name
                Kind = 'OfficeRootValue'
                ValueName = $property.Name
                Url = ''
                Value = $property.Value
            }
        }
    }
    # Keep root-value records distinct even when they share the same key path.
    # The removal phase de-duplicates key targets separately.
    return @($records)
}

function Get-ClickOnceRegistryInventory
{
    $records = @()
    $baseRoot = 'Registry::HKEY_CURRENT_USER\Software\Classes\Software\Microsoft\Windows\CurrentVersion\Deployment\SideBySide\2.0'
    if (-not (Test-Path -LiteralPath $baseRoot)) { return @() }

    $identityPattern = $ownedIdentityPattern
    $keys = @((Get-Item -LiteralPath $baseRoot)) +
        @(Get-ChildItem -LiteralPath $baseRoot -Recurse -ErrorAction SilentlyContinue)

    foreach ($key in $keys)
    {
        if ($key.PSChildName -match $identityPattern)
        {
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $baseRoot
                Name = $key.PSChildName
                Kind = 'ClickOnceRegistryKey'
            }
        }

        $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
        foreach ($property in ($properties.PSObject.Properties | Where-Object Name -notmatch '^PS'))
        {
            $searchParts = @([string]$property.Name)
            if ($property.Value -is [byte[]])
            {
                $bytes = [byte[]]$property.Value
                $searchParts += [System.Text.Encoding]::Unicode.GetString($bytes)
                $searchParts += [System.Text.Encoding]::UTF8.GetString($bytes)
                $searchParts += [System.Text.Encoding]::ASCII.GetString($bytes)
            }
            else
            {
                $searchParts += [string]$property.Value
            }
            if (($searchParts -join ' ') -notmatch $identityPattern) { continue }
            $records += [pscustomobject]@{
                Path = $key.PSPath
                Root = $baseRoot
                Name = $key.PSChildName
                Kind = 'ClickOnceRegistryValue'
                ValueName = $property.Name
            }
        }
    }
    return @($records)
}

$powerPointProcesses = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
if ($powerPointProcesses.Count -gt 0)
{
    $processIds = ($powerPointProcesses | Select-Object -ExpandProperty Id) -join ', '
    throw "PowerPoint 正在运行（PID: $processIds）。请保存文件并退出 PowerPoint 后重试；脚本不会强制结束它。"
}

$registryRoots = @(
    'HKCU:\Software\Microsoft\Office\PowerPoint\Addins',
    'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Addins',
    'HKCU:\Software\kingsoft\Office\WPP\Addins',
    'HKCU:\Software\kingsoft\Office\WPP\AddinsWL',
    'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\PowerPoint\Addins',
    'Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins',
    'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins',
    'Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Addins',
    'Registry::HKEY_LOCAL_MACHINE\Software\kingsoft\Office\WPP\Addins',
    'Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\kingsoft\Office\WPP\Addins'
)
$registryRecords = @(Get-RegistryInventory $registryRoots)
$uninstallRoots = @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
$uninstallRecords = @(Get-UninstallInventory $uninstallRoots)
$vstoRecords = @(Get-VstoRegistryInventory)
$clickOnceRegistryRecords = @(Get-ClickOnceRegistryInventory)

$manifestFiles = @($registryRecords | Where-Object { $_.ManifestPath }) |
    Select-Object -ExpandProperty ManifestPath -Unique
$externalManifestPaths = @()
foreach ($manifestFile in $manifestFiles)
{
    if (-not (Test-Path -LiteralPath $manifestFile -PathType Leaf)) { continue }
    if (Test-PathUnderRoot -Candidate $manifestFile -Root $repoRoot)
    {
        continue
    }
    if ([System.IO.Path]::GetFileName($manifestFile) -match '(?i)^CaptainMusX\.SlideSCI(?:\.Focus)?\.vsto$')
    {
        $externalManifestPaths += $manifestFile
    }
}

$knownExternalManifestPaths = @(
    'C:\Program Files\CaptainMusX.SlideSCI.Focus.vsto',
    'C:\Program Files (x86)\CaptainMusX.SlideSCI.Focus.vsto',
    'D:\Program Files\CaptainMusX.SlideSCI.Focus.vsto',
    'D:\Program Files (x86)\CaptainMusX.SlideSCI.Focus.vsto',
    'C:\Program Files\CaptainMusX.SlideSCI.vsto',
    'C:\Program Files (x86)\CaptainMusX.SlideSCI.vsto',
    'D:\Program Files\CaptainMusX.SlideSCI.vsto',
    'D:\Program Files (x86)\CaptainMusX.SlideSCI.vsto',
    (Join-Path $env:LOCALAPPDATA 'Programs\SlideSCI Focus\CaptainMusX.SlideSCI.Focus.vsto')
)
foreach ($knownPath in $knownExternalManifestPaths)
{
    if ((Test-Path -LiteralPath $knownPath -PathType Leaf) -and $knownPath -notin $externalManifestPaths)
    {
        $externalManifestPaths += $knownPath
    }
}

$knownExternalDirectories = @(
    'C:\Program Files\SlideSCI Focus',
    'C:\Program Files (x86)\SlideSCI Focus',
    'D:\Program Files\SlideSCI Focus',
    'D:\Program Files (x86)\SlideSCI Focus',
    (Join-Path $env:LOCALAPPDATA 'Programs\SlideSCI Focus'),
    'F:\SlideSCI-Installed'
)
$externalDirectories = @($knownExternalDirectories | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
foreach ($manifestTarget in $externalManifestPaths)
{
    $manifestParent = [System.IO.Path]::GetDirectoryName($manifestTarget)
    $hasInnoMarker = (Test-Path -LiteralPath (Join-Path $manifestParent 'unins000.exe') -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $manifestParent 'Application Files') -PathType Container)
    if ($hasInnoMarker -and -not (Test-PathUnderRoot -Candidate $manifestParent -Root $repoRoot) -and $manifestParent -notin $externalDirectories)
    {
        $externalDirectories += $manifestParent
    }
}

$clickOnceRoot = Join-Path $env:LOCALAPPDATA 'Apps\2.0'
$clickOnceBranches = @(Get-ClickOnceBranches $clickOnceRoot)

if ([string]::IsNullOrWhiteSpace($BackupDirectory))
{
    $BackupDirectory = Join-Path $artifactRoot "pre-install-cleanup-$timestamp"
}
$backupPath = [System.IO.Path]::GetFullPath($BackupDirectory)
New-Item -ItemType Directory -Force -Path $backupPath | Out-Null

$inventory = [pscustomobject]@{
    Timestamp = (Get-Date).ToString('o')
    Repository = $repoRoot
    Registry = $registryRecords
    UninstallRegistry = $uninstallRecords
    VstoRegistry = $vstoRecords
    ClickOnceRegistry = $clickOnceRegistryRecords
    ExternalManifestFiles = $externalManifestPaths
    ExternalDirectories = $externalDirectories
    ClickOnceBranches = $clickOnceBranches
}
$inventoryPath = Join-Path $backupPath 'inventory.json'
[System.IO.File]::WriteAllText(
    $inventoryPath,
    ($inventory | ConvertTo-Json -Depth 8),
    (New-Object System.Text.UTF8Encoding -ArgumentList $false)
)

$actions = [System.Collections.Generic.List[string]]::new()
$failures = [System.Collections.Generic.List[string]]::new()
$actions.Add("Backup: $inventoryPath")

$registryTargets = @(
    (@($registryRecords | Select-Object -ExpandProperty Path) +
        @($uninstallRecords | Select-Object -ExpandProperty Path) +
        @($vstoRecords | Where-Object { $_.Kind -notin @('SolutionMetadataRootValue', 'OfficeRootValue') } | Select-Object -ExpandProperty Path) +
        @($clickOnceRegistryRecords | Where-Object { $_.Kind -ne 'ClickOnceRegistryValue' } | Select-Object -ExpandProperty Path)) |
        Sort-Object -Unique
)
$registryValueTargets = @(
    @($vstoRecords | Where-Object { $_.Kind -in @('SolutionMetadataRootValue', 'OfficeRootValue') }) +
    @($clickOnceRegistryRecords | Where-Object { $_.Kind -eq 'ClickOnceRegistryValue' })
)
foreach ($registryTarget in $registryTargets)
{
    $actions.Add("Registry target: $registryTarget")
}
foreach ($registryValueTarget in $registryValueTargets)
{
    $actions.Add("Registry value target: $($registryValueTarget.Path)::$($registryValueTarget.ValueName)")
}
foreach ($manifestTarget in $externalManifestPaths)
{
    $actions.Add("File target: $manifestTarget")
}
foreach ($directoryTarget in $externalDirectories)
{
    $actions.Add("Directory target: $directoryTarget")
}
foreach ($cacheTarget in $clickOnceBranches)
{
    $actions.Add("ClickOnce target: $cacheTarget")
}

if ($Preview)
{
    $actions.Add('Preview only: no targets were changed.')
    $logPath = Join-Path $backupPath 'cleanup.log'
    [System.IO.File]::WriteAllLines($logPath, $actions, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
    $actions
    return
}

foreach ($registryValueTarget in $registryValueTargets)
{
    try
    {
        Remove-ItemProperty -LiteralPath $registryValueTarget.Path -Name $registryValueTarget.ValueName -Force -ErrorAction Stop
        $actions.Add("Removed registry value: $($registryValueTarget.Path)::$($registryValueTarget.ValueName)")
    }
    catch
    {
        $failures.Add("Registry value removal failed: $($registryValueTarget.Path)::$($registryValueTarget.ValueName) - $($_.Exception.Message)")
    }
}

foreach ($registryTarget in $registryTargets)
{
    try
    {
        Remove-Item -LiteralPath $registryTarget -Recurse -Force -ErrorAction Stop
        $actions.Add("Removed registry: $registryTarget")
    }
    catch
    {
        $failures.Add("Registry removal failed: $registryTarget - $($_.Exception.Message)")
    }
}

foreach ($manifestTarget in $externalManifestPaths)
{
    try
    {
        if (Test-Path -LiteralPath $manifestTarget -PathType Leaf)
        {
            [System.IO.File]::Delete($manifestTarget)
            $actions.Add("Removed file: $manifestTarget")
        }
    }
    catch
    {
        $failures.Add("File removal failed: $manifestTarget - $($_.Exception.Message)")
    }
}

foreach ($directoryTarget in $externalDirectories)
{
    try
    {
        $directoryInfo = Get-Item -LiteralPath $directoryTarget -Force -ErrorAction Stop
        if (($directoryInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        {
            $failures.Add("Skipped reparse-point directory: $directoryTarget")
            continue
        }
        [System.IO.Directory]::Delete($directoryTarget, $true)
        $actions.Add("Removed directory: $directoryTarget")
    }
    catch
    {
        $failures.Add("Directory removal failed: $directoryTarget - $($_.Exception.Message)")
    }
}

foreach ($cacheTarget in $clickOnceBranches)
{
    try
    {
        $cacheInfo = Get-Item -LiteralPath $cacheTarget -Force -ErrorAction Stop
        if (($cacheInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        {
            $failures.Add("Skipped reparse-point ClickOnce branch: $cacheTarget")
            continue
        }
        [System.IO.Directory]::Delete($cacheTarget, $true)
        $actions.Add("Removed ClickOnce branch: $cacheTarget")
    }
    catch
    {
        $failures.Add("ClickOnce removal failed: $cacheTarget - $($_.Exception.Message)")
    }
}

$logPath = Join-Path $backupPath 'cleanup.log'
foreach ($failure in $failures) { $actions.Add($failure) }
[System.IO.File]::WriteAllLines($logPath, $actions, (New-Object System.Text.UTF8Encoding -ArgumentList $false))

$actions
if ($failures.Count -gt 0)
{
    throw "SlideSCI Focus 清理未完全成功；详情见 $logPath"
}

[pscustomobject]@{
    BackupDirectory = $backupPath
    RegistryRemoved = $registryTargets.Count
    ExternalFilesRemoved = $externalManifestPaths.Count
    ExternalDirectoriesRemoved = $externalDirectories.Count
    ClickOnceBranchesRemoved = $clickOnceBranches.Count
    Log = $logPath
}
