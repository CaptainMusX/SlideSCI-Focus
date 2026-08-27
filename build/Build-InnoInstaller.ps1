[CmdletBinding()]
param(
    [string]$SourceDirectory = "",
    [string]$OutputDirectory = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot ".." )).Path
$issPath = Join-Path $repoRoot "build\SlideSCI.iss"
$artifactRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot "artifacts"))
$projectPath = Join-Path $repoRoot "SlideSCI\SlideSCI.csproj"

function Find-InnoCompiler
{
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $candidates = @(
        "D:\Program Files\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        "D:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    )
    foreach ($candidate in $candidates)
    {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }

    throw "找不到 Inno Setup 6 的 ISCC.exe。"
}

function Find-SignTool
{
    $command = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    $kitRoot = Join-Path $programFilesX86 'Windows Kits\10\bin'
    if (Test-Path -LiteralPath $kitRoot -PathType Container)
    {
        $candidate = Get-ChildItem -LiteralPath $kitRoot -Recurse -Filter signtool.exe -File -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($candidate) { return $candidate.FullName }
    }

    return $null
}

function Test-UnderRoot
{
    param(
        [Parameter(Mandatory = $true)][string]$Candidate,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $candidateFull = [System.IO.Path]::GetFullPath($Candidate).TrimEnd('\', '/')
    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    return $candidateFull.StartsWith($rootFull + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

if (-not (Test-Path -LiteralPath $issPath -PathType Leaf))
{
    throw "找不到 Inno Setup 脚本: $issPath"
}
if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf))
{
    throw "找不到项目文件: $projectPath"
}

$projectXml = New-Object System.Xml.XmlDocument
$projectXml.Load($projectPath)
$versionNode = $projectXml.SelectSingleNode("//*[local-name()='ApplicationVersion']")
$appVersion = if ($versionNode -and $versionNode.InnerText) { $versionNode.InnerText } else { 'local' }

if ([string]::IsNullOrWhiteSpace($SourceDirectory))
{
    $SourceDirectory = Join-Path $repoRoot 'SlideSCI\bin\Release'
}
$releasePath = [System.IO.Path]::GetFullPath($SourceDirectory)
if (-not (Test-Path -LiteralPath $releasePath -PathType Container))
{
    throw "Release 输出目录不存在: $releasePath；请先运行 build\Build-Installer.ps1。"
}

$stagingPath = [System.IO.Path]::GetFullPath((Join-Path $artifactRoot ("inno-source-{0}" -f $appVersion)))
if (-not (Test-UnderRoot -Candidate $stagingPath -Root $artifactRoot))
{
    throw "内部 staging 目录越界: $stagingPath"
}
if (Test-Path -LiteralPath $stagingPath -PathType Container)
{
    [System.IO.Directory]::Delete($stagingPath, $true)
}
New-Item -ItemType Directory -Force -Path $stagingPath | Out-Null

$sourceFiles = @(Get-ChildItem -LiteralPath $releasePath -Recurse -File -ErrorAction Stop |
    Where-Object {
        if ($_.FullName -match '\\app\.publish(\\|$)' -or
            $_.FullName -match '\\latex-converter\\node_modules(\\|$)' -or
            $_.Extension -ieq '.pdb')
        {
            return $false
        }

        # NuGet XML documentation files are not read at runtime and account
        # for most of the uncompressed install footprint.
        if ($_.Extension -ieq '.xml' -and
            (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($_.FullName, '.dll')) -PathType Leaf))
        {
            return $false
        }

        return $true
    })
foreach ($sourceFile in $sourceFiles)
{
    $relativePath = $sourceFile.FullName.Substring($releasePath.Length).TrimStart('\', '/')
    $destination = Join-Path $stagingPath $relativePath
    $destinationDirectory = [System.IO.Path]::GetDirectoryName($destination)
    New-Item -ItemType Directory -Force -Path $destinationDirectory | Out-Null
    Copy-Item -LiteralPath $sourceFile.FullName -Destination $destination -Force
}

$certificatePath = Join-Path $env:LOCALAPPDATA 'SlideSCI\build-signing\SlideSCI-Build-Signing.cer'
if (Test-Path -LiteralPath $certificatePath -PathType Leaf)
{
    Copy-Item -LiteralPath $certificatePath -Destination (Join-Path $stagingPath 'SlideSCI-Build-Signing.cer') -Force
}

foreach ($requiredName in @('CaptainMusX.SlideSCI.vsto', 'CaptainMusX.SlideSCI.dll.manifest', 'CaptainMusX.SlideSCI.dll', 'latex-converter\latex-to-svg.js', 'latex-converter\package.json'))
{
    $requiredPath = Join-Path $stagingPath $requiredName
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf))
    {
        throw "直接 VSTO staging 缺少必需项: $requiredPath"
    }
}

$applicationManifestPath = Join-Path $stagingPath 'CaptainMusX.SlideSCI.dll.manifest'
[xml]$applicationManifest = Get-Content -LiteralPath $applicationManifestPath -Raw
$rsaKeyNode = $applicationManifest.SelectSingleNode('//*[local-name()="RSAKeyValue"]')
if (-not $rsaKeyNode)
{
    throw "应用清单缺少 VSTO Inclusion 所需的 RSA 公钥。"
}
$modulusNode = $rsaKeyNode.SelectSingleNode('./*[local-name()="Modulus"]')
$exponentNode = $rsaKeyNode.SelectSingleNode('./*[local-name()="Exponent"]')
if (-not $modulusNode -or -not $exponentNode)
{
    throw "应用清单中的 RSA 公钥不完整。"
}
$vstoPublicKey = '<RSAKeyValue><Modulus>' + $modulusNode.InnerText +
    '</Modulus><Exponent>' + $exponentNode.InnerText + '</Exponent></RSAKeyValue>'
$trustIncludePath = Join-Path $artifactRoot 'VstoTrust.generated.iss'
[System.IO.File]::WriteAllText(
    $trustIncludePath,
    ('#define VstoPublicKey "' + $vstoPublicKey + '"'),
    (New-Object System.Text.UTF8Encoding -ArgumentList $false)
)
$manifestReferences = @()
foreach ($node in $applicationManifest.SelectNodes('//*[local-name()="dependentAssembly"][@dependencyType="install"]'))
{
    $manifestReferences += [string]$node.codebase
}
foreach ($node in $applicationManifest.SelectNodes('//*[local-name()="file"]'))
{
    $manifestReferences += [string]$node.name
}
$missingReferences = @()
foreach ($reference in $manifestReferences | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
{
    $relativeReference = $reference.Replace('/', '\')
    if ([System.IO.Path]::IsPathRooted($relativeReference)) { throw "清单包含绝对路径: $reference" }
    $resolvedReference = Join-Path $stagingPath $relativeReference
    if (-not (Test-Path -LiteralPath $resolvedReference -PathType Leaf)) { $missingReferences += $reference }
}
if ($missingReferences.Count -gt 0)
{
    throw "直接 VSTO staging 有缺失的清单文件: $($missingReferences -join ', ')"
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory))
{
    $OutputDirectory = Join-Path $artifactRoot 'inno'
}
$outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
if (-not (Test-UnderRoot -Candidate $outputPath -Root $artifactRoot))
{
    throw "OutputDirectory 必须位于 artifacts 目录下: $outputPath"
}
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

$isccPath = Find-InnoCompiler
$signToolPath = Find-SignTool
$pfxPath = Join-Path $env:LOCALAPPDATA 'SlideSCI\build-signing\SlideSCI-Build-Signing.pfx'
$logPath = Join-Path $outputPath 'inno-build.log'
$setupPath = Join-Path $outputPath ("SlideSCI-{0}-Setup.exe" -f $appVersion)

$isccArguments = @(
    "/DAppVersion=$appVersion",
    "/DSourceRoot=$stagingPath",
    "/O$outputPath",
    $issPath
)

Push-Location $repoRoot
try
{
    & $isccPath @isccArguments 2>&1 | Tee-Object -FilePath $logPath
    if ($LASTEXITCODE -ne 0) { throw "Inno Setup 编译失败，exit code=$LASTEXITCODE" }
}
finally
{
    Pop-Location
}

if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf))
{
    throw "Inno Setup 编译完成但找不到安装包: $setupPath"
}

$signatureResult = 'not-signed'
if ($signToolPath -and (Test-Path -LiteralPath $pfxPath -PathType Leaf))
{
    & $signToolPath sign /fd SHA256 /f $pfxPath /p '' $setupPath 2>&1 |
        Tee-Object -FilePath (Join-Path $outputPath 'inno-sign.log')
    if ($LASTEXITCODE -ne 0) { throw "安装包签名失败，exit code=$LASTEXITCODE" }
    $signatureResult = 'signed-with-local-build-certificate'
}

$hash = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash
$size = (Get-Item -LiteralPath $setupPath).Length
[pscustomobject]@{
    ReleaseDirectory = $releasePath
    StagingDirectory = $stagingPath
    Installer = $setupPath
    Bytes = $size
    StagedFiles = $sourceFiles.Count
    SHA256 = $hash
    SignTool = $signToolPath
    Signature = $signatureResult
    Log = $logPath
}
