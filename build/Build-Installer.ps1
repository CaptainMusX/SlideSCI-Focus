[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [string]$OutputDirectory = "",
    [switch]$BundleLatexRuntime
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot ".." )).Path
$solutionPath = Join-Path $repoRoot "SlideSCI-Focus.sln"
$latexDirectory = Join-Path $repoRoot "SlideSCI\latex-converter"
$artifactRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot "artifacts"))
$artifactRootWithSeparator = $artifactRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
$projectPath = Join-Path $repoRoot "SlideSCI\SlideSCI.csproj"
$bundleLatexRuntimeValue = if ($BundleLatexRuntime) { "true" } else { "false" }
$projectXml = New-Object System.Xml.XmlDocument
$projectXml.Load($projectPath)
$versionNode = $projectXml.SelectSingleNode("//*[local-name()='ApplicationVersion']")
$publishVersion = if ($versionNode -and $versionNode.InnerText) { $versionNode.InnerText } else { "local" }

if ([string]::IsNullOrWhiteSpace($OutputDirectory))
{
    $OutputDirectory = Join-Path $artifactRoot "SlideSCI-Focus-$publishVersion"
}

$outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
$outputPathWithoutSeparator = $outputPath.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
$isArtifactRoot = $outputPathWithoutSeparator.Equals($artifactRoot, [System.StringComparison]::OrdinalIgnoreCase)
$isUnderArtifactRoot = $outputPath.StartsWith($artifactRootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)
if ($isArtifactRoot -or -not $isUnderArtifactRoot)
{
    throw "OutputDirectory 必须是 artifacts 目录下的专用输出目录。"
}

if (-not (Test-Path -LiteralPath $solutionPath -PathType Leaf))
{
    throw "找不到解决方案文件: $solutionPath"
}

function Get-RequiredCommandPath
{
    param([string]$CommandName)

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if (-not $command)
    {
        throw "找不到必需的命令: $CommandName"
    }
    return $command.Source
}

function Find-MSBuild
{
    $vswherePath = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vswherePath -PathType Leaf)
    {
        $found = @(& $vswherePath -latest -products '*' -find "MSBuild\Current\Bin\MSBuild.exe" 2>$null)
        if ($found.Count -gt 0 -and (Test-Path -LiteralPath $found[0] -PathType Leaf))
        {
            $installationPath = (& $vswherePath -latest -products '*' -property installationPath 2>$null | Select-Object -First 1)
            $installationVersion = (& $vswherePath -latest -products '*' -property installationVersion 2>$null | Select-Object -First 1)
            $majorVersion = if ($installationVersion -match '^(\d+)') { [int]$Matches[1] } else { 18 }
            $vstoPath = Join-Path $installationPath "MSBuild\Microsoft\VisualStudio\v$majorVersion.0"
            $officeTargets = Join-Path $vstoPath "OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"
            if (Test-Path -LiteralPath $officeTargets -PathType Leaf)
            {
                return [PSCustomObject]@{
                    Path = $found[0]
                    VisualStudioVersion = "$majorVersion.0"
                    VSToolsPath = $vstoPath
                }
            }
        }
    }

    $candidates = @(
        "D:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    )
    foreach ($candidate in $candidates)
    {
        if (Test-Path -LiteralPath $candidate -PathType Leaf)
        {
            $candidateRoot = Split-Path (Split-Path (Split-Path (Split-Path $candidate -Parent) -Parent) -Parent) -Parent
            $version = if ($candidate -match "\\(\d+)\\Community\\") { "$($Matches[1]).0" } else { "17.0" }
            $vstoPath = Join-Path $candidateRoot "MSBuild\Microsoft\VisualStudio\v$version"
            $officeTargets = Join-Path $vstoPath "OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"
            if (Test-Path -LiteralPath $officeTargets -PathType Leaf)
            {
                return [PSCustomObject]@{
                    Path = $candidate
                    VisualStudioVersion = $version
                    VSToolsPath = $vstoPath
                }
            }
        }
    }

    throw "找不到包含 OfficeTools targets 的 Visual Studio MSBuild。"
}

function Find-StrongNameTool
{
    $programFilesX86 = ${env:ProgramFiles(x86)}
    $candidates = @(
        (Join-Path $programFilesX86 "Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\sn.exe"),
        (Join-Path $programFilesX86 "Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\sn.exe")
    )

    foreach ($candidate in $candidates)
    {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }

    throw "找不到 .NET 强名称工具 sn.exe。"
}

function Ensure-LocalSigningMaterial
{
    $signingDirectory = Join-Path $env:LOCALAPPDATA "SlideSCI-Focus\build-signing"
    New-Item -ItemType Directory -Force -Path $signingDirectory | Out-Null

    # Keep the direct-deployment channel on its own stable manifest identity.
    # Reusing the certificate from the earlier ClickOnce wrapper makes VSTO
    # reuse that subscription's original deployment URL (for example D:\...).
    $certificateSubject = "CN=SlideSCI Focus Direct Deployment Signing"
    $certificate = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq $certificateSubject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if (-not $certificate)
    {
        $certificate = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $certificateSubject `
            -FriendlyName "SlideSCI Focus Direct Deployment Signing" `
            -CertStoreLocation Cert:\CurrentUser\My `
            -KeyExportPolicy Exportable `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
            -KeySpec Signature `
            -NotAfter (Get-Date).AddYears(2)
    }

    $pfxPath = Join-Path $signingDirectory "SlideSCI-Focus-Build-Signing.pfx"
    $certificatePath = Join-Path $signingDirectory "SlideSCI-Focus-Build-Signing.cer"
    $certutilPath = Get-RequiredCommandPath "certutil.exe"
    & $certutilPath -exportPFX -user -f -p "" My $certificate.Thumbprint $pfxPath | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $pfxPath -PathType Leaf))
    {
        throw "无法导出本机 ClickOnce 签名证书。"
    }

    Export-Certificate -Cert $certificate -FilePath $certificatePath -Force | Out-Null

    return [PSCustomObject]@{
        PfxPath = $pfxPath
        CertificatePath = $certificatePath
        Thumbprint = $certificate.Thumbprint
        Subject = $certificate.Subject
    }
}

$nodePath = Get-RequiredCommandPath "node.exe"
$npmPath = Get-RequiredCommandPath "npm.cmd"
$msbuild = Find-MSBuild
$snPath = Find-StrongNameTool
$signing = Ensure-LocalSigningMaterial
$snkPath = Join-Path (Split-Path $signing.PfxPath -Parent) "SlideSCI-Focus-Build-Assembly.snk"

if (-not (Test-Path -LiteralPath (Join-Path $latexDirectory "node_modules\mathjax-full") -PathType Container))
{
    Push-Location $latexDirectory
    try
    {
        if (Test-Path -LiteralPath (Join-Path $latexDirectory "package-lock.json") -PathType Leaf)
        {
            & $npmPath ci --ignore-scripts --no-audit --no-fund
        }
        else
        {
            & $npmPath install --ignore-scripts --no-audit --no-fund
        }
        if ($LASTEXITCODE -ne 0) { throw "npm 安装 LaTeX 依赖失败。" }
    }
    finally
    {
        Pop-Location
    }
}

if (-not (Test-Path -LiteralPath $snkPath -PathType Leaf))
{
    & $snPath -k $snkPath | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $snkPath -PathType Leaf))
    {
        throw "无法生成程序集强名称密钥。"
    }
}

if (Test-Path -LiteralPath $outputPath -PathType Container)
{
    [System.IO.Directory]::Delete($outputPath, $true)
}
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
$publishDirectoryWithSlash = $outputPath.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar

$msbuildArguments = @(
    $solutionPath,
    "/t:Publish",
    "/p:Configuration=$Configuration",
    "/p:VisualStudioVersion=$($msbuild.VisualStudioVersion)",
    "/p:VSToolsPath=$($msbuild.VSToolsPath)",
    "/p:PublishDir=$publishDirectoryWithSlash",
    "/p:PublishUrl=$publishDirectoryWithSlash",
    "/p:SlideSCIFocusBundleLatexRuntime=$bundleLatexRuntimeValue",
    "/p:SlideSCIFocusManifestKeyFile=$($signing.PfxPath)",
    "/p:SlideSCIFocusAssemblyKeyFile=$snkPath",
    "/p:SlideSCIFocusManifestCertificateThumbprint=$($signing.Thumbprint)",
    "/p:SignManifests=true",
    "/p:SignAssembly=true",
    "/m",
    "/v:minimal"
)

Push-Location $repoRoot
try
{
    $restoreArguments = @($msbuildArguments)
    $restoreArguments[1] = "/t:Restore"
    $restoreArguments += "/p:RestorePackagesConfig=true"
    & $msbuild.Path @restoreArguments
    if ($LASTEXITCODE -ne 0) { throw "NuGet 包还原失败，MSBuild exit code=$LASTEXITCODE" }

    $cleanArguments = @($msbuildArguments)
    $cleanArguments[1] = "/t:Clean"
    & $msbuild.Path @cleanArguments
    if ($LASTEXITCODE -ne 0) { throw "清理 VSTO 构建输出失败，MSBuild exit code=$LASTEXITCODE" }

    & $msbuild.Path @msbuildArguments
    if ($LASTEXITCODE -ne 0) { throw "VSTO 发布失败，MSBuild exit code=$LASTEXITCODE" }
}
finally
{
    Pop-Location
}

$installerPath = Join-Path $outputPath "setup.exe"
$deploymentManifestPath = Join-Path $outputPath "CaptainMusX.SlideSCI.Focus.vsto"
$applicationFilesDirectory = Join-Path $outputPath "Application Files"
$publishedCertificatePath = Join-Path $outputPath "SlideSCI-Focus-Build-Signing.cer"
$signing.CertificatePath | Copy-Item -Destination $publishedCertificatePath -Force
$hasInstaller = Test-Path -LiteralPath $installerPath -PathType Leaf
$hasDeploymentManifest = Test-Path -LiteralPath $deploymentManifestPath -PathType Leaf
$hasApplicationFiles = Test-Path -LiteralPath $applicationFilesDirectory -PathType Container
if (-not $hasInstaller -or -not $hasDeploymentManifest -or -not $hasApplicationFiles)
{
    throw "发布完成但未找到完整的 setup.exe / CaptainMusX.SlideSCI.Focus.vsto / Application Files。"
}

$publishedApplicationDirectory = Get-ChildItem -LiteralPath $applicationFilesDirectory -Directory | Select-Object -First 1
if (-not $publishedApplicationDirectory)
{
    throw "发布目录缺少 Application Files 子目录。"
}

$publishedMathJaxEntry = Join-Path $publishedApplicationDirectory.FullName "latex-converter\node_modules\mathjax-full\js\mathjax.js.deploy"
if ($BundleLatexRuntime -and -not (Test-Path -LiteralPath $publishedMathJaxEntry -PathType Leaf))
{
    throw "已要求捆绑 LaTeX 运行时，但发布目录缺少 MathJax 文件。"
}

$deploymentText = Get-Content -Raw -LiteralPath $deploymentManifestPath
if ($deploymentText -notmatch "<Signature" -or $deploymentText -notmatch "publisherIdentity")
{
    throw "部署清单未检测到签名信息。"
}

$checksumPath = Join-Path $outputPath "SHA256SUMS.txt"
$checksumLines = Get-ChildItem -LiteralPath $outputPath -Recurse -File |
    Where-Object { $_.FullName -ne $checksumPath } |
    Sort-Object FullName |
    ForEach-Object {
        $relative = $_.FullName.Substring($outputPath.Length).TrimStart('\', '/')
        $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        "$hash  $relative"
    }
[System.IO.File]::WriteAllLines($checksumPath, $checksumLines, (New-Object System.Text.UTF8Encoding -ArgumentList $false))

[PSCustomObject]@{
    Version = $publishVersion
    Configuration = $Configuration
    OutputDirectory = $outputPath
    Installer = $installerPath
    DeploymentManifest = $deploymentManifestPath
    Certificate = $signing.CertificatePath
    CertificateSubject = $signing.Subject
    CertificateThumbprint = $signing.Thumbprint
    BundleLatexRuntime = $BundleLatexRuntime
    Node = $nodePath
    MSBuild = $msbuild.Path
}
