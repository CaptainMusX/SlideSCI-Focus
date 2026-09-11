[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$assembly = Join-Path $repoRoot 'SlideSCI\bin\Release\CaptainMusX.SlideSCI.Focus.dll'
if (-not (Test-Path -LiteralPath $assembly)) { throw 'Build Release first using build/Build-Installer.ps1.' }
$output = Join-Path $repoRoot ('artifacts\core-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output | Out-Null
Copy-Item -LiteralPath $assembly -Destination $output
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$executable = Join-Path $output 'CoreRegression.exe'
# Exercise the actual deployment configuration as the host configuration.
# Loading only the DLL hides configuration-system failures inside VSTO.
Copy-Item -LiteralPath ($assembly + '.config') -Destination ($executable + '.config')
& $compiler /nologo /target:exe "/out:$executable" "/reference:$assembly" /reference:System.Configuration.dll (Join-Path $PSScriptRoot 'CoreRegression.cs')
if ($LASTEXITCODE -ne 0) { throw 'Regression harness compilation failed.' }
& $executable
if ($LASTEXITCODE -ne 0) { throw 'Regression tests failed.' }
