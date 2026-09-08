param()
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$output=Join-Path $root 'artifacts/title-smoke'
New-Item -ItemType Directory -Force $output | Out-Null
$assembly=Join-Path $root 'SlideSCI/bin/Release/CaptainMusX.SlideSCI.Focus.dll'
Copy-Item $assembly $output -Force
$ppt='C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll'
$office='C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'
$csc=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
& $csc /nologo /target:exe "/out:$output/TitleSmoke.exe" "/reference:$assembly" "/link:$ppt" "/link:$office" /reference:C:\Windows\assembly\GAC\stdole\7.0.3300.0__b03f5f7f11d50a3a\stdole.dll /reference:System.Drawing.dll /reference:System.Configuration.dll (Join-Path $PSScriptRoot 'TitleSmoke.cs')
if($LASTEXITCODE -ne 0){throw 'Compilation failed'}
& "$output/TitleSmoke.exe" $output
if($LASTEXITCODE -ne 0){throw 'PowerPoint title smoke failed'}
