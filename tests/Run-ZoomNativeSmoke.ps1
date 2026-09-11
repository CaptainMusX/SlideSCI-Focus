$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$output=Join-Path $root 'artifacts/zoom-smoke'
New-Item -ItemType Directory -Force $output | Out-Null
$assembly=Join-Path $root 'SlideSCI/bin/Release/CaptainMusX.SlideSCI.Focus.dll'
Copy-Item (Join-Path (Split-Path $assembly) '*.dll') $output -Force
$ppt='C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll'
$office='C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'
$csc=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
& $csc /nologo /target:exe "/out:$output/ZoomNativeSmoke.exe" "/reference:$assembly" "/link:$ppt" "/link:$office" /reference:System.Drawing.dll (Join-Path $PSScriptRoot 'ZoomNativeSmoke.cs')
if($LASTEXITCODE -ne 0){throw 'Compilation failed'}
& "$output/ZoomNativeSmoke.exe" $output
if($LASTEXITCODE -ne 0){throw 'PowerPoint smoke failed'}
