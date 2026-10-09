# Command-line build for EBHMaterialPlanning (Visual Studio 2019 MSBuild).
#   .\build.ps1   builds the add-on (Release | x64)
param([string]$Configuration = "Release")

$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
& $msbuild "$PSScriptRoot\EBHMaterialPlanning.sln" "/p:Configuration=$Configuration" "/p:Platform=x64" "/v:m" "/nologo"
if ($LASTEXITCODE -ne 0) { Write-Host "BUILD FAILED"; exit 1 }
Write-Host "BUILD OK"