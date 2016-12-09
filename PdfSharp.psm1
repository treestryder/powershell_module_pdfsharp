Set-StrictMode -Version Latest

# Fail at the first exception.
trap {}

Add-Type -Path "$PSScriptRoot/inc/pdfsharp.dll"
Get-ChildItem "$PSScriptRoot/inc/*.ps1" | foreach { . $_ }