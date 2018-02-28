Set-StrictMode -Version Latest

# Fail at the first exception.
trap {}

Add-Type -Path "$PSScriptRoot/PdfSharp/pdfsharp.dll"
Get-ChildItem "$PSScriptRoot/inc/*.ps1" | ForEach-Object { . $_ }
