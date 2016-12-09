<#
.Synopsis
    Uses the PDFSharp framework to set properties on a PDF document.
#>
function Set-PdfProperties {
    [CmdletBinding()]
        param (
        [Parameter(Mandatory=$true, ParameterSetName='File')]
        [string]$Path,
        [Parameter(ParameterSetName='File')]
		[string]$OutPath,
        [Parameter(Mandatory=$true, ParameterSetName='PdfDocument')]
        [PdfSharp.Pdf.PdfDocument]$PdfDocument,
        [string]$Title,
        [string]$Author,
        [string]$Subject,
        [object]$CreationDate,
		[string]$UserPassword,
		[string]$OwnerPassword,
        [hashtable]$CustomProperties,
        [switch]$Passthru
    )

	if ([string]::IsNullOrWhiteSpace($OutPath)) {
		$OutPath = $Path
	}
    Write-Verbose "Setting properties for the PDF $Path"
    if ($PdfDocument -eq $null) { 
        $PdfDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($Path)
    }
    else {
        $Passthru = $true
    }
    if ([string]::IsNullOrWhiteSpace($Title) -eq $false) { 
        $PdfDocument.Info.Title = $Title
    }
    if ([string]::IsNullOrWhiteSpace($Author) -eq $false) {
        $PdfDocument.Info.Author = $Author
    }
    if ([string]::IsNullOrWhiteSpace($Subject) -eq $false) {
        $PdfDocument.Info.Subject = $Subject
    }
    $d = Get-Date
    if ( [DateTime]::TryParse($CreationDate, [ref] $d) ) {
        $PdfDocument.Info.CreationDate = $d
    }
    if ([string]::IsNullOrWhiteSpace($UserPassword) -eq $false) {
        $PdfDocument.SecuritySettings.UserPassword = $UserPassword
    }
	if ([string]::IsNullOrWhiteSpace($OwnerPassword) -eq $false) {
        $PdfDocument.SecuritySettings.OwnerPassword = $OwnerPassword
    }

    # Add custom properties to PDF.
    if ($CustomProperties -ne $null) {
        foreach ($key in $CustomProperties.Keys) {
            $pdfString = New-Object -TypeName PdfSharp.Pdf.PdfString -Argument $CustomProperties[$key]
            $PdfDocument.Info.Elements.Add('/' + $key, $pdfString)
        }
    }
    if ($Passthru) {
        Write-Output $PdfDocument
    }
    else {
        $PdfDocument.Save($OutPath)
    }
    $PdfDocument.Close();
    $PdfDocument.Dispose()
}
