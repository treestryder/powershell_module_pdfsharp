<#
.Synopsis
    Uses the PDFSharp framework to read and return a PDF Document object.
#>
function Get-Pdf {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [object[]]$Path
    )

    process {
        foreach ($p in $Path) {
            if ($p -is [System.IO.DirectoryInfo]) { continue }
            if ($p -is [System.IO.FileInfo]) { $p = $p.FullName }
            if ([PdfSharp.Pdf.IO.PdfReader]::TestPdfFile($p) -eq 0) {    #(-not (Test-Path $p)) {
                Write-Warning "File not found or not a valid PDF: $p"
                continue
            }
            Write-Verbose $p
            [PdfSharp.Pdf.IO.PdfReader]::Open($p);
        }
    }
}
