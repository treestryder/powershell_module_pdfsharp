<#
.Synopsis
    Uses the PDFSharp framework to produce a PDF file from one or more image files.
#>
function Convert-ImageToPdf {
    [CmdletBinding()]
    [OutputType()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string[]]$ImagePath,
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string]$Title,
        [string]$Author,
        [string]$Subject,
        [object]$CreationDate,
		[string]$UserPassword,
		[string]$OwnerPassword,
        [hashtable]$CustomProperties
    )

    begin {
        Write-Verbose "Creating PDF $Path"
        $PdfDocument = New-Object PdfSharp.Pdf.PdfDocument
        $PdfDocument = Set-PdfProperties -PdfDocument $PdfDocument -Title:$Title -Author:$Author -Subject:$Subject -CreationDate:$CreationDate -UserPassword:$UserPassword -OwnerPassword:$OwnerPassword -CustomProperties:$CustomProperties
    }

    process {
        foreach ($image in $ImagePath) {
            if (-not (Test-Path $image)) {
                Write-Warning "Image file not found $image"
                continue
            }
            Write-Verbose "    Adding image $image"
            $ximage = [PdfSharp.Drawing.XImage]::FromFile($image)
            $page = $PdfDocument.Pages.Add()
            $xgraphics  = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
            $xgraphics.DrawImage($ximage,0,0)
            $ximage.Dispose()
        }
    }

    end {
        if ($PdfDocument.PageCount -gt 0) {
            Write-Verbose "    Saving PDF"
            $pdfDocument.Save($Path)
            $pdfDocument.Close()
        } else {
            Write-Warning "    PDF $Path was not created as there no valid pages."
        }

        $PdfDocument.Dispose()
    }
}
