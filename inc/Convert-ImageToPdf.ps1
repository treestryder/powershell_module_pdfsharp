<#
.Synopsis
    Uses the PDFSharp framework to produce a PDF file from one or more image files.
.Notes
    The method [PdfSharp.Drawing.XImage]::FromStream() will not run in the Powershell ISE.
#>
function Convert-ImageToPdf {
    [CmdletBinding()]
    [OutputType()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [object[]]$Path,
        [Parameter(Mandatory=$true)]
        [object]$Destination,
        [string]$Title,
        [string]$Author,
        [string]$Subject,
        [object]$CreationDate,
		[string]$UserPassword,
		[string]$OwnerPassword,
        [hashtable]$CustomProperties,
        [switch]$Force
    )

    begin {
        Add-Type -AssemblyName System.Drawing
        $PdfDocument = $null
        if ($Destination -is [System.IO.FileInfo]) { $Destination = $Destination.FullName }
        if ((Test-Path $Destination) -and -not $Force) {
            Write-Verbose "Appending to PDF $Destination"
            $PdfDocument = Get-Pdf -Path $Destination
        }
        else {
            Write-Verbose "Creating PDF $Destination"
            $PdfDocument = New-Object PdfSharp.Pdf.PdfDocument
        }
        $PdfDocument = Set-PdfProperty -PdfDocument $PdfDocument -Title:$Title -Author:$Author -Subject:$Subject -CreationDate:$CreationDate -UserPassword:$UserPassword -OwnerPassword:$OwnerPassword -CustomProperties:$CustomProperties
    }

    process {
        foreach ($image in $Path) {
            if ($image -is [System.IO.DirectoryInfo]) { continue }
            if ($image -is [System.IO.FileInfo]) { $image = $image.FullName }
            Write-Debug $image
            if (-not (Test-Path $image)) {
                Write-Warning "Image file not found: $image"
                continue
            }
            
            $ximage = $null
            $original = $null
            $xgraphics = $null
            $ximage = $null
            $ms = $null
            try {
                $original = New-Object System.Drawing.Bitmap -ArgumentList $image -ErrorAction Stop
                $frameCount = $original.GetFrameCount([System.Drawing.Imaging.FrameDimension]::Page)
                for ($i = 0; $i -lt $frameCount; $i++) {
                    Write-Verbose "    Adding image: $image frame: $i"
                    $null = $original.SelectActiveFrame([System.Drawing.Imaging.FrameDimension]::Page, $i)
                    $ms = New-Object System.IO.MemoryStream
                    $original.Save( $ms, $original.RawFormat)
                    $ximage = [PdfSharp.Drawing.XImage]::FromStream($ms)
                    $xgraphics  = [PdfSharp.Drawing.XGraphics]::FromPdfPage($PdfDocument.Pages.Add())
                    $xgraphics.DrawImage($ximage,0,0)
                    $xgraphics.Dispose()
                    $ximage.Dispose()
                    $ms.Dispose()
                }
                $original.Dispose()
            }
            catch {
                $pdfDocument.Close()
                $PdfDocument.Dispose()
                throw ('Failed to add image {0}: {1}' -f $image, $_.ToString())
            }
            finally {
                if ($xgraphics -ne $null) { $xgraphics.Dispose() }
                if ($ximage -ne $null) { $ximage.Dispose() }
                if ($ms -ne $null) { $ms.Dispose() }
                if ($original -ne $null) { $original.Dispose() }
            }
        }
    }

    end {
        if ($PdfDocument.PageCount -gt 0) {
            Write-Verbose "    Saving PDF"
            $pdfDocument.Save($Destination)
            $pdfDocument.Close()
        } else {
            Write-Warning "PDF $Destination was not created as there were no valid pages."
        }

        $PdfDocument.Dispose()
    }
}
