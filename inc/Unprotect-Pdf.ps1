function Unprotect-PDF {
    <#
     .Synopsis
     Removes the password from a PDF file.
    
     .Parameter Path
     Full path or array of PDF paths to decrypt. Excepts piped input. Required.
    
     .Example
     Unprotect-PDF EncryptedFile.pdf
    
     .Example
    Unprotect-PDF -Path encrypted.pdf
     
     .Example
     dir *.pdf | Unprotect-PDF -Password password -Verbose
    
    #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true,
                HelpMessage='One or more PDF files to process')]
            [ValidateNotNullorEmpty()]
            [object[]]$Path,
            [Parameter(Mandatory=$true,
                HelpMessage='The password to decrypt the PDF files')]
            [ValidateNotNullorEmpty()]
            [string]$Password
        )

        process {
            foreach ($file in $Path) {
                if ($file -is [System.IO.FileInfo]) {$file = $file.FullName}
                if ($file -isnot [string]) { continue }
                Write-Verbose $file
                try {
                    $document = [PdfSharp.Pdf.IO.PdfReader]::Open($file, $password)
                    $document.Save($file)
                }
                catch {
                    Write-Error "Error removing password from $file : $_"
                }
            }
        }
    }
