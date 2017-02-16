<#
.Synopsis
    Uses the PDFSharp framework to set properties on a PDF document.
#>
function Set-PdfProperty {
    [CmdletBinding()]
        param (
        [Parameter(
            ParameterSetName='PipedFile',
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [object]$InputObject,
        [Parameter(
            ParameterSetName='File',
            Mandatory=$true,
            Position=0)]
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

    process {
        if ($InputObject -ne $null) {
            if ($InputObject -is [System.IO.FileInfo]) {
                $Path = $InputObject.FullName
                $PdfDocument = $null
                $OutPath = $Path
            }
            elseif ($InputObject -is [string]) {
                $Path = $InputObject
                $PdfDocument = $null
                $OutPath = $Path
            }
            elseif ($InputObject -is [System.IO.DirectoryInfo]) {
                continue
            }
            else {
                Write-Verbose "Skipping invalid input object type: $($InputObject.GetType())"
                continue
            }
        }
        if ([string]::IsNullOrWhiteSpace($OutPath)) {
            $OutPath = $Path
        }
        Write-Verbose "Setting PDF properties for $Path"
        if ($PdfDocument -eq $null) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($OwnerPassword)) {
                    $PdfDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($Path, $OwnerPassword)
                }
                elseif (-not [string]::IsNullOrWhiteSpace($UserPassword)) {
                    $PdfDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($Path, $UserPassword)
                }
                else {
                    $PdfDocument = [PdfSharp.Pdf.IO.PdfReader]::Open($Path)
                }
            }
            catch {
                throw "Error opening PDF $Path : $($_.ToString())"
            }
        }
        else {
            $Passthru = $true
        }

        try {
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
        }
        catch {
            throw "Error setting properties on $Path : $($_.ToString())"
        }
        if ($Passthru) {
            Write-Output $PdfDocument
        }
        else {
            try {
                $PdfDocument.Save($OutPath)
            }
            catch {
                throw "Error setting properties on $Path : $($_.ToString())"
            }
        }
        $PdfDocument.Close();
        $PdfDocument.Dispose()
    }
}
