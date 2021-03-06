# define output directory
$outputDirectory = '\\ixion\las_m\Archiv\2D\ElectronicLabbooks'
$outputDirectoryAllPages = "\\ixion\las_m\Archiv\2D\ElectronicLabbooks\Full_$(Get-Date -format yyyy)"
$outputFile = "$($outputDirectory)\$(Get-Date -format yyyyMMdd)_Summary.pdf"

# enable onenote provider
Enable-Onenote

# get list of all pages which where last modified during the last week
$allPages = dir "OneNote:\Lithium6_Labbook\Labbook $(Get-Date -format yyyy)" -recurse | where-object { (!$_.PSIsContainer) }

$lastfile = gci $outputDirectory | ?{-not $_.PsIsContainer} | sort LastWriteTime | select -last 1
$lastDate = $lastfile.Name.Substring(0,8)

$changedPages = dir "OneNote:\Lithium6_Labbook\Labbook $(Get-Date -format yyyy)" -recurse | where-object {($_.Name -gt $lastDate) -and ($_.Name.Length -eq 8) -and (!$_.PSIsContainer) }

# clean up full labbook folder in order to circumvent "overwrite?" question
New-Item -ItemType Directory -Force -Path $outputDirectoryAllPages
Remove-Item "$outputDirectoryAllPages\*.pdf" | Where { ! $_.PSIsContainer }
# export those pages to the archive
$exportedFileNames = $changedPages | export-onenote -output $outputDirectory -format pdf

$allPages | export-onenote -output $outputDirectoryAllPages -format pdf

# add pdfsharp library
Add-Type -Path '.\PdfSharp.dll' 
# merge all resulting pdf files
$output = New-Object PdfSharp.Pdf.PdfDocument            
$PdfReader = [PdfSharp.Pdf.IO.PdfReader]            
$PdfDocumentOpenMode = [PdfSharp.Pdf.IO.PdfDocumentOpenMode]                        
        
foreach($i in $exportedFileNames) {            
    $input = New-Object PdfSharp.Pdf.PdfDocument            
    $input = $PdfReader::Open($i."ExportedFile", $PdfDocumentOpenMode::Import)            
    $input.Pages | %{$output.AddPage($_)}            
}                        
            
           
$output.Save($outputFile) 

# remove single files
foreach($i in $exportedFileNames) {            
    Remove-Item $i."ExportedFile"         
} 


# print pages to minolta, staple and punch
& 'C:\Program Files\SumatraPDF\SumatraPDF.exe' "-print-to-default" $outputFile 