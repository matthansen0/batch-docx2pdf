##########################################################
####           Batch Convert .docx to PDF             ####
####                                                  ####
####   https://github.com/matthansen0/batch-docx2pdf  ####
##########################################################


$path = Read-Host -Prompt "Please enter the file system path to your docx files."
$msWord = New-Object -ComObject Word.Application

Get-ChildItem -Path $path -Filter *.doc? -ErrorAction Stop | ForEach-Object {
        $doc = $msWord.Documents.Open($_.FullName)
        $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
        $doc.SaveAs([ref] $pdf_filename, [ref] 17)
        $doc.Close() 
    }
$msWord.Quit()

Write-Host "Conversion is complete. PDF files have been saved to $path." -ForegroundColor Green