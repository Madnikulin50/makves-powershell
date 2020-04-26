
##by https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer
param(
    [string]$template = "C:\work\makves\makves-powershell\Reports\template.docx",
    [string]$outFilePath = ".\report.docx",
    [string]$find_String = "{/CN/}" 
)
##Make new Object MS word
[psobject]$word = New-Object -ComObject Word.Application
$word.Visible = $true
[psobject]$doc = $word.Documents.Open($template)

function findAndReplaceWholeDoc {
    param (
        [psobject]$document,
        [string]$find = "",
        [string]$replaceWith
    )
   

    $document.Paragraphs | ForEach-Object {
        $rng = $_.Range;
        [psobject]$findReplace = $rng.Find
        $findReplace.ClearFormatting();
        $res = $findReplace.Execute($find,  $True, $false, $null, $null, $null, $null, $null, $null, $replaceWith,
        2, $null, $null, $null, $null)
       
        Write-Host "Find in paragraph" $res
    }

    $document.Shapes | ForEach-Object {
        $rng = $_.TextFrame.TextRange;
        [psobject]$findReplace = $rng.Find
        $findReplace.ClearFormatting();
        $res = $findReplace.Execute($find,  $True, $false, $null, $null, $null, $null, $null, $null, $replaceWith,
        2, $null, $null, $null, $null)
       
        Write-Host "Find in shapes" $res
    }


   
    #$findReplace.Execute($find,  $True, $false, $null, $null, $null, $null, $null, $null, $replaceWith,
    #2, $null, $null, $null, $null)
    

}
##FindAndReplaceWholeDoc -Document $doc -Find $find_String -ReplaceWith "Ivan"
#FindAndReplaceWholeDoc -Document $doc -Find "{/TD/}" -ReplaceWith "morning"
FindAndReplaceWholeDoc -Document $doc -Find "USRSUM" -ReplaceWith "9440"
##$doc.Save("report.doc")
#$doc.shapes