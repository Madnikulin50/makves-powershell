param(
    [string]$template = "https://github.com/Madnikulin50/makves-powershell/blob/develop/Reports/template.docx?raw=true",
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [bool]$makeDOCX = $false,
    [bool]$sendToEmail = $false
    )

## Init web server 
$pair = "${makves_user}:${makves_pwd}" ##Make a string for auth pair login/pass
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)
$basicAuthValue = "Basic $base64"
[psobject]$pairPassLog = @{ Authorization = $basicAuthValue} ##make a headers for invok-webrequest auth

## Make string from date for filenames   
$dateString = ((Get-date).ToString('yyMMddhhmmss'))

## Make tmp file
New-Item -Path "$PSScriptRoot\tmp" -ItemType Directory
[string]$tmpFile = "$PSScriptRoot\tmp\tmp_$dateString.docx" 
[psobject]$WebClient = New-Object System.Net.WebClient
$WebClient.DownloadFile($template, $tmpFile)


##Make new Object MS word
[psobject]$word = New-Object -ComObject Word.Application
$word.Visible = $true
[psobject]$doc = $word.Documents.Open($tmpFile)

##Get stats from makves server
[psobject]$ldapStat = mkvsrequest -typeReport "/ldap/stat" -headers $pairPassLog
[psobject]$ldapExplore = mkvsrequest -typeReport "/ldap/explore" -headers $pairPassLog
[psobject]$usersStat = mkvsrequest -typeReport "/ldap/stat?type=users" -headers $pairPassLog
[psobject]$compsStat = mkvsrequest -typeReport "/ldap/stat?type=computers" -headers $pairPassLog
[psobject]$eventsStat = mkvsrequest -typeReport "/events/stat?total=true&stat=true" -headers $pairPassLog
[psobject]$filesStat = mkvsrequest -typeReport "/file/stat?total=true&duplicates=true&compliance=true&stolled=true&byTime=true&byComputer=true&byType=true&byCompliance=true" -headers $pairPassLog
##Replace in template
findAndReplaceWholeDoc -document $doc -find "ID0001" -replaceWith $usersStat.total
findAndReplaceWholeDoc -document $doc -find "ID0002" -replaceWith $compsStat.total
findAndReplaceWholeDoc -document $doc -find "ID0003" -replaceWith $eventsStat.total
findAndReplaceWholeDoc -document $doc -find "ID0004" -replaceWith $filesStat.totalFiles
findAndReplaceWholeDoc -document $doc -find "ID0005" -replaceWith $usersStat.totalGroups
findAndReplaceWholeDoc -document $doc -find "ID0006" -replaceWith $usersStat.totalDisabled
findAndReplaceWholeDoc -document $doc -find "ID0007" -replaceWith $usersStat.totalInactive
findAndReplaceWholeDoc -document $doc -find "ID0008" -replaceWith $usersStat.totalPasswordExpired
findAndReplaceWholeDoc -document $doc -find "ID0009" -replaceWith $usersStat.totalDontExpirePassword
findAndReplaceWholeDoc -document $doc -find "ID0010" -replaceWith $usersStat.totalLockout
findAndReplaceWholeDoc -document $doc -find "ID0011" -replaceWith $usersStat.totalEmptyGroups
findAndReplaceWholeDoc -document $doc -find "ID0012" -replaceWith $usersStat.totalMnsLogonAccount
findAndReplaceWholeDoc -document $doc -find "ID0013" -replaceWith $compsStat.totalDisabled
findAndReplaceWholeDoc -document $doc -find "ID0014" -replaceWith $compsStat.totalInactive
findAndReplaceWholeDoc -document $doc -find "ID0015" -replaceWith $filesStat.totalFolders
findAndReplaceWholeDoc -document $doc -find "ID0016" -replaceWith $filesStat.totalDuplicates
findAndReplaceWholeDoc -document $doc -find "ID0017" -replaceWith $filesStat.totalStolled
findAndReplaceWholeDoc -document $doc -find "ID0018" -replaceWith $filesStat.totalByCompliance
findAndReplaceWholeDoc -document $doc -find "ID0019" -replaceWith (convertCapacity -bytes $filesStat.totalSize)
findAndReplaceWholeDoc -document $doc -find "ID0020" -replaceWith (convertCapacity -bytes $filesStat.totalDuplicatesSize)
findAndReplaceWholeDoc -document $doc -find "ID0021" -replaceWith (convertCapacity -bytes $filesStat.totalStolledSize)
findAndReplaceWholeDoc -document $doc -find "ID0022" -replaceWith $eventsStat.totalByCompliance
findAndReplaceWholeDoc -document $doc -find "ID0023" -replaceWith $eventsStat.avgByDay
findAndReplaceWholeDoc -document $doc -find "ID0024" -replaceWith ([math]::Round($eventsStat.avgByHour, 2))
findAndReplaceWholeDoc -document $doc -find "ID0025" -replaceWith ($ldapExplore.countByDomain.param)


##MAKE COLORS AND TEXT FOR RISK-FACTOR PANEL
if($ldapStat.risk -le 0.3){
    $riskPanel11 = $doc.Shapes.Range("Rectangle 11")
    $riskPanel11.fill.ForeColor.RGB = 5296274
    $riskPanel11.TextFrame.TextRange.text = "–»— -‘¿ “Œ– Õ»« »…"
    }
    elseif($ldapStat.risk -le 0.8){
        $riskPanel11 = $doc.Shapes.Range("Rectangle 11")
        $riskPanel11.fill.ForeColor.RGB = 49407
        $riskPanel11.TextFrame.TextRange.text = "–»— -‘¿ “Œ– —–≈ƒÕ»…"
        }
        elseif($ldapStat.risk -le 1){
            $riskPanel11 = $doc.Shapes.Range("Rectangle 11")
            $riskPanel11.fill.ForeColor.RGB = 255
            $riskPanel11.TextFrame.TextRange.text = "–»— -‘¿ “Œ– ¬€—Œ »…"
            }

if($filesStat.risk -le 0.3){
    $riskPanel17 = $doc.Shapes.Range("Rectangle 17")
    $riskPanel17.fill.ForeColor.RGB = 5296274
    $riskPanel17.TextFrame.TextRange.text = "–»— -‘¿ “Œ– Õ»« »…"
    }
    elseif($filesStat.risk -le 0.8){
        $riskPanel17 = $doc.Shapes.Range("Rectangle 17")
        $riskPanel17.fill.ForeColor.RGB = 49407
        $riskPanel17.TextFrame.TextRange.text = "–»— -‘¿ “Œ– —–≈ƒÕ»…"
        }
        elseif($filesStat.risk -le 1){
            $riskPanel17 = $doc.Shapes.Range("Rectangle 17")
            $riskPanel17.fill.ForeColor.RGB = 255
            $riskPanel17.TextFrame.TextRange.text = "–»— -‘¿ “Œ– ¬€—Œ »…"
            }

if($eventsStat.risk -le 0.3){
    $riskPanel7 = $doc.Shapes.Range("Rectangle 7")
    $riskPanel7.fill.ForeColor.RGB = 5296274
    $riskPanel7.TextFrame.TextRange.text = "–»— -‘¿ “Œ– Õ»« »…"
    }
    elseif($eventsStat.risk -le 0.8){
        $riskPanel7 = $doc.Shapes.Range("Rectangle 7")
        $riskPanel7.fill.ForeColor.RGB = 49407
        $riskPanel7.TextFrame.TextRange.text = "–»— -‘¿ “Œ– —–≈ƒÕ»…"
        }
        elseif($eventsStat.risk -le 1){
            $riskPanel7 = $doc.Shapes.Range("Rectangle 7")
            $riskPanel7.fill.ForeColor.RGB = 255
            $riskPanel7.TextFrame.TextRange.text = "–»— -‘¿ “Œ– ¬€—Œ »…"
            }

##SAVE TO
if (!$makeDOCX){
    $doc.ExportAsFixedFormat("$PSScriptRoot\report_$dateString.pdf", 17)
    
}
else {
    $doc.ExportAsFixedFormat("$PSScriptRoot\report_$dateString.pdf", 17)
    $doc.SaveAs([ref]"$PSScriptRoot\report_$dateString.docx")
    }

##CLOSE DOC and WORD
$doc.Close()
$word.Quit()

##remove TMP folder
Remove-Item "$PSScriptRoot\tmp" -Recurse 

##Convert bytes
function convertCapacity (
    $bytes, 
    [int]$precision
    ) 
{
    foreach ($i in ("¡"," ¡","Ã¡","√¡","“¡")) {
        if (($bytes -lt 1000) -or ($i -eq "“¡")){
           $bytes = ($bytes).tostring("F0" + "$precision")
           return $bytes + " $i"
           } else {
               $bytes /= 1KB
           }
        }
    }

##Change marks in document
function mkvsrequest { ##return JsonObject
    param(
        [string]$typeReport = "",
        [psobject]$headers
    )
    ##Make full URI
    $uri = $makves_url + $typeReport

    if (!$makves_url) {
        $uri = ""
        Add-Type -AssemblyName 'System.Net.Http'
    }
   
    if ($uri) { 
        Try
        {
            $response = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers ##request
        }
        Catch {
            Write-Host "Error send data to server:" + $PSItem.Exception.Message
            return
        }
    }
    $responseObj = ConvertFrom-Json $([String]::new($response.Content))
return $responseObj
}

function responseToCSV { ##Return CSV File in target folder
    param (
        [psobject]$inputObj,
        [string]$CSVPath = "",
        [string[]]$columns = @()
    )
    if(!$columns){
        $columns = "*"
    }
    
    $inputObj.items | Select-Object $columns | export-csv -encoding utf8 -path $CSVPath -UseCulture -notypeinformation
   }

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
       
       Write-Host "$Find in paragraph" $res
        }

    $document.Shapes | ForEach-Object {
        $rng = $_.TextFrame.TextRange;
        [psobject]$findReplace = $rng.Find
        $findReplace.ClearFormatting();
        $res = $findReplace.Execute($find,  $True, $false, $null, $null, $null, $null, $null, $null, $replaceWith,
        2, $null, $null, $null, $null)
       
      Write-Host "$Find in shapes" $res
       
    }

}
