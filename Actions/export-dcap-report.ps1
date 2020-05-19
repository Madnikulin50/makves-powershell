param(
    [string]$template = "{{.ROOT}}/powershell/dcap-template.docx",
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$tmpFile = "",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [bool]$makeDOCX = $false,
    [string]$outfile = "{{.ROOT}}/{{.OUT}}"
    )

## Init web server 
$pair = "${makves_user}:${makves_pwd}" ##Make a string for auth pair login/pass
$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)
$basicAuthValue = "Basic $base64"
[psobject]$pairPassLog = @{ Authorization = $basicAuthValue} ##make a headers for invok-webrequest auth

## Make string from date for filenames   
$dateString = ((Get-date).ToString('yyMMddhhmmss'))

$scriptPath = "{{.ROOT}}"

if ($tmpFile -eq "") {
    $tmpFile = "$scriptPath\temp_$dateString.docx" 
}
 

## Make tmp file
Copy-Item $template $tmpFile


##Convert bytes
function convertCapacity (
    $bytes, 
    [int]$precision
    ) 
{
    foreach ($i in ("bytes","KB","MB","GB","TB")) {
        if (($bytes -lt 1000) -or ($i -eq "TB")){
           $bytes = ($bytes).tostring("F0" + "$precision")
           return $bytes + " $i"
           } else {
               $bytes /= 1KB
           }
        }
    }

    
function convertRisk (
   [float]$data
    ) 
{
    return [math]::Round($data * 100)
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
            Write-Host "Error send data to server" $typeReport ":" $PSItem.Exception.Message
            return $false
        }
    }
    $responseObj = ConvertFrom-Json $([String]::new($response.Content))
    return $responseObj
}

function findAndReplaceWholeDoc2 { 
    param (
        [psobject]$document,
        [psobject]$word,
        $replace
    )
    try {
        ForEach ($storyRge in $Document.StoryRanges) {
            Do {
                try {
                    [psobject]$findReplace = $storyRge.Find
                    if ($null -ne $findReplace) {
                        $findReplace.ClearFormatting();
                        $replace | ForEach-Object { 
                            $res = $findReplace.Execute($_.find,  $True, $false, $null, $null, $null, $null, $null, $null, $_.replace,
                            2, $null, $null, $null, $null)
                            if ($true -eq $res) {
                                Write-Host "$_.find in paragraph and replace"
                            }
                        }
                    }
                }
                catch {
                    Write-Host "Error replace $find"
                 } 
                #check for linked Ranges
                $storyRge = $storyRge.NextStoryRange
            } Until (!$storyRge) #null is False
    
        }
        
        return
    }
    catch {
       Write-Host "Error replace $find"
    }
}


function main() {
    ##Make new Object MS word
    

    ##Get stats from makves server
    $ldapStat = mkvsrequest -typeReport "/ldap/stat" -headers $pairPassLog
    if ($false -eq $ldapStat) {
        return
    }
    $ldapExplore = mkvsrequest -typeReport "/ldap/explore" -headers $pairPassLog
    if ($false -eq $ldapExplore) {
        return
    }
    $usersStat = mkvsrequest -typeReport "/ldap/stat?type=users" -headers $pairPassLog
    if ($false -eq $usersStat) {
        return
    }
    $compsStat = mkvsrequest -typeReport "/ldap/stat?type=computers" -headers $pairPassLog
    if ($false -eq $compsStat) {
        return
    }
    $eventsStat = mkvsrequest -typeReport "/events/stat?total=true&stat=true" -headers $pairPassLog
    if ($false -eq $eventsStat) {
        return
    }
    $filesStat = mkvsrequest -typeReport "/file/stat?total=true&duplicates=true&compliance=true&stolled=true&byTime=true&byComputer=true&byType=true&byCompliance=true" -headers $pairPassLog
    if ($false -eq $filesStat) {
        return
    }

    $mbStat = mkvsrequest -typeReport "/mailbox/stat" -headers $pairPassLog
    if ($false -eq $mbStat) {
        return
    }

    $usersTopRisk = mkvsrequest -typeReport '/ldap/explore?q=%7B%22filter%22:%7B%22type%22:%7B%22type%22:%22in%22,%22filter%22:%5B%22user%22,%22group%22%5D%7D%7D,%22sort%22:%5B%7B%22colId%22:%22basic_score%22,%22sort%22:%22desc%22%7D%5D,%22page%22:1%7D' -headers $pairPassLog
    if ($false -eq $usersTopRisk) {
        return
    }

    $filesTopRisk = mkvsrequest -typeReport '/file/explore?q=%7B%22filter%22:%7B%7D,%22sort%22:%5B%7B%22colId%22:%22basic_score%22,%22sort%22:%22desc%22%7D%5D,%22page%22:1%7D' -headers $pairPassLog
    if ($false -eq $filesTopRisk) {
        return
    }
   


    $replace = @()

    
    $replace += New-Object PSObject -Property @{
        Find= "ID0001"
        Replace=$usersStat.total   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0002"
        Replace=$compsStat.total   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0003"
        Replace=$eventsStat.total   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0004"
        Replace=$filesStat.totalFiles   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0005"
        Replace=$usersStat.totalGroups   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0006"
        Replace=$usersStat.totalDisabled   
    }

    
    $replace += New-Object PSObject -Property @{
        Find= "ID0007"
        Replace=$usersStat.totalInactive   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0008"
        Replace=$usersStat.totalPasswordExpired   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0009"
        Replace=$usersStat.totalDontExpirePassword   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0010"
        Replace=$usersStat.totalLockout   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0011"
        Replace=$usersStat.totalEmptyGroups   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0012"
        Replace=$usersStat.totalMnsLogonAccount   
    }

    
    $replace += New-Object PSObject -Property @{
        Find= "URLOW"
        Replace=$usersStat.totalsRiskMinimal   
    }
    $replace += New-Object PSObject -Property @{
        Find= "URMED"
        Replace=$usersStat.totalsRiskMedium   
    }
    $replace += New-Object PSObject -Property @{
        Find= "URHIGH"
        Replace=$usersStat.totalsRiskCritical   
    }

    #
    $replace += New-Object PSObject -Property @{
        Find= "RU0001"
        Replace=$usersTopRisk.items[0].cn   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RU0002"
        Replace=$usersTopRisk.items[1].cn   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RU0003"
        Replace=$usersTopRisk.items[2].cn   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RU0004"
        Replace=$usersTopRisk.items[3].cn   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RU0005"
        Replace=$usersTopRisk.items[4].cn   
    }
    

    $replace += New-Object PSObject -Property @{
        Find= "URISK001"
        Replace=(convertRisk -data $usersTopRisk.items[0].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "URISK002"
        Replace=(convertRisk -data $usersTopRisk.items[1].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "URISK003"
        Replace=(convertRisk -data $usersTopRisk.items[2].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "URISK004"
        Replace=(convertRisk -data $usersTopRisk.items[3].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "URISK005"
        Replace=(convertRisk -data $usersTopRisk.items[4].basic_score)   
    }


    
    $replace += New-Object PSObject -Property @{
        Find= "ID0013"
        Replace=$compsStat.totalDisabled   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0014"
        Replace=$compsStat.totalInactive   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0015"
        Replace=$filesStat.totalFolders   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0016"
        Replace=$filesStat.totalDuplicates   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0017"
        Replace=$filesStat.totalStolled   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0018"
        Replace=$filesStat.totalByCompliance   
    }


    
    $replace += New-Object PSObject -Property @{
        Find= "COUNTSTD001"
        Replace=$filesStat.byCompliance[0].count   
    }
    $replace += New-Object PSObject -Property @{
        Find= "COUNTSTD002"
        Replace=$filesStat.byCompliance[1].count   
    }
    $replace += New-Object PSObject -Property @{
        Find= "COUNTSTD003"
        Replace=$filesStat.byCompliance[2].count   
    }
    $replace += New-Object PSObject -Property @{
        Find= "COUNTSTD004"
        Replace=$filesStat.byCompliance[3].count   
    }
    $replace += New-Object PSObject -Property @{
        Find= "COUNTSTD005"
        Replace=$filesStat.byCompliance[4].count   
    }

    $replace += New-Object PSObject -Property @{
        Find= "STD001"
        Replace=$filesStat.byCompliance[0].compliance   
    }
    $replace += New-Object PSObject -Property @{
        Find= "STD002"
        Replace=$filesStat.byCompliance[1].compliance   
    }
    $replace += New-Object PSObject -Property @{
        Find= "STD003"
        Replace=$filesStat.byCompliance[2].compliance   
    }
    $replace += New-Object PSObject -Property @{
        Find= "STD004"
        Replace=$filesStat.byCompliance[3].compliance   
    }
    $replace += New-Object PSObject -Property @{
        Find= "STD005"
        Replace=$filesStat.byCompliance[4].compliance   
    }


    #
    $replace += New-Object PSObject -Property @{
        Find= "RFILE0001"
        Replace=$filesTopRisk.items[0].folder+"\"+$filesTopRisk.items[0].name   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RFILE0002"
        Replace=$filesTopRisk.items[1].folder+"\"+$filesTopRisk.items[1].name   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RFILE0003"
        Replace=$filesTopRisk.items[2].folder+"\"+$filesTopRisk.items[2].name   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RFILE0004"
        Replace=$filesTopRisk.items[3].folder+"\"+$filesTopRisk.items[3].name   
    }
    $replace += New-Object PSObject -Property @{
        Find= "RFILE0005"
        Replace=$filesTopRisk.items[4].folder+"\"+$filesTopRisk.items[4].name   
    }

    $replace += New-Object PSObject -Property @{
        Find= "FRISK001"
        Replace=(convertRisk -data $filesTopRisk.items[0].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "FRISK002"
        Replace=(convertRisk -data $filesTopRisk.items[1].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "FRISK003"
        Replace=(convertRisk -data $filesTopRisk.items[2].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "FRISK004"
        Replace=(convertRisk -data $filesTopRisk.items[3].basic_score)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "FRISK005"
        Replace=(convertRisk -data $filesTopRisk.items[4].basic_score)   
    }


    $replace += New-Object PSObject -Property @{
        Find= "FRLOW"
        Replace=$filesStat.totalsRiskMinimal   
    }
    $replace += New-Object PSObject -Property @{
        Find= "FRMED"
        Replace=$filesStat.totalsRiskMedium   
    }
    $replace += New-Object PSObject -Property @{
        Find= "FRHIGH"
        Replace=$filesStat.totalsRiskCritical   
    }

    $replace += New-Object PSObject -Property @{
        Find= "ID0019"
        Replace=(convertCapacity -bytes $filesStat.totalSize)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0020"
        Replace=(convertCapacity -bytes $filesStat.totalDuplicatesSize)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0021"
        Replace=(convertCapacity -bytes $filesStat.totalStolledSize)   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0022"
        Replace=$eventsStat.totalByCompliance   
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0023"
        Replace=$eventsStat.avgByDay
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0024"
        Replace=([math]::Round($eventsStat.avgByHour, 2))  
    }
    $replace += New-Object PSObject -Property @{
        Find= "ID0025"
        Replace= ($usersStat.countByDomain[0].param)  
    }


    #
    $replace += New-Object PSObject -Property @{
        Find= "MB0001"
        Replace=$mbStat.total  
    }
    $replace += New-Object PSObject -Property @{
        Find= "MB0002"
        Replace=$mbStat.totalFolders  
    }
    $replace += New-Object PSObject -Property @{
        Find= "MB0003"
        Replace= (convertCapacity -bytes $mbStat.totalSize)  
    }
    $replace += New-Object PSObject -Property @{
        Find= "MRLOW"
        Replace= $mbStat.totalsRiskMinimal  
    }
    $replace += New-Object PSObject -Property @{
        Find= "MRMED"
        Replace= $mbStat.totalsRiskMedium  
    }
    $replace += New-Object PSObject -Property @{
        Find= "MRHIGH"
        Replace= $mbStat.totalsRiskCritical  
    }


    [psobject]$word = New-Object -ComObject Word.Application
    $word.Visible = $falses
    [psobject]$doc = $word.Documents.Open($tmpFile)

    findAndReplaceWholeDoc2 -document $doc -word $word -replace $replace

    ##Replace in template_url    
    if ($false) {
        ##MAKE COLORS AND TEXT FOR RISK-FACTOR PANEL
        try {
            $riskPanel11 = $doc.Shapes.Range("Rectangle 11")
            if($ldapStat.risk -le 0.3){
                $riskPanel11.fill.ForeColor.RGB = 5296274
                $riskPanel11.TextFrame.TextRange.text = "РИСК-ФАКТОР НИЗКИЙ"
                }
                elseif($ldapStat.risk -le 0.8){
                    $riskPanel11.fill.ForeColor.RGB = 49407
                    $riskPanel11.TextFrame.TextRange.text = "РИСК-ФАКТОР СРЕДНИЙ"
                    }
                    elseif($ldapStat.risk -le 1){
                    $riskPanel11.fill.ForeColor.RGB = 255
                        $riskPanel11.TextFrame.TextRange.text = "РИСК-ФАКТОР КРИТИЧЕСКИЙ"
                        }

        }
        catch {
            Write-Host "Error found Rectangle 11"
        }

        try {
            $riskPanel17 = $doc.Shapes.Range("Rectangle 17")
            if($filesStat.risk -le 0.3){
                $riskPanel17.fill.ForeColor.RGB = 5296274
                $riskPanel17.TextFrame.TextRange.text = "РИСК-ФАКТОР НИЗКИЙ"
                }
                elseif($filesStat.risk -le 0.8){
                    $riskPanel17.fill.ForeColor.RGB = 49407
                    $riskPanel17.TextFrame.TextRange.text = "РИСК-ФАКТОР СРЕДНИЙ"
                    }
                    elseif($filesStat.risk -le 1){
                        $riskPanel17.fill.ForeColor.RGB = 255
                        $riskPanel17.TextFrame.TextRange.text = "РИСК-ФАКТОР КРИТИЧЕСКИЙ"
                        }

        }
        catch {
            Write-Host "Error found Rectangle 17"
        }

        try {
            $riskPanel17 = $doc.Shapes.Range("Rectangle 7")
            if($eventsStat.risk -le 0.3){
                $riskPanel7.fill.ForeColor.RGB = 5296274
                $riskPanel7.TextFrame.TextRange.text = "РИСК-ФАКТОР НИЗКИЙ"
                }
                elseif($eventsStat.risk -le 0.8){
                    $riskPanel7.fill.ForeColor.RGB = 49407
                    $riskPanel7.TextFrame.TextRange.text = "РИСК-ФАКТОР СРЕДНИЙ"
                    }
                    elseif($eventsStat.risk -le 1){
                        $riskPanel7.fill.ForeColor.RGB = 255
                        $riskPanel7.TextFrame.TextRange.text = "РИСК-ФАКТОР КРИТИЧЕСКИЙ"
                        }

        }
        catch {
            Write-Host "Error found Rectangle 7"
        }
    }

    
    Write-Host "Start save"
    $doc.Save()

    ##SAVE TO
    if ($false -eq $makeDOCX){
        $doc.ExportAsFixedFormat($outfile, 17)
    }
    else{        
        $doc.SaveAs($outfile)
    }
    Write-Host "End save"
    ##CLOSE DOC and WORD
    $doc.Close()
    $word.Quit()

    ##remove TMP file
    Remove-Item $tmpFile

}

main


