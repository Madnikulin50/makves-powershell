##Send report to Email
param (
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$csvPath = "C:\Program Files (x86)\makves\reports",
    [bool]$files = $true,
    [string[]]$columnsFileExplore = @("type","folder","name","size","computer"), 
    [bool]$ldap = $true,
    [string[]]$columnsldapExplore = @("type","cn","mail","phone","bad_pwd","operating_system"), 
    [bool]$events = $true,
    [string[]]$columnsEventsExplore = @("event_id","category","severity","time", "computer","user", "contents"),
    [string]$EmailFrom ="", ##Var of sender email
    [string]$EmailTo = "", ##var of User mail
    [string]$Body = "", ##Body of mail
    [string]$SmtpServer = "" ##server smtp like "smtp.makves.ru"
     )
     $getCredentials = Get-Credential
     ## Init web server 
 ##Make a string for auth pair login/pass
    
    $pair = "${makves_user}:${makves_pwd}"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $basicAuthValue = "Basic $base64"
##make a headers for invok-webrequest auth
    $headers = @{ Authorization = $basicAuthValue}
   
##make string from date for filenames   
    $dateString = ((Get-date).ToString('yyMMddhhmmss'))


    function csvFromServer { ##create CSV file 
    param (
        [string]$typeReport = "",
        [string]$csvFileName = "",
        [string[]]$columns = @()
    )
    ##make full path to CSV file
    $csvFile = $csvPath + $csvFileName
    ##ake full URI
    $uri = $makves_url + $typeReport
    
    if ($columns -eq @()){
        $columns = "*"
    }

    if ($makves_url -eq "") {
        $uri = ""
        Add-Type -AssemblyName 'System.Net.Http'
    }
    
    if ($uri -ne "") {
        Try
        {
            $response = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers ##
        }
        Catch {
            Write-Host "Error send data to server:" + $PSItem.Exception.Message
            return
        }
    }
    ##create object from WebRequest
    $jsonObj = ConvertFrom-Json $([String]::new($response.Content))
    ##export to json object items to CSV
    $jsonObj.items | Select-Object $columns | export-csv -encoding utf8 -path $csvFile -UseCulture -notypeinformation

    Write-Host "File was create: "  $csvFile
return
}
   ##File/explore
if ($files -eq $true){
    $filesCSVName = '\'+ $dateString + '_file_explore.csv'
    csvFromServer -typeReport "/file/explore" -csvFileName $filesCSVName -columns $columnsFileExplore
    }
    ##ldap/explore
if ($ldap -eq $true){
    $ldapCSVName = '\'+ $dateString + '_ldap_explore.csv'
    csvFromServer -typeReport "/ldap/explore" -csvFileName $ldapCSVName -columns $columnsldapExplore
    }    
    ##events/explore
if ($events -eq $true){
    $eventsCSVName = '\'+ $dateString + '_events_explore.csv'
    csvFromServer -typeReport "/events/explore" -csvFileName $eventsCSVName -columns $columnsEventsExplore
    }

## ZIP files
$filesPathToZip = $csvPath +'\' + $dateString + '*'
$destFileName = $csvPath +'\'+ $dateString + '_reports.zip'
Compress-Archive -Path $filesPathToZip -DestinationPath $destFileName -CompressionLevel Optimal

##Send-mailmessage -....  -Attachments

send-mailmessage -from $EmailFrom -to $EmailTo -subject "Маквес отчеты" -Attachment $destFileName -smtpServer $SmtpServer -Credential $getCredentials
