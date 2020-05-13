param (
    [string]$URL = "https://1cweb.tass.ru/zup_centr/ws/Upload.1cws",
    [string]$fn = "c:\work\res.xml",
    [string]$request = @"
    <Envelope xmlns="http://schemas.xmlsoap.org/soap/envelope/">
        <Body>
            <UploadEmployees xmlns="EmployeesData"/>
        </Body>
    </Envelope>
"@,
[string]$makves_url = "http://127.0.0.1:8000",
[string]$makves_user = "admin",
[string]$makves_pwd = "admin"

)
if (Test-Path $fn) {
    Remove-Item $fn
}

## Init web server 
$uri = $makves_url + "/identity/upload"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}


 

$resp = Invoke-WebRequest -Uri $URL -Headers (@{SOAPAction='Read'
Authorization='Basic ЧЧЧЧЧЧЧЧ='}) -Method Post -Body $request -ContentType application/xml 

$body = [System.Text.Encoding]::UTF8.GetBytes($resp.Content);

$body | Out-File -FilePath $fn

$Body = $ReturnXml.Envelope.InnerText
    $fname = "1c.xml"
    $boundary = [guid]::NewGuid().ToString()
    $template = "
    --$boundary
    Content-Disposition: form-data; name=""file""; filename="""+$fn+"""

    "+$Body+"
    --$boundary
    Content-Type:text/plain; charset=utf-8
    Content-Disposition: form-data; name=""metadata""

    {""name"":"""+$fn+"""}}

    --$boundary--"

Try {
    $upload_file = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $template -ContentType "multipart/form-data;boundary=$boundary" -Verbose
    Write-Host  "Send data to server:" + $uri + " res:" $upload_file
}
Catch {
    Write-Host "Error send data to server:" + $PSItem.Exception.Message
}

Write-Host 'done'

