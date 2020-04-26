param (
    [string]$makves_url = "http://localhost:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$fn = "result.xml"
)


## Init web server 
$uri = $makves_url + "/identity/upload"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue }

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}
 

 
$Body =  Get-Content $fn
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
