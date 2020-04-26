param (
     [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
 )

 ## Init web server 
$uri = $makves_url + "/file/explore"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}

if ($uri -ne "") {
    Try
    {
        $response = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers
    }
    Catch {
        Write-Host "Error send data to server:" + $PSItem.Exception.Message
        return
    }
}

$jsonObj = ConvertFrom-Json $([String]::new($response.Content))

$jsonObj.items | ForEach-Object {
    $cur = $_
    Write-Host $cur.type " " $cur.folder " " $cur.name 
}

