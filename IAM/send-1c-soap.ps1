param (
    [string]$URL = "https://1cweb.tass.ru/zup_centr/ws/Upload.1cws",
    [string]$request = @"
    <Envelope xmlns="http://schemas.xmlsoap.org/soap/envelope/">
        <Body>
            <UploadEmployees xmlns="EmployeesData"/>
        </Body>
    </Envelope>
"@,
    [string]$makves_url = "", ##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
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
 
#$cert = Get-ChildItem -Path Cert:\LocalMachine\My | where-Object {$_.Subject -like 'Subject of certificate'} 
 
#if($cert -ne $null)
#{
    Try
    {
 
        # Sending SOAP Request To Server 
        $soapWebRequest = [System.Net.WebRequest]::Create($URL) 
        $soapWebRequest.ClientCertificates.Add($cert)
        $soapWebRequest.Headers.Add("SOAPAction","Provide method name of the API")
        $soapWebRequest.ContentType = "text/xml;charset=utf-8"
        $soapWebRequest.Accept      = "text/xml"
        $soapWebRequest.Method      = "POST"
        $soapWebRequest.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        $soapWebRequest.UseDefaultCredentials = $true
    
        #Initiating Send
        $requestStream = $soapWebRequest.GetRequestStream() 
        $SOAPRequest.Save($requestStream) 
        $requestStream.Close() 
       
        #Send Complete, Waiting For Response.
        $resp = $soapWebRequest.GetResponse() 
        $responseStream = $resp.GetResponseStream() 
        $soapReader = [System.IO.StreamReader]($responseStream) 
        $ReturnXml = [Xml] $soapReader.ReadToEnd() 
        $responseStream.Close() 
 
    }
    Catch
    {
        Throw $ReturnXml.Envelope.InnerText
    }
 
    $Body = $ReturnXml.Envelope.InnerText
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

    
 
#else
#{
#    $Return = "Certificate not found"
#}
 
#$SOAPRequest = $SOAPRequest.OuterXML