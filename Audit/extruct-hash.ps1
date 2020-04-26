
##$computers = Get-ADComputer -filter *  | Select -Exp Name
$path = "C:\Users\fscontrol_dev\Desktop\Request_Files"
$file = "fingerprints.csv"
$hashlen = 1048576
$filters = "*.xlsx", "*.xls", "*.docx", "*.doc", "*.pdf"



function filterLength($fn) {
    return $true                
    foreach ( $sz in $sizelist )
    {
        if ($sz -eq $fn.Length) {
            return $true
        }        
    }
    return $false
}

Function getFileHash([String] $FileName, $HashName = "SHA1") {
    if ($hashlen -eq 0) {
        $FileStream = New-Object System.IO.FileStream($FileName, "Open", "Read") 
        $StringBuilder = New-Object System.Text.StringBuilder 
        [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($FileStream) | ForEach-Object { [Void]$StringBuilder.Append($_.ToString("x2")) } 
        $FileStream.Close() 
        $StringBuilder.ToString()
    }
    else {
        $StringBuilder = New-Object System.Text.StringBuilder 
        $binaryReader = New-Object System.IO.BinaryReader(New-Object System.IO.FileStream($FileName, "Open", "Read"))
    
        $bytes = $binaryReader.ReadBytes($hashlen)
        $binaryReader.Close() 
        if ($bytes -ne 0) {
            [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($bytes) | ForEach-Object { [Void]$StringBuilder.Append($_.ToString("x2")) }
        }
        $StringBuilder.ToString()
    }
}


function inspect($fn) {
    $hash = getFileHash($fn)    
    $message = "$($fn.Length);$($hash)"
    try {
        $message | Out-File -FilePath $file -Encoding UTF8 -Append
    } catch {
 
    }
}

Get-ChildItem -Path $path -include $filters -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { filterLength $_ } | ForEach-Object {
    inspect $_ $file
}
  