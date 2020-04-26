##Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
##$computers = Get-ADComputer -filter *  | Select -Exp Name
$paths = "\\fs.pewete.com\pwt\"
$file = "res1.csv"
$filters = "*.xlsx", "*.xls", "*.docx", "*.doc", "*.pdf"
$hashlen = 1048576

$fpdata = @("
1572655;e8feae4d90b5dbc706c7157bfb004af75729d8f2
165;45526a9156d50d12ee8fc19a91cf076c0150459c
771438;50fd8c62046da3b0158d2376dbf39eb1c198b368
922193;25e076a7802037b13437c900d4ea6e04f248f3d8
922182;bdd61e51c9ed7e532d2d95d967cab6e01a4b192c
627964;d0f04159e7760a3a2946174be1a889179fd68621
27773;207dd7c7e4d6e8a06ffb97d92583b01c0a0be49a
15061;9e65661795438b257e71df6442d62b13af48ce67
13673;c835d3b6d2d3aa5d4b8f7b7b0d2aebb57b716fe2
900710;9ab5ace12f7933d00c1b049ca94c5f1ff90761bd
16287788;bd140502c8e4d46e083896ccef86cdc15ddaed05
3572073;b448fb020662936de4456f07872a2c84924796da
812427;317845015da2bf23e801a2901d1d784c81070e81
2737061;d1d4a072a61c6568f622af45650e6c39ad8e42f3
629891;f4688376b7ead5b5cd536dac25e73f174e360110
46675783;0064fc982ac07ff366cf765b76eb27253736d7e3
11767524;44a44bedff204134fac449409ce8841a8b186f6a
654926;8a18a16414738f1207fa81cdc460c73b0c46fdf8
158496;0d91c7c492a31f0890e09452d7e1b5ac16b6aff1
1826505;65ee0459a67d37d1af2e14f73cb69813fce0929d
1983866;34e0c5bf3f25d8d106128b3d3be05aca8512e67f
793906;91b2aa3aa834e8d2c93028ccdd0e2144f6001d80
4073984;ff4fb0fb16cb4c54f118b2b1d4223fe097e969db
964608;c9f4797d30f5d20f6d9916e7402da2b00593c3f5
223744;ae936c3e9afb4aac2cf8a79601cbaef1d4aa21f3
1114679;beb2babd1d02014014240801f8987baea6324204
1124131;142a663643cc386e5d17f4bd9d65059380cb3378
72713;0d56e2f758cc2c261dfa45a549c8c1fac9cfe8d2
40011;2ccc546fd909fd3bb408fe1de23fe2b79f73986d
88064;ed39e681a0d10782a43df360fbc99fecfa1515b2
18799748;8197b58e46a587879b78cbee41ad6ee08a1b8b08
2608532;d60b6acb0317c47f5b256a7f3002619e291e70c8
119851;c983e4c6fe23354699e131f9cdb8cfe3c82cca25
7918258;7cb2e5172ab6f24451d9207c8cb9b277166c5e4c
")

$lh = @{label = "Length"; expression = { $_.L } }

$fp = $fpdata | ConvertFrom-Csv -Delimiter ";" -Header "L", "Hash" | Select-Object -Property $lh, Hash

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

function filterLength($fn) {   
    if ($_.PSIsContainer -eq $true) {
        return $false
    }
    Write-Host $fn
    try {
        foreach ( $f in $fp ) {
            if ($f.Length -eq $fn.Length) {
                $h = getFileHash($fn)
                if ($f.Hash -eq $h) {
                    return $true
                }
            }        
        }
        return $false
    }
    Catch {
        return $false
    }

}
function inspect($fn, $file) {    
    
    $message = $env:computername+";$($fn);$($fn.Length)"
    try {
        $message | Out-File -FilePath $file -Encoding UTF8 -Append
    }
    catch {
     
    }
}

function exec($path) {
   
    Write-Host "file $($file)"
    Write-Host "filters $($filters)"
        
    Write-Host "Start $($path)"
               
    try {
        Get-ChildItem -Path $path -include $filters -Recurse -force -file -ErrorAction SilentlyContinue | Where-Object { filterLength $_ } | ForEach-Object {
            inspect $_ $file
        }
    }
    catch {
        Write-Host "Error get files from $($path):$($_.Exception.Message)"
        Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue | Where-Object { filterLength $_ } | ForEach-Object {
            inspect $_ $file
        }
    }
    Write-Host "Finish $($path)"  
}

Write-Host "Start"
$paths | Foreach-Object {
    exec $_
}

Get-PSDrive -PSProvider FileSystem |  Foreach-Object {
    exec $_.Root
}



Write-Host "Finish all"