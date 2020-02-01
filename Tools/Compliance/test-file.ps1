param (
    [string]$file = './test',
    [string]$kb = './fn.json'
 )

Write-Host "file: " $file
Write-Host "kb: " $kb

$kbcontent = Get-Content -Path $kb -Encoding UTF8
Write-Host "kb content: " $kbcontent

$kbdata = $kbcontent | ConvertFrom-Json


function check($item) {
    $ext = $item.Extension
    $path = $item.FullName
    $name = $item.Name
    Write-Host "Start check of: " $path " ext " $ext
    if ($ext -ne ".txt") {
        Write-Host "File without text"
    }
    $text = Get-Content -Path $path -Encoding UTF8
    
    $kbdata | ForEach-Object  {
        if ($_ -eq $null) {
            return
        }
        $result = 0.
        $cur = $_
        Write-Host "Start inspect " +  $name + " by " $cur.category
        $cur.masks | ForEach-Object  {
            $cm = $_
            $mask = $cm.mask -replace "&lt;", '<'
            $r = $text -match $mask
            if ($r.Length -ne 0) {
                Write-Host "Found " $cm.descr "(" $r.Length ") in " +  $name
                #$r | ForEach-Object  {  Write-Host $_  }
                $result += [double]$cm.weight / 100.  - $result * [double]$cm.weight / 100.
            }
        }
        $result = $result * 100
        Write-Host "Finish check of: " $path " result " $result
    
    }
}

function inspect($item) {
    $path = $item.FullName    
        
    if ($item.PSIsContainer -eq $false) {
        check $item
    } else {
        Get-ChildItem $path -Recurse | Foreach-Object {
            Try
            {
                inspect $_
            }
            Catch {
                Write-Host "inspectFile error:" + $PSItem.Exception.Message
            }
        }
    }
}

$cur =  Get-Item $file

inspect $cur


Write-Host "Finish check : " $file