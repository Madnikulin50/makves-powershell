<#
 .Synopsis
  Сбор данных о файлах в папке

 .Description
  Сбор данных о файлах в папке

 .Parameter File
 Имя файла из которого выделяется текст

 .Example
   # Пример запуска без выделения текста
   Get-Text -File "c:\\work\\test\\" -Outfilename folder_test

#>
function Get-Text {
[CmdletBinding()]
Param(
    [Parameter(ValueFromPipeline)]
    $File
)
    $scriptFolder = $MyInvocation.MyCommand.Module.ModuleBase
   
    Import-Module $scriptFolder"/compliance.dll" -Verbose

    
    function Get-MKVS-DocText([String] $FileName) {
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Word.DisplayAlerts = 0
        $text = ""
        Try
        {
            $catch = $false
            Try{
                $Document = $Word.Documents.Open($FileName, $null, $null, $null, "")
            }
            Catch {
                Write-Host 'Doc is password protected.`r`n'
                if ($logfilename -ne "") {
                    'Doc is password protected.`r`n' | Out-File -FilePath $logfile -Encoding UTF8 -Append
                }
                $catch = $true
            }
            if ($catch -eq $false) {
                $Document.Paragraphs | ForEach-Object {
                    $text += $_.Range.Text
                }
                
            }
        }
        Catch {
            Write-Host $PSItem.Exception.Message
            $Document.Close()
            $Word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
            Remove-Variable Word
        }
        $Document.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
        Remove-Variable Word        
        return $text
    }

    function Get-MKVS-XlsText([String] $FileName) {
        $excel = New-Object -ComObject excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = 0
        $text = ""
        $password
        Try    
        {
            $catch = $false
            Try{
                $wb =$excel.Workbooks.open($path, 0, 0, 5, "")
            }
            Catch{
                Write-Host 'Book is password protected.'
                if ($logfilename -ne "") {
                    'Book is password protected.`r`n' | Out-File -FilePath $logfile -Encoding UTF8 -Append
                }
                $catch = $true
            }
            if ($catch -eq $false) {
                foreach ($sh in $wb.Worksheets) {
                    #Write-Host "sheet: " $sh.Name            
                    $endRow = $sh.UsedRange.SpecialCells(11).Row
                    $endCol = $sh.UsedRange.SpecialCells(11).Column
                    #Write-Host "dim: " $endRow $endCol
                    for ($r = 1; $r -le $endRow; $r++) {
                        for ($c = 1; $c -le $endCol; $c++) {
                            $t = $sh.Cells($r, $c).Text
                            $text += $t
                            #Write-Host "text cel: " $r $c $t
                        }
                    }
                }
            }
        }
        Catch {
            Write-Host $PSItem.Exception.Message
            if ($logfilename -ne "") {
                "Extruct Excel text: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
        }
        #Write-Host "text: " $text
        $excel.Workbooks.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        Remove-Variable excel
        return $text
    }

    $info = New-Object System.IO.FileInfo($file)
    switch ($info.Extension) {
        ".doc" {
            return Get-MKVS-DocText $FileName
        }
        ".docx" {
            return Get-MKVS-DocText $FileName
        }
        ".xls" {
            return Get-MKVS-XlsText $FileName
        }
        ".xlsx" {
            return Get-MKVS-XlsText $FileName
        }
    }
    return ""
}

<#
 .Synopsis
  Сбор данных о файлах в папке

 .Description
  Сбор данных о файлах в папке

 .Parameter File
 Имя файла из которого выделяется текст

 .Example
   # Пример запуска без выделения текста
   Search-File-Compliance -File "c:\\work\\test\\"
#>
function Search-FileCompliance {
    [CmdletBinding()]
    Param(
        [parameter(ValueFromPipeline)]
        [string[]]$File,
        [string[]]$KB=("")
    )
    Begin {
        $scriptFolder = $MyInvocation.MyCommand.Module.ModuleBase
        Import-Module $scriptFolder"/compliance.dll" -Verbose
        $objects = @()
    }

    Process {
        $File | Write-Host 
        $obj = $File | Get-Compliance -KB $KB
        Write-Host $obj
        $objects += $obj
    }
    
    End {
        Write-Verbose "Finish Search-File-Compliance"
        $objects | Write-Host 
        return $objects
    }
}


<#
 .Synopsis
  Сбор данных о файлах в папке

 .Description
  Сбор данных о файлах в папке

 .Parameter File
 Имя файла из которого выделяется текст

 .Example
   # Пример запуска без выделения текста
   Search-Text-Compliance -Text "c:\\work\\test\\"
#>
function Search-TextCompliance {
    [CmdletBinding()]
    Param(
        [parameter(ValueFromPipeline=$True)]
        [string[]]$Text,
        
        [string[]]$KB=("")
    )
    Begin {
        $scriptFolder = $MyInvocation.MyCommand.Module.ModuleBase
        Import-Module $scriptFolder"/compliance.dll" -Verbose
        $objects = @()
    }

    Process {
        $objects += Get-Compliance -File $Text -KB $KB
    }
    
    End {
        Write-Verbose "Finish Search-File-Compliance"
        return $objects
    }
}

Export-ModuleMember -Function Search-FileCompliance
Export-ModuleMember -Function Search-TextCompliance
Export-ModuleMember -Function Get-Text