Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$excel = New-Object -ComObject excel.application
$excel.visible = $true

$extens = "csv"
$fileLocation = Read-Host -Prompt 'Program current path '
$folderLoc = $fileLocation + "\rep"
$files = (Get-ChildItem $folderLoc)


ForEach($file in $files) {
    cd $fileLocation
    $path = ($file.fullname).substring(0, ($file.FullName).lastindexOf("."))
    $workbook = $excel.workbooks.open($path) 
    $path2 = $fileLocation + "\result\" + $file.Name.substring(0, ($file.Name).lastindexOf(".")) +"."+ $extens
    $workbook.saveas($path2, 6)
    $workbook.close()  	
}

$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()