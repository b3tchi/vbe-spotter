$excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
Write-Host $excel.Name

Write-Host $excel.workbooks(1).worksheets(1).Name

#$vbaarr = $excel.workbooks(1).worksheets(1).UsedRange.Value
$vbaarr = $excel.Run('arrRet')

Write-Host Get-Member $vbaarr
Write-Host $vbaarr.GetType()
#Write-Host $vbaarr.Value.ToString()

Write-Host $vbaarr[1]
$array2 = New-Object 'object[,]' 2,2
$array2[1,1] = 'test'
$ok = $excel.Run('arrIn',$array2)