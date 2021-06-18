# $app = New-Object -ComObject 'shell.application'
# $handles = (Get-Process -Name iexplore).MainWindowHandle
# foreach ($handle in $handles){
#    $window = $app.windows() | Where-Object {$_.HWND -eq $handle}
#    Write-Output $window.Document.documentElement
# #    $window.Document.documentElement | Out-File ".\Desktop\test\$handle.txt"
# }


$app = New-Object -ComObject 'shell.application'
$handles = (Get-Process -Name MSACCESS).MainWindowHandle
foreach ($handle in $handles){
   $window = $app.windows() | Where-Object {$_.HWND -eq $handle}
#    $window = $app.windows() | Where-Object {$_.HWND -eq $handle}
   Write-Output $window.CurrentProject.Path
#    $window.Document.documentElement | Out-File ".\Desktop\test\$handle.txt"
}