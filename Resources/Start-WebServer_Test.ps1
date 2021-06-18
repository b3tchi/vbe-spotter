$Hso = New-Object Net.HttpListener
$Hso.Prefixes.Add("http://localhost:8000/")
$Hso.Start()
While ($Hso.IsListening) {
    $HC = $Hso.GetContext()
    $HRes = $HC.Response
    # Write-Output $HC.Request
    # $HRes.Headers.Add("Content-Type","text/plain")
    # $Buf = [Text.Encoding]::UTF8.GetBytes((Get-Content (Join-Path $Pwd ($HC.Request).RawUrl)))
    $Buf =Get-Content(Join-Path $Pwd ($HC.Request).RawUrl) -Raw
    $HRes.ContentLength64 = $Buf.Length
    $HRes.OutputStream.Write($Buf,0,$Buf.Length)
    $HRes.Close()
}
$Hso.Stop()