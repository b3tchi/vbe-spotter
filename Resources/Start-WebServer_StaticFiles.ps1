#!/snap/bin/powershell

Import-Module $PSScriptRoot/AccessRunDb.ps1
Import-Module $PSScriptRoot/Config.ps1
Import-Module $PSScriptRoot/Router.ps1

##VARIABLES##
# $htmlFilesPath = "C:\Users\czJaBeck\Documents\Vbox\svelte_Template_IE_XMLTest\public"
# $dbFullPath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

$app = GetApp $dbFullPath
# $db = $app.CurrentDb()

##HTTP LISTENER PREPARATION##
$Hso = New-Object Net.HttpListener
$Hso.Prefixes.Add("http://localhost:$srvPort/")
$Hso.Start()
# Register-EngineEvent PowerShell.Exiting –Action {
#     Write-Host "Close Event"
# }
try{
  Write-Host "$(Get-Date -Format s) Custom  Powershell webserver started."
  While ($Hso.IsListening) {
    $HC = $Hso.GetContext()
    $HRes = $HC.Response
    # Write-Output $HC.Request
    # $HRes.Headers.Add("Content-Type","text/plain")
    $RequestItem = $HC.Request
    [string]$RequestText = $HC.Request | Select-Object -Property HttpMethod, Url, HasEntityBody, ContentType  | ConvertTo-Json -Compress
    # [string]$RequestText = $HC.Request | Format-List

    $RECEIVED = '{0} {1}' -f $RequestItem.httpMethod, $RequestItem.Url.LocalPath
    Write-Host $RECEIVED

    # stop powershell webserver, nothing to do here
    if($RECEIVED -eq "GET /quit"){
      Write-Host "$(Get-Date -Format s) Stopping powershell webserver..."
      $HRes.Close()
      break
    }

    switch($RequestItem.httpMethod){
      "GET"{
        $RequestedUrl = ($RequestItem).RawUrl
        #masking for /
        if($RequestedUrl -eq "/"){
          $RequestedUrl = "/index.html"
        }
        #adjustment for css
        if($RequestedUrl -like "*.css"){
          Write-Host "Css"
          $HRes.Headers.Add("Content-Type","text/css")
        }
        $Path = (Join-Path $htmlFilesPath $RequestedUrl)
        Write-Host $Path
        if(Test-Path $Path -PathType Leaf){
          # Buf =Get-Content(Join-Path $Pwd ($HC.Request).RawUrl) -Raw
          $Buf = [Text.Encoding]::UTF8.GetBytes((Get-Content $Path -Raw))
          $HRes.ContentLength64 = $Buf.Length
          $HRes.OutputStream.Write($Buf,0,$Buf.Length)
        }else{
          Write-Host "file not found"
        }
      break
      }

      "POST"{
        # "OPTIONS"{
        # Write-Host "Post"
        # only if there is body data in the request
        if ($RequestItem.HasEntityBody){

          # set default message to error message (since we just stop processing on error)
          # $RESULT = "Received corrupt or incomplete form data"

          # check content type
          if ($RequestItem.ContentType){

            if($RECEIVED -eq "POST /command"){
              # read complete header (inkl. file data) into string

              $inputStream = $RequestItem.InputStream
              $Encoding = $RequestItem.ContentEncoding

              $READER = New-Object System.IO.StreamReader($inputStream, $Encoding)
              $DATA = $READER.ReadToEnd()
              $READER.Close()
              $RequestItem.InputStream.Close()

              Write-Host "Request Data:"
              Write-Host $DATA

              $jsonQ = $DATA | ConvertFrom-Json
              # TODO Prepare response Script

              $resp = AccessCmd $app $jsonQ.name $jsonQ.arguments

              request $RequestText $DATA

              if ($resp.Status -eq 500){
                $Hres.StatusCode = 500
              }

              $JSONRESPONSE = $resp | ConvertTo-Json

              Write-Host $JSONRESPONSE
              $HRes.AddHeader("Content-Type","text/json")
              $HRes.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
              $HRes.AddHeader("Server", "Powershell Webserver/1.2 on ")

              # return HTML answer to caller
              $BUFFER = [Text.Encoding]::UTF8.GetBytes($JSONRESPONSE )
              $HRes.ContentLength64 = $BUFFER.Length
              $HRes.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
            }

            if($RECEIVED -eq "POST /procedure"){
              # read complete header (inkl. file data) into string

              $inputStream = $RequestItem.InputStream
              $Encoding = $RequestItem.ContentEncoding

              $READER = New-Object System.IO.StreamReader($inputStream, $Encoding)
              $DATA = $READER.ReadToEnd()
              $READER.Close()
              $RequestItem.InputStream.Close()

              Write-Host "Request Data:"
              Write-Host $DATA

              $pson = $DATA | ConvertFrom-Json
              # TODO Prepare response Script

              request $RequestText $DATA

              $resp = AccessProcedure $app $pson

              $JSONRESPONSE = $resp | ConvertTo-Json
              # $JSONRESPONSE = AccessCmd $app "DbMsg" "Test Messagebox"

              $HRes.AddHeader("Content-Type","text/json")
              $HRes.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
              $HRes.AddHeader("Server", "Powershell Webserver/1.2 on ")

              # return HTML answer to caller
              $BUFFER = [Text.Encoding]::UTF8.GetBytes($JSONRESPONSE )
              $HRes.ContentLength64 = $BUFFER.Length
              $HRes.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
            }

            if($RECEIVED -eq "POST /query"){
              # read complete header (inkl. file data) into string

              $inputStream = $RequestItem.InputStream
              $Encoding = $RequestItem.ContentEncoding

              $READER = New-Object System.IO.StreamReader($inputStream, $Encoding)
              $DATA = $READER.ReadToEnd()
              $READER.Close()
              $RequestItem.InputStream.Close()

              Write-Host "Request Data:"
              Write-Host $DATA

              $jsonQ = $DATA | ConvertFrom-Json
              # TODO Prepare response Script

              request $RequestText $DATA
              $JSONRESPONSE = AccessJSON $app $jsonQ.name
              # $JSONRESPONSE = AccessCmd $app "DbMsg" "Test Messagebox"

              $HRes.AddHeader("Content-Type","text/json")
              $HRes.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
              $HRes.AddHeader("Server", "Powershell Webserver/1.2 on ")

              # return HTML answer to caller
              $BUFFER = [Text.Encoding]::UTF8.GetBytes($JSONRESPONSE )
              $HRes.ContentLength64 = $BUFFER.Length
              $HRes.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
            }

            if($RECEIVED -eq "POST /action"){
              # if($RECEIVED -eq "OPTIONS /query"){

              # retrieve boundary marker for header separation
              # $BOUNDARY = $NULL
              # if ($RequestItem.ContentType -match "boundary=(.*);")
              # {	$BOUNDARY = "--" + $MATCHES[1] }
              # else
              # { # marker might be at the end of the line
              # 	if ($RequestItem.ContentType -match "boundary=(.*)$")
              # 	{ $BOUNDARY = "--" + $MATCHES[1] }
              # }
              # if ($BOUNDARY)
              # { # only if header separator was found

              # read complete header (inkl. file data) into string

              $inputStream = $RequestItem.InputStream
              $Encoding = $RequestItem.ContentEncoding


              $READER = New-Object System.IO.StreamReader($inputStream, $Encoding)
              $DATA = $READER.ReadToEnd()
              $READER.Close()
              $RequestItem.InputStream.Close()

              # }
              Write-Host "Request Data:"
              Write-Host $DATA
              # TODO Prepare response Script

              # $JSONRESPONSE = AccessJSON $app "Test"
              $JSONRESPONSE = AccessCmd $app "DbMsg" "Test Messagebox"

              $HRes.AddHeader("Content-Type","text/json")
              $HRes.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
              $HRes.AddHeader("Server", "Powershell Webserver/1.2 on ")

              # return HTML answer to caller
              $BUFFER = [Text.Encoding]::UTF8.GetBytes($JSONRESPONSE )
              $HRes.ContentLength64 = $BUFFER.Length
              $HRes.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
            }
          }
        }
      break
      }
    }
    $HRes.Close()
  }
}finally{
  #Close Listener
  $Hso.Stop()
  $Hso.Close()

  #Close MS ACCESS
  # CloseDb $app

  Write-Host "Closed"
}
