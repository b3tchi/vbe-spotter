#!/snap/bin/powershell

# Import-Module $PSScriptRoot/AccessRunDb.ps1
Import-Module $PSScriptRoot/Config.ps1

$app = GetApp $dbFullPath

function request(){
  param(
    [string]$headers
    ,[string]$data
  )

  # Write-Host 'headers:'
  # Write-Host $headers
  # Write-Host 'data:'
  # Write-Host $data
  # Write-Host 'end'

  $respAr = $app.Run("HttpRequest", [ref]$headers, [ref]$data) #use [ref] for optional COM parameters

}
