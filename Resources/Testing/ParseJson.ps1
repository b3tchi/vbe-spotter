$jsonS = "{""name"":""getItems"",""keys"":[]}"

$json = $jsonS | ConvertFrom-Json

$json | Add-Member -NotePropertyName Status -NotePropertyValue Done

$json2 = $json | ConvertTo-Json

Write-Host $json2