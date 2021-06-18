function ConvertFrom-Xml {
  param([parameter(Mandatory, ValueFromPipeline)] [System.Xml.XmlNode] $node)
  process {
    if ($node.DocumentElement) { $node = $node.DocumentElement }
    $oht = [ordered] @{}
    $name = $node.Name
    if ($node.FirstChild -is [system.xml.xmltext]) {
      $oht.$name = $node.FirstChild.InnerText
    } else {
      $oht.$name = New-Object System.Collections.ArrayList 
      foreach ($child in $node.ChildNodes) {
        $null = $oht.$name.Add((ConvertFrom-Xml $child))
      }
    }
    $oht
  }
}

# $xmlObject = Get-Content -Path C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\Table1.xml | ConvertTo-Xml 
$rawText = Get-Content -Path C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\Table1.xml

$xmlObject = [XML]($rawText)

$jsonObject = $xmlObject | ConvertFrom-XML | ConvertTo-JSON -Depth 3

Write-Output $jsonObject

