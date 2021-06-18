$jscriptCode = {
    function Create-ScriptEngine()
    {
      param([string]$language = $null, [string]$code = $null);
      if ( $language )
      {
        $sc = New-Object -ComObject ScriptControl;
        $sc.Language = $language;
        if ( $code )
        {
          $sc.AddCode($code);
        }
        $sc.CodeObject;
      }
    }
$jscode = @"
function main(s)
{
    return 'a'
}
"@
$js = Create-ScriptEngine "JScript" $jscode;
$str = "abcd";
$js.main($str);

}

$job = Start-Job -ScriptBlock $jscriptCode -runAs32 
$output = $job | Wait-Job | Receive-Job
$output
Remove-Job $job