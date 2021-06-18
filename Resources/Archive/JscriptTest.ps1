#ScriptBlock: $var = {...code...}
$jscriptCode = {
  
  function ScriptEngine() {
 
    $sc = New-Object -ComObject MSScriptControl.ScriptControl.1;
    $sc.Language = "JScript";
    $sc.AddCode($args[0]);
    $sc.CodeObject;
  } 

  # write-host "There are a total of $($args.count) arguments"
  
  $jscode = $args[0];
  $jsoncode = $args[1];
  $js = ScriptEngine $jscode;

  # $str = "abcd";
  # $js.eval($jsoncode);
  write-host $js.main($jsoncode);
  # write-host $js.main($jsoncode);


}# -runas32 | wait-job | receive-job

$jscode = Get-Content -Path .\Testing\jScriptLocal.js -Raw
$jsoncode = Get-Content -Path .\Testing\json2.js -Raw
# $jscode = Get-Content -Path .\Testing\jsTest2.js 
# $jscode = "function mainx(s){ return 'a'; }"

# write-host $jscode

# run in isolated 32bit session
$job = Start-Job -ScriptBlock $jscriptCode -runAs32 -ArgumentList @($jscode, $jsoncode)
$output = $job | Wait-Job | Receive-Job
# $output.GetType() #return output
$output
Remove-Job $job

# $jscode = @"
# function jslen(s)
# {
#   return 'a'
# }
# "@

# $jscode = @"
# function jslen(s)
# {
#   getObject()
#   return s.length;

#   var Fs = new ActiveXObject("Scripting.FileSystemObject");
#   eval(Fs.OpenTextFile("json2.js", 1).ReadAll());

#   var shell = new ActiveXObject("WScript.Shell");

#   var myObj = {name: "John", age: 31, city: "New York"};
#   var myJSON = JSON.stringify(myObj);

#   shell.Popup(myJSON);
# }
# "@

# "Done"