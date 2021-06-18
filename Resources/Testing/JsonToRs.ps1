
Import-Module ./AccessRunDb.ps1

# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()

# $scriptPath = $PSScriptRoot
# $scriptPath = Split-Path -Parent $PSCommandPath
# $scriptPath = "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps"
# $shelperName = "shelper.accdb"

# $shelperPath = "$scriptPath\$shelperName"
$dbFullPath = "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

# $jsonS = "{""parameters"":{""command"":""TestCommand""}}"

$jsonS = @"
{
"data":[
    {"01_Items":[
        {"FormID":1,"FormAction":1,"ItemID":1,"Field1":"Test"}
        ,{"FormID":2,"FormAction":1,"ItemID":2,"Field1":"TestA"}
        ]
    }
    ,{"01_Assignment":[
        {"FormID":1,"FormAction":1,"ItemID":1, "AssignmentID": 1}
        ,{"FormID":2,"FormAction":1,"ItemID":2, "AssignmentID": 2}
        ]
    }
    ]

,"parameters":
    {"command":"TestCommand"
    ,"option1":"value1"
    }

}
"@

$json = $jsonS | ConvertFrom-Json

# write-host $json[1] #get first object in array
# write-host $json."data" #get first object in array

$data = $json."data" #get first object in array

# write-host $data

$app = GetApp $dbFullPath # $shelperPath

#GetDb
$db = $app.CurrentDb()

foreach($item in $data){
    
    #SPLIT HERE TO FUNCTION
    ConvertToRs $db $item

}
#END OF SCRIPT
