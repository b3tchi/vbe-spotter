

// run();


// function includeFile (filename) {
//     var fso = new ActiveXObject ("Scripting.FileSystemObject");
//     var fileStream = fso.openTextFile (filename);
//     var fileData = fileStream.readAll();
//     fileStream.Close();
//     eval(fileData);
// }
function run() {

    var Fs = new ActiveXObject("Scripting.FileSystemObject");
    var jsonlib = Fs.OpenTextFile("json2.js", 1).ReadAll()
    var shell = new ActiveXObject("WScript.Shell");
    
    // main();
// 

    shell.Popup(main(jsonlib));
}

function main(jsonlib) {
        
    eval(jsonlib);

    // var myObj = {name: "John", age: 31, city: "New York"};

    var db = GetObject("C:\\Users\\czJaBeck\\Documents\\Vbox\\LocalWeb_Ps\\TestDb.accdb");
    // var Workbook = GetObject('', 'Microsoft.Access')
    // var db = GetObject('', "Access.Application");

    // var shell = new ActiveXObject("WScript.Shell");

    // shell.Popup(myJSON);
    var dbs = db.CurrentDb();

    var rs = dbs.OpenRecordset("AllItems");

    rs.MoveLast();
    rs.MoveFirst();

    var data = [];
    var fldCount = rs.Fields.Count;
    var rcount = rs.RecordCount;

    // var recTemplate = {};
    // for (var index = 0; index < fldCount; index++) {
    //     var fldName = rs.Fields(index).Name;
    //     recTemplate[fldName] = null;
    // }
    // shell.Popup(JSON.stringify(recTemplate));

    while(rs.EOF != true){

        var rec = {};

        //add recordset
        for (var i2 = 0; i2 < fldCount; i2++) {

            var fldName = rs.Fields(i2).Name;
            rec[fldName] = rs.Fields(i2).Value;

        }

        data.push(rec);
        // shell.Popup(rs.Fields(1).Value)
        rs.MoveNext();

    }

    // shell.Popup(JSON.stringify(data));

    rs.close();
    dbs.close();

    var respText = JSON.stringify(data);

    return respText.toString();
    // return 'a';

}
