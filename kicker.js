// FileName: kicker.js

var baseFolder = "C:\\\\temp";
var subFolder = "updater-alpha";

var scripts = [
//    "sample1.js",
    "JRE_Updater.vbs",
];


var fs = new ActiveXObject("Scripting.FileSystemObject");
for ( var i = 0; i < scripts.length; i++ ) {
    var file = baseFolder + "\\" + subFolder + "\\" + scripts[i];
    if (fs.FileExists(file) == true) {
        var wsh = WScript.CreateObject ("WScript.Shell");
        wsh.Run(file);
        wsh = null;
    }
}
fs = null;
