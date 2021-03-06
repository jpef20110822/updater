// FileName: download.js

var url = "https://github.com/jpef20110822/updater/archive/master.zip";
var file = "master.zip";

var baseFolder = ".";
var subFolder = "updater-master";

var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject('Shell.Application');

// download file
try {
    var http = new ActiveXObject("Msxml2.ServerXMLHTTP");
    http.open("GET", url, false);
    http.send();
    if (http.status == "200") {
        var strm = WScript.CreateObject("Adodb.Stream");
        var adTypeBinary = 1, adSaveCreateOverWrite = 2;
        strm.Type = adTypeBinary;
        strm.Open();
        strm.Write(http.responseBody);
        strm.Savetofile(baseFolder + "\\" + file, adSaveCreateOverWrite);
    }
} catch(e) {
    throw e;
}


// delete old folder
if (fso.FolderExists(baseFolder + "\\" + subFolder) == true) {
    fso.DeleteFolder(baseFolder + "\\" + subFolder);
}

// unzip and copy
var dst = shell.NameSpace(fso.getFolder(baseFolder).Path);
var zip = shell.NameSpace(fso.getFile(baseFolder + "\\" + file).Path);
dst.CopyHere(zip.Items(), 4 + 16);
