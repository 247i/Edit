var module_name = "SaveCopy";
var module_ver = "0.001a";

function SaveCopy()
{	
  var thisdoc = newEditor();
  thisdoc.assignActiveEditor();
  var newdoc="";
  var dat = thisdoc.Text();
  
  var newdoc = newEditor();
  newdoc.newFile();
  newdoc.Text(dat);
  newdoc.fileName("");
    
	var wsh = new ActiveXObject("WScript.Shell");
	wsh.sendKeys("^{F4}");
	wsh.sendKeys("~");
}

function Init(){
  addMenuItem("Save Copy", "", "SaveCopy");
}
