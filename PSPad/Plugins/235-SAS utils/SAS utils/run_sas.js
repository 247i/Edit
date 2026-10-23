/*
*******************************************************************************

PROGRAM NAME:
run_sas.js

DESCRIPTION:
Run a SAS program in batch mode from PSPad and on completion, ppen the log 
(and listing file, if created) or, if already opened in PSPad, refresh them.

PROGRAMMER:
John Hendrickx <John2.Hendrickx@gmail.com>

*******************************************************************************
*/
// Note: the "-nologo" and "-icon" options are strognly recommended for the 
// sas_options variable. If they are not used, PSPad will prompt the to reload 
// the LOG and LST files after running SAS. Using "-nologo" and "-icon", these 
// files will be automatically reloaded.
// Add other options like -autoexec here
var sas_options="-nologo -icon ";

function Init(){
	addMenuItem("Run SAS", "SAS", "run_sas","CTRL+SHIFT+R");
}

function run_sas() {
	logClear();
	var editor = NewEditor();
	editor.assignActiveEditor(); 
	editor.saveFile();
	var fname = editor.fileName();
	var re=/\.sas$/i;
	if (!re.test(fname)) return;
	outfile=fname.replace(re,".lst");
	logfile=fname.replace(re,".log");
// 	logAddLine(outfile);

	var WshShell = new ActiveXObject("WScript.Shell"); 
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var dir=getVarValue("%FilePath%");
	WshShell.CurrentDirectory=dir;
// 	logAddLine(WshShell.CurrentDirectory);
	
	var SAS_ver = WshShell.RegRead("HKEY_LOCAL_MACHINE\\SOFTWARE\\SAS Institute Inc.\\The SAS System\\CurrentVersion\\CurrentVersion");
// 	logAddLine(SAS_ver);
	var SAS_path = WshShell.RegRead("HKEY_LOCAL_MACHINE\\SOFTWARE\\SAS Institute Inc.\\The SAS System\\" + SAS_ver + "\\DefaultRoot");
// 	logAddLine(SAS_path);

	var run_cmd="\""+SAS_path+"\\sas.exe\" -sysin \""+fname+"\""+sas_options;
		
// 	logAddLine(run_cmd);

	/*
	0: intWindowStyle=Hides the window and activates another window. 
	1:bWaitOnReturn=true, wait for the program to finish executing
	*/
	WshShell.run(run_cmd,0,1);
// 	logAddLine("finished");
	
// Check if the LST file is already open in PSPad. If so, reload it. If not, open 
// it if it exists.
	var qtab2=out_index(outfile);
// 	logAddLine(qtab2);
	if (qtab2==-1) {
		if (fso.FileExists(outfile)) {
			editor.newFile();
			editor.openFile(outfile);
		}
	}
	else {
		editor.assignEditorByIndex(qtab2);
		editor.activate();
		editor.reloadFile();
	}	

// Check if the LOG file is already open in PSPad. If so, reload it. If not, open 
// it.
	var qtab1=out_index(logfile);
// 	logAddLine(qtab1);
	if (qtab1==-1) {
		editor.newFile();
		editor.openFile(logfile);
	}
	else {
		editor.assignEditorByIndex(qtab1);
		editor.activate();
		editor.reloadFile();
	}
}

function out_index(name) {
	var tmped = NewEditor();
	tmped.assignActiveEditor(); 
// 	logAddLine(name);
// 	logAddLine(editorsCount());
	var currfile;
	for (i=0;i<editorsCount();i++) {
		tmped.assignEditorByIndex(i);
		currfile=tmped.fileName();
// 		logAddLine(i+currfile);
		if (currfile.toUpperCase()==name.toUpperCase()) return(i);
	}
	return(-1);
}

