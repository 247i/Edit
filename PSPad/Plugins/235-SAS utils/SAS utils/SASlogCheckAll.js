/*
*******************************************************************************

PROGRAM NAME:
SASlogCheckAll.js

DESCRIPTION:
Check all SAS log files in the current directory for errors, warnings, or
relevant notes. Results will be written to log_errors.txt, log_warnings.txt and 
log_notes.txt respectively. Empty files will be created if no errors, warnings,
notes are present

PROGRAMMER:
John Hendrickx <John2.Hendrickx@gmail.com>

*******************************************************************************
*/
var module_name  = "SASlogCheckAll";
var module_ver   = "0.1";

var dir, editor;

function Init(){
	addMenuItem("SASlogCheckAll", "SAS", "SASlogCheckAll","CTRL+ALT+SHIFT+L");
}

function SASlogCheckAll() {
	logClear();
	dir=getVarValue("%FilePath%");
	editor=newEditor();
	CloseOpenFiles();
 	editor.newfile(dir+"log_errors.txt");
 	editor.newfile(dir+"log_warnings.txt");
 	editor.newfile(dir+"log_notes.txt");

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	f = fso.GetFolder(dir);
	fc = new Enumerator(f.files);
	var currfile
	for (; !fc.atEnd(); fc.moveNext()) {
		if (!/\.log$/i.test(fc.item())) continue;
		currfile=fso.GetFile(fc.item());
		doCheck(currfile);
	}
 	editor.assignEditorByName(dir+"log_notes.txt");
 	editor.saveFile();
 	editor.assignEditorByName(dir+"log_warnings.txt");
 	editor.saveFile();
	editor.assignEditorByName(dir+"log_errors.txt");
 	editor.saveFile();
 	editor.assignActiveEditor();
 	if (fso.FileExists(dir+"log_errors.txt.bak")) fso.DeleteFile(dir+"log_errors.txt.bak");
 	if (fso.FileExists(dir+"log_warnings.txt.bak")) fso.DeleteFile(dir+"log_warnings.txt.bak");
 	if (fso.FileExists(dir+"log_notes.txt.bak")) fso.DeleteFile(dir+"log_notes.txt.bak");
}

function doCheck(thisfile) {
	var errorfree=1;
	var warningfree=1;
	var notefree=1;
	var filecnt=thisfile.OpenAsTextStream(1, -2); // for reading, using system default for unicode/ascii
	var currline, linenr;
	while (!filecnt.AtEndOfStream) {
		currline=filecnt.readline();
		if (/^ERROR:/.test(currline)) {
			linenr=filecnt.Line-1;
			editor.assignEditorByName(dir+"log_errors.txt");
			if (errorfree) {
				print_prog_name(thisfile.Name);
				errorfree=0;
			}
			editor.appendText(padnum(linenr)+" "+currline+"\n");
			//editor.assignEditorByName(thisfile);
		}
		if (/^WARNING:/.test(currline)) {
			linenr=filecnt.Line-1;
			editor.assignEditorByName(dir+"log_warnings.txt");
			if (warningfree) {
				print_prog_name(thisfile.Name);
				warningfree=0;
			}
			editor.appendText(padnum(linenr)+" "+currline+"\n");
			//editor.assignEditorByName(thisfile);
		}
		if (/^NOTE:/.test(currline)) {
			if ( (/not referenced/i.test(currline)) ||
				(/never been refere/i.test(currline)) ||
				(/not resolved/i.test(currline)) ||
				(/is not in the report def/i.test(currline)) ||
				(/repeats of by values/i.test(currline)) ||
				(/current word or quoted string has become more/i.test(currline)) ||
				(/outside the axis range/i.test(currline)) ||
				(/is unknown/i.test(currline)) ||
				(/invalid/i.test(currline)) ||
				(/was not found, but appears on a delete/i.test(currline)) ||
				(/uninitialized/i.test(currline)) ||
				(/division by zero/i.test(currline)) ||
				(/w.d format/i.test(currline)) ||
				(/mathematical operations could not/i.test(currline)) ||
				(/cannot be determined/i.test(currline)) ||
				(/has been truncated/i.test(currline)) ||
				(/extraneous information/i.test(currline)) ||
				(/could not be loaded/i.test(currline)) ||
				(/groups are not created because/i.test(currline)) ||
				(/due to looping/i.test(currline)) ||
				(/end of macro/i.test(currline)) ||
				(/ending execution/i.test(currline)) ||
				(/errorabend/i.test(currline)) ||
				(/errorcheck=strict/i.test(currline))  )
			{
				linenr=filecnt.Line-1;
				editor.assignEditorByName(dir+"log_notes.txt");
				if (notefree) {
					print_prog_name(thisfile.Name);
					notefree=0;
				}
				editor.appendText(padnum(linenr)+" "+currline+"\n");
				//editor.assignEditorByName(thisfile);
			}
		} // if /^NOTE/
	}
	filecnt.Close();
}

function print_prog_name(logname) {
	var log=logname;
	var re=/\.log$/i;
	var progname=log.replace(re,".sas");
	editor.appendText("\n########################################################\n");
	editor.appendText("### PROGRAM: "+progname+"\n");
	editor.appendText("########################################################\n\n");
}

function CloseOpenFiles() {
	var fname;
	for (i=0; i < editorsCount(); i++) {
		editor.assignEditorByIndex(i);
		fname=editor.FileName();
		if (/log_(errors|warnings|notes)\.txt$/i.test(fname)) closeEditorByIndex(i);
	}
}

function padnum(n) {
	var numstr='      ' + n.toString();
	return numstr.substr(numstr.length-6);
}
