/*
*******************************************************************************

PROGRAM NAME:
SASlogCheckAll.js

DESCRIPTION:
Check a SAS log file for errors, warnings, or relevant notes. Results will be 
written to the PSPad log window.

PROGRAMMER:
John Hendrickx <John2.Hendrickx@gmail.com>

*******************************************************************************
*/
var module_name  = "SASlogCheck";
var module_ver   = "0.1";

function Init(){
	addMenuItem("SASlogCheck", "SAS", "SASlogCheck","CTRL+ALT+L");
}

function SASlogCheck() {
	var currFile = newEditor();
	currFile.assignActiveEditor();
	var logname=currFile.FileName();
	if (!/\.log$/i.test(logname)) return;
	logClear();
	var firstlog=0;

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var thisfile=fso.GetFile(logname);
		
	var filecnt=thisfile.OpenAsTextStream(1, -2); // for reading, using system default for unicode/ascii
	var currline, LogLine;
	while (!filecnt.AtEndOfStream) {
		currline=filecnt.readline();
		LogLine=(/^ERROR|^WARNING/.test(currline));

		if (/^NOTE/.test(currline)) {
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
					LogLine = "True";
			}
		} // if /^NOTE/
		
			 		
		if (LogLine) {
			var linenr=filecnt.Line-1;
			if (firstlog == 0) firstlog = linenr;
			logAddLine(thisfile.Name + "(" + linenr + "): " + currline);
		}
	} // end of for-loop
	
	if (firstlog == 0) {
		logAddLine("No errors, warnings, or relevant notes");
		currFile.setCaretPos(1, 1);
	} else {
		currFile.setCaretPos(1, firstlog);
	}
}