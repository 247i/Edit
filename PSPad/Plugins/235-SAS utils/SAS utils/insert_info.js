/*
*******************************************************************************

PROGRAM NAME:
insert_info.js

DESCRIPTION:
Insert date: insert the current date as ddmmmyyyy (e.g. 20MAY2009)
Insert file name: insert the file name, e.g. insert_info.js
Insert file path: insert the path to the current file 
Report Macros: A list of SAS macros will be written to the PSPad log window
and copied to the clipboard

PROGRAMMER:
John Hendrickx <John2.Hendrickx@gmail.com>

*******************************************************************************
*/
function Init(){
	addMenuItem("Insert date", "Insert Info", "ins_date","CTRL+SHIFT+D");
	addMenuItem("Insert file name", "Insert Info", "ins_fn","CTRL+SHIFT+N");
	addMenuItem("Insert file path", "Insert Info", "ins_path","CTRL+ALT+N");
	addMenuItem("Report Macros", "Insert Info", "ReportMacros");
}

function ins_fn() {
	editor=newEditor();
	editor.assignActiveEditor(); 
	editor.SelText(getVarValue("%FileName%")); 
}

function ins_path() {
	editor=newEditor();
	editor.assignActiveEditor(); 
	editor.SelText(getVarValue("%FilePath%")); 
}

function ins_date() {
	editor=newEditor();
	editor.assignActiveEditor(); 
	editor.SelText(common.date9()); 
}


function ReportMacros() {
	logClear();
	var editor = NewEditor();
	editor.assignActiveEditor(); 
	var alltxt = editor.text();
	
	var alltxt3=noMacroLabels(alltxt);
	
	// search for all macro sytnax
	var re=/%\w+/ig;
	var mtch=new Array();
	mtch=alltxt3.match(re);
// 	logaddline(mtch);
	if (mtch==null) return("");
	
	// sort and remove duplicates
	var mtch2=common.RemoveDuplicates(mtch);
// 	logaddline(mtch2);

	// transform to string, remove builtin macro words that can be used outside
	// a macro
	var mtchstr=mtch2.toString();
// 	logaddline(mtchstr);
	var rmv=/%(eval|global|q?left|length|let|q?lowcase|inc(lude)?|put|(nr)?b?quote|q?scan|(nr)?q?str|q?substrn?|superq|q?sysfunc|q?trim|q?upcase)($|,)/ig;
	var mtchstr2=mtchstr.replace(rmv,"");
// 	logaddline(mtchstr2);
	var rmv2=/%(by|do|else|end|goto|if|index|length|let|local|macro|mend|put|return|symdel|symexist|sysevalf|then|to|while|until)($|,)/ig;
	var mtchstr3=mtchstr2.replace(rmv2,"");
	var mtchstr4=mtchstr3.replace(/,$/,"");

	setClipboardText(mtchstr4);
	logaddline(mtchstr4);
}

function noMacroLabels(instring) {
	// use [\s\S] to match anything, including newlines, for a multiline search
	// use *? for a non-greedy search, i.e. match the first "*/" occurrence
	// remove all macro labels (e.g. %exit:)
	var re=/%\w+?:/ig;
	var nomaclab=instring.replace(re,"");
	
	return(nomaclab);
}

