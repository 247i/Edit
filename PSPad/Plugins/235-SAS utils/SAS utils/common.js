/*
*******************************************************************************

PROGRAM NAME:
common.js

DESCRIPTION:
Contains functions called by other scripts

PROGRAMMER:
John Hendrickx <John2.Hendrickx@gmail.com>

*******************************************************************************
*/
function date9() {
	var today = new Date();
	var d="0"+today.getDate();
	var dd=d.substr(d.length-2);
		
	var monthname=new Array("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG",
		"SEP","OCT","NOV","DEC");
	var mmm=monthname[today.getMonth()];

	var date9=dd+mmm+today.getFullYear();
	return(date9);
}

//http://www.eggheadcafe.com/community/aspnet/3/10007603/remove-duplicates.aspx
function RemoveDuplicates(arr) {
	// sort the array, return without duplicates.
	arr=arr.sort(function(x,y){ 
		var a = String(x).toUpperCase(); 
		var b = String(y).toUpperCase(); 
		if (a > b) 
			return 1 
		if (a < b) 
			return -1 
		return 0; 
    }); 
	var result=new Array();
	var lastValue="";
	var curValue;
	for (var i=0; i<arr.length; i++) {
		curValue=arr[i];
		if (curValue.toUpperCase() != lastValue.toUpperCase()) {
			result[result.length] = curValue;
		}
		lastValue=curValue;
	}
	return(result);
}

function noMacrosComments(instring) {
	// use [\s\S] to match anything, including newlines, for a multiline search
	// use *? for a non-greedy search, i.e. match the first "*/" occurrence
	// First remove all multiline comments /* ... */
	var re=/\/\*[\s\S]*?\*\//ig;
	var nocomm=instring.replace(re,"");
	
	// remove all single line comments
	var re=/(^|\s+)\*[\s\S]*?;/ig;
	var nocomm2=nocomm.replace(re,"");
	
	// remove all in-file macro specifications
	var re=/%macro[\s\S]*?%mend[\s\S]*?;/ig;
	var nocomm3=nocomm2.replace(re,"");
	return(nocomm3);
}

function noComments(instring) {
	// use [\s\S] to match anything, including newlines, for a multiline search
	// use *? for a non-greedy search, i.e. match the first "*/" occurrence
	// First remove all multiline comments /* ... */
	var re=/\/\*[\s\S]*?\*\//ig;
	var nocomm=instring.replace(re,"");
	
	// remove all single line comments
	var re=/(^|\s+)\*[\s\S]*?;/ig;
	var nocomm2=nocomm.replace(re,"");
	return(nocomm2);
}
