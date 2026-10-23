// JavaScript Document
//*******************************************************************************
//	filename   : swap.js
//	description: Swap characters from left and right sides of cursor. For correct accidentally typed.
//	created    : 15.11.2006
//	author     : Serge Balance
//
//	You may distribute this script freely, but please keep this header intact.
//*******************************************************************************
var MODULE_NAME  = "Swap characters";
var MODULE_VER   = "0.7";
var MODULE_TITLE = "Swap characters from left and right sides of cursor";
function Init() {
    menuName = "&" + MODULE_NAME;
    subMenu = "&" + "Edit";
    addMenuItem(menuName, subMenu, "main", "CTRL+ALT+K");
}

function main()
{
	var	ed = neweditor();
	ed.assignActiveEditor();
	var text = new String(ed.lineText());
	var curx = ed.caretX();
	var leftsymb = new String(text.substr(curx-2,1));
	ed.command('ecDeleteLastChar');
	ed.command('ecRight');
//	ed.command('ecInsertMode');
	ed.command('ecInsertLine');
	ed.command('ecLineBreak');
	ed.lineText(leftsymb);
	ed.command('ecLineEnd');
	ed.command('ecDeleteChar');
	ed.command('ecLineStart');
	ed.command('ecDeleteLastChar');
// 	ed.command('ecChar{leftsymb}'); // not working at all
//	echo(leftsymb);
}

function indexInText(text, curx, cury) { //function for lib WSH-PSPad
// calculate linear coordinate in String text, based on cursor coordinates 
var i = 1; // from first line, not zero
var pos = 0;
var cursymb = new String();
while (i < cury) {
	pos = text.indexOf("\n",pos);
	pos++; // next search from next position
	i++;
}
cursymb = text.charAt(pos);
if ((cursymb == "\n") || (cursymb == "\r") || (cursymb == "\f")) {
	pos++;
}
pos = pos + curx - 1; // Array index from 0, string index from 1
return (pos);
}
