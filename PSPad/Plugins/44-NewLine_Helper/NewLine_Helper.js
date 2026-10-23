/*****************************************************************************
* This software is provided 'as-is', without any express or implied warranty.
* In no event will the authors be held liable for any damages arising from
* the use of this software.
*
* Permission is granted to anyone to use this software for any purpose,
* including commercial applications, and to alter it and redistribute it freely.
*
********************************************************************************/

/*
paste the following line into a new text file
(\S+),~#LF#$1#LF#This is line1,This is line2,The last line
then select everything up to the first "This" (\S+),~#LF#$1#LF#
the tilde(~) serves as the argument delimiter    */

function NewLine_Helper() {
var startStr, findStr, newStr, temp;
var ed = newEditor();
ed.assignActiveEditor();
  if(ed.selText() == '') return;
  
// if you need to change the argument delimiter edit the next line
temp = ed.selText().split("~");
  
  if(temp.length != 2) return;
findStr = temp[0];
newStr = temp[1];
  if(findStr.length == 0 || newStr.length == 0) return;
ed.selText('');

/*    Had trouble with literal string interpretation.
      You may need to add your own conversion below
      if you run across issues.
                                                          */

  newStr = newStr.replace(/#LF#/g, "\n");
  newStr = newStr.replace(/#SP#/g, "\s");
  newStr = newStr.replace(/#TAB#/g, "\t");

startStr = ed.Text();
var rex = new RegExp(findStr,'ig');
var newStr = startStr.replace(rex, newStr);
ed.text(newStr);
}

function Init(){
addMenuItem("New Line Helper", "", "NewLine_Helper");
}
