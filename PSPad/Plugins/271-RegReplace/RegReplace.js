/*******************************************************************************
 * (c) 2010 - http://www.hofrichter.net - sven hofrichter - pspad@hofrichter.net
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * author ......... : Sven Hofrichter
 * date ........... : 23.08.2010
 * version ........ : 1.1
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * description .... : This script was made to enable the support of regulare
 *                    expressions, while replacing newline-sequences in multi-
 *                    lined text. It was made, because of the missing feature
 *                    in the standard-dialog of PSPad and because of the
 *                    admin-rights-riquirering installation of a DLL by the
 *                    other wonderful userscript 'phReplace'. If you don't have
 *                    the administration rights or if you don't want to install
 *                    DLLs you can you this little script ;-)  
 * introduction ... : The script has to have located in:
 *                    - <PSPad-path>\Script\JScript
 *                    And then you can replace text, by doing this 
 *                    1. press <CTRL> + <SHIFT> + <R>
 *                    2. type in your regulare expression for substitution text.
 *                       These standard-pattern are supported by this script: 
 *                       - /.../.../
 *                       - s/.../.../
 *                       AND these optionial flags are supported:
 *                       - g, m, i -> example: s/.../.../gim
 *                    3. press ok and the text will be replaced!
 * additional informations:
 * At the first view lines, you find some variables. You can change them, 
 * however you want.
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * history:
 * 20.08.2010 - 1.0 - hofsve - initial version
 * 23.08.2010 - 1.1 - hofsve - added linebreak-support ("\r" and/or "\n") for
 *                    the "target-pattern". 
 ******************************************************************************/

// you can change this:
var SHORT_CUT_KEYS = 'CTRL+SHIFT+R';
var showConfirmDialog  = true; // Needs access-rights to an IE-library. If you
                               // don't like it: disable this feature. The IE-
                               // LIB will ONLY BE USED, if you want to get a
                               // confirm-dialog
var confirmDialogLimit = 100;  // 'n' characters will be shown in the confirm-
                               // dialog

//------------------------------------------------------------------------------
// DO NOT CHANGE THIS:
var module_name = "RegReplace";
var module_ver = "1.1";
function regReplace() {
	var editor = newEditor();
	var text = '';
	var isWholeFile = false;
	editor.assignActiveEditor();
	if (editor.selText().length > 0) {
	  isWholeFile = false;
		text = applyRegularExpression(editor.selText(), isWholeFile);
	} else {
	  isWholeFile = true;
		text = applyRegularExpression(editor.Text(), isWholeFile);
	}
	if (text != null) {
  	if (isWholeFile) {
      editor.Text(text);
    } else {
      editor.selText(text);
    }
  }
}

function applyRegularExpression(selectedText, isWholeFile) {
	var regulareExpression  = null;
	var origRegulareExpression = null;
  try {
	  var defaultExpression = "s/" + (isWholeFile ? '' : selectedText) + "//g";
		// if the version of PSPad is not a newer one, the following will fail:
		regulareExpression = InputText( "Please enter the regulare expression:", defaultExpression);
		if (regulareExpression == defaultExpression || regulareExpression == '') {
		  echo("Replacing text cancled!");
      return;
    }
	}
	catch(e) {
		return null;
	}
	
	origRegulareExpression = regulareExpression;
	var separator = regulareExpression.replace(/^\w*([^\w]).*/, "$1"); // find the separator-symbol
	if (separator != '/') {
    echo("Only slash '/' is supported as separator for search- or -replace-expressions!");
    return null;
  }
	var regPattern = new Array();
  //regPattern = regulareExpression.split(new RegExp(separator));
  var endPosition = regulareExpression.indexOf(separator);
  var startPosition = 0; 
  while (endPosition != -1) {
    var tmpResult = regulareExpression.substring(startPosition, endPosition);
    regPattern.push(tmpResult == separator ? '' : tmpResult);
    regulareExpression = regulareExpression.substring(endPosition + 1, regulareExpression.length);
    endPosition = regulareExpression.indexOf(separator);
    startPosition = 0; 
  }
	if (regPattern == null || regPattern.length == 0) {
    echo("Couldn't translate your expression into a search- or -replace-expression!");
    return null;
  }
  var tmpResult = regulareExpression.substring(startPosition, regulareExpression.length);
  regPattern.push(tmpResult == separator ? '' : tmpResult);
  if (regPattern.length != 4) {
    echo("An error occured while interprating the search- or -replace-expression!\nWrong number for extracted pieces (expected 3 or 4):\n- " + regPattern.join('\n- '));
    return null;
  }

  // now we start trieing to work with the pattern
  regPattern[0] = regPattern[0].toLowerCase();
  regPattern[3] = regPattern[3].toLowerCase();
  if (regPattern[0] != '' && regPattern[0] != 's') {
    echo("Wrong prefix '" + regPattern[0] + "' in pattern!Expecting -nothing- or 's'\n>>> Supported patterns: 's/.../.../' -or- '/.../.../' (plus optional pattern-parameters 'g', 'i', 'm', 's')");
    return null;
  }
  var errorTextArray = new Array();
  for (var i = 0; i < regPattern[3].length; i++) {
     var subText = regPattern[3].substring(i, i + 1);
     if (!subText.match(/[gim]/)) {
        errorTextArray.push(subText);
     }
  }
  //if (regPattern[3] != '' && regPattern[3] != 'g' && regPattern[3] != 'i' && regPattern[3] != 'm' && regPattern[3] != 's') {
  if (errorTextArray.length > 0) {
    echo("Wrong suffix (optional parameters) in pattern: '" + errorTextArray.join("', '") + "'!\nExpected syntax: 'g', 'i' -or- 'm'");
    return null;
  }
  
  var nlStr = '';
  var newLineChar = '';  
  if (regPattern[1].match(/(\\n|\\r|\\r\\n|\\n\\r)/g) || regPattern[2].match(/(\\n|\\r|\\r\\n|\\n\\r)/g)) {
    nlStr = '#N#';
    while(selectedText.match(new RegExp(nlStr))) {
      nlStr += nlStr;
    }
    while (regPattern[2].match(new RegExp(nlStr))) {
      nlStr += nlStr;
    }
    var treffer = ''; //selectedText.match(/([\n\r])/g);
    var treffer2= regPattern[2].match(/([\\n\\r])/g);
    newLineChar = treffer && treffer.length > 0 ? treffer[0] : '';
    if (newLineChar == '') {
      var treffer2= regPattern[2].match(/([\\n\\r])/g);
      var tempTreffer2 = treffer2[0].replace(/[\\n]+/g, "\n").replace(/[\\r]+/g, "\r");
      newLineChar = tempTreffer2 && tempTreffer2.length > 0 ? tempTreffer2 : '';
    }
    regPattern[1] = regPattern[1].replace(/(\\n|\\r|\\r\\n|\\n\\r)/g, nlStr);
    regPattern[2] = regPattern[2].replace(/(\\n|\\r|\\r\\n|\\n\\r)/g, nlStr);
    selectedText  = selectedText.replace(/[\n\r]+/g, nlStr);
  }
	var mRegExp = null;
  if (regPattern[3] != '') {
    mRegExp = new RegExp(regPattern[1], regPattern[3]);
	} else {
    mRegExp = new RegExp(regPattern[1]);
  }
  selectedText = selectedText.replace(mRegExp, regPattern[2]);
	if (nlStr != '') {
	  var replaceBack = new RegExp(nlStr, 'gm');
	  while (selectedText.match(replaceBack)) {
      selectedText = selectedText.replace(replaceBack, newLineChar);
    }
  }
   
  if (confirm(origRegulareExpression, selectedText)) {
    return selectedText;
  }
  return null;
}
function confirm(origRegulareExpression, newText) {
  if (!showConfirmDialog) {
    return true;
  }
  var limit = confirmDialogLimit > newText.length ? newText.length : confirmDialogLimit; 
  var text  = "Do you want to replace the current Text by this one?";
      text += "\n--------------"; 
      text += "\nYour Pattern: " + origRegulareExpression;
      text += "\n--------------"; 
      text += "\nReplaced text:"; 
      text += "\n\n" + newText.substring(0, limit); 
  var ie = CreateObject("InternetExplorer.Application"); 
  ie.width = 0;
  ie.height = 0;
  ie.left = 10000000;
  ie.visible = true;
  ie.navigate("about:blank");
  var returnValue = ie.Document.parentWindow.confirm(text); 
  ie.quit(); 
  return returnValue; 
}
function Init() {
	addMenuItem( "RegReplace (v" + module_ver + ")", "", "regReplace", SHORT_CUT_KEYS);
}
