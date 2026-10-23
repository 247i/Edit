/*------------------------------------------------------------
/ New Cpp class
/
/   Copyright (C) 2012 Ales Cernigoj
/   This program is free software: you can redistribute it and/or modify
/   it under the terms of the GNU General Public License.
/
/   TODO: Get some meaningful path into path input box.
/
/------------------------------------------------------------*/
var module_name = "New Cpp class";
var module_ver = "0.2";

/*------------------------------------------------------------
    Flags
/------------------------------------------------------------*/
var restoreLastEntry = true;   // if true, will try to restore last entry.

function Init()
{
    addMenuItem("New Cpp class", "New Cpp class", "Start", "Ctrl+Alt+C");
    addMenuItem("-","New Cpp class","","");
    addMenuItem("About","New Cpp class","aboutScript","");

    setLastEntry();
}

function aboutScript()
{
    echo(
        "\n" + module_name + " " + module_ver + "\n\n" +
        "_________________\n" +
        "DESCRIPTION\n\n" +
        "   PSPad Editor extension to create new cpp class files. (Cpp, h).\n"
        );
  return;
}

/*------------------------------------------------------------
    Globals
/------------------------------------------------------------*/
var fso = new ActiveXObject("Scripting.FileSystemObject");
var lastEntry = 0;

var htmlbody = 
    "<!--Line--><font color='DeepSkyBlue'><b>Enter class name</b></font>" +
    "<!--Line--><hr>" +
    "<!--Line--><table broder=0>" +
    "<!--Line--><tr>" +
    "<td> <b>Class name:</b> </td>" +
    "<td title='Enter class name.'> <input name='className' onkeyup=document.all('sourceFile').value=document.all('className').value.concat('.cpp');document.all('headerFile').value=document.all('className').value.concat('.h'); size='40' maxlength='512'> </td>" +
    "</tr>" +
    "<!--Line--><tr>" +
    "<td> <b>Base class:</b> </td>" +
    "<td title='Optional base class.'> <input name='baseClassName' size='40' maxlength='512'> </td>" +
    "</tr>" +
    "<!--Line--><tr>" +
    "<td> <b>Source file:</b> </td>" +
    "<td title='Source filename of the class.'> <input name='sourceFile' size='40' maxlength='512'> </td>" +
    "</tr>" +
    "<!--Line--><tr>" +
    "<td> <b>Header file:</b> </td>" +
    "<td title='Header filename of the class.'> <input name='headerFile' size='40' maxlength='512'> </td>" +
    "</tr>" +
    "<!--Line--><tr>" +
    "<td> Path: </td>" +
    "<td title='Enter full path destination to class files.'> <input name='path' size='40' maxlength='1024'> </td>" +
    "</tr>" +
    "</table>" +
    "<!--Line--><button style='width:70px;height:25px' name='OK' AccessKey='O' Onclick=document.all('ButtonHandler').value='Ok';><u>O</u>K</button>" +
    "&nbsp;&nbsp;&nbsp;&nbsp;" +
    "<button style='width:70px;height:25px' name='Cancel' AccessKey='C' Onclick=document.all('ButtonHandler').value='Cancel';><u>C</u>ancel</button>" +
    "<!--Line--><input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>";
/*------------------------------------------------------------
/ createMinIe
/
/ Arguemnts: none
/
/ Returns:
/   Internet Explorer object
/
/ Description:
/   Creates an hiden minimal offline IE object. 
/------------------------------------------------------------*/
function createMinIe()
{
    var ie = CreateObject("InternetExplorer.Application");
    ie.Offline = true;
    ie.Navigate("about:blank");

    while(ie.Busy)
    {
        sleep(100);
    }

    ie.MenuBar = false;
    ie.StatusBar = false;
    ie.AddressBar = false;
    ie.ToolBar = false;
    ie.showbrowserbar("{30D02401-6A81-11D0-8274-00C04FD5AE38}", false);
    ie.showbrowserbar("{EFA24E61-B078-11D0-89E4-00C04FC9E26E}", false);
    ie.showbrowserbar("{EFA24E62-B078-11D0-89E4-00C04FC9E26E}", false);
    ie.showbrowserbar("{EFA24E63-B078-11D0-89E4-00C04FC9E26E}", false);

    var tmpLeft = Math.floor((ie.Document.parentWindow.screen.width/2) - 200);
    var tmpTop = Math.floor((ie.Document.parentWindow.screen.height/2) - 200);

    ie.Left = Math.max(0, tmpLeft);
    ie.Top = Math.max(0, tmpTop);

    return ie;
}
/*------------------------------------------------------------
/ showPageIe
/
/ Arguemnts: ie             Internet explorer object
/            contentHtml    html content to show
/
/ Returns:
/   none
/
/ Description:
/   Shows IE window with specified html content.
/------------------------------------------------------------*/
function showPageIe(ie, title, contentHtml)
{
    ie.Document.body.innerHTML = contentHtml;

    ie.document.title = title;
    ie.Visible = true;

    while(ie.Busy)
    {
        sleep(100);
    }

    var shell = CreateObject("WScript.Shell");
    shell.AppActivate(ie.document.Title);
}
/*------------------------------------------------------------
/ GetPSPadTemplatePath
/
/ Arguemnts: none
/
/ Returns:
/   PsPad templates path
/
/ Description:
/   Gets pspad templates path from module path -2 + hardcoded string
/------------------------------------------------------------*/
function GetPSPadTemplatePath()
{
    var modPath = modulePath();
    var arr = new Array();
    arr = modPath.split(/\\/);

    var pspadPath = arr[0];
    for (var i=1; i<arr.length-2; i++)
    {
        pspadPath = pspadPath + "\\" + arr[i];
    }

    var templatesPath = pspadPath + "\\Template\\";

    return templatesPath;
}
/*------------------------------------------------------------
/ GetFiles
/
/ Arguemnts: path               - path to files
/            extensionFilter    - file extension regex
/
/ Returns:
/   File object.
/       longnames   : array
/       basenames  : array
/       namesNoExt : array
/       path        : string
/
/ Description:
/   Gets files with specified extension.
/------------------------------------------------------------*/
function GetFiles(path, extensionFilter)
{
    var longnames = new Array();
    var namesNoExt = new Array();
    var basenames = new Array();

    if (fso.FolderExists(path))
    {
        var folder = fso.GetFolder(path);
        var folderColl = new Enumerator(folder.Files);
        var i = 0;

        longnames[i] = "[None]";
        basenames[i] = "[None]";
        namesNoExt[i] = "[None]";

        for(; !folderColl.atEnd(); folderColl.moveNext())
        {
            var name = folderColl.item().Name;
            var dotIndex = name.lastIndexOf(".");
            var extension = name.substring(dotIndex, name.length);

            var regex = new RegExp(extensionFilter, "i");
            if ( extension.match(regex) )
            {
                //  + quotes converts items to strings
                longnames[++i] = folderColl.item() + "";
                basenames[i] = name;
                namesNoExt[i] = name.substring(0, dotIndex);
            }
        }
    }

    return { longnames:longnames, basenames:basenames, namesNoExt:namesNoExt, path:path };
}
/*------------------------------------------------------------
/ readFile
/
/ Arguemnts: filename - (IN) file name
/
/ Returns:
/   Entire file content
/
/ Description:
/   Reads entire file and returns its content.
/-------------------------------------------------------------*/
function readFile(filename)
{
    var content="";
    if (fso.FileExists(filename))
    {
        var fhandle = fso.OpenTextFile(filename, 1);
        while(fhandle.AtEndOfStream==false)
            content += fhandle.ReadLine() + "\n";

        fhandle.Close();
    }
    return content;
}
/*------------------------------------------------------------
/ addHtmlSelectAfter
/
/ Arguemnts: selects        - (IN) array of dropdown entries
/            selectedIndex  - (IN) index of the selected entry
/            htmlContent    - (IN) content to modify
/            lineNumber     - (IN) insert after this line
/
/ Returns:
/   Modified content.
/
/ Description:
/   
/-------------------------------------------------------------*/
function addHtmlSelectAfter(selects, selectedIndex, htmlContent, lineNumber)
{
    var lines = new Array();
    var newHtmlContent = "";
    lines = htmlContent.split("<!--Line-->");

    for(var i=0; i<lines.length; i++)
    {
        if (i-1 == lineNumber)
        {
            newHtmlContent += "<tr><td>Templates</td><td title='Choose cpp template to use. List of templates from pspad templates directory.'><select id='myselectid' style='width:100%'>";
            for(var j=0; j<selects.length; j++)
            {
                // last is selected
                if (j == selectedIndex)
                {
                    newHtmlContent += "<option value='" + j + "' selected='selected'>" + selects[j] + "</option>";
                }
                else
                {
                    newHtmlContent += "<option value='" + j + "'>" + selects[j] + "</option>";
                }
            }
            newHtmlContent += "</select></td></tr>";
        }
        newHtmlContent += lines[i];
    }

    return newHtmlContent;
}
/*------------------------------------------------------------
/ saveLastEntry
/
/ Arguemnts: entry  - text to save to file
/
/ Returns:
/   none
/
/ Description:
/   Saves some text to temporary file.
/-------------------------------------------------------------*/
function saveLastEntry(entry)
{
    var oTmp = fso.GetSpecialFolder(2);
    var sMytmp = oTmp + "\\" + module_name;

    if (fso.FolderExists(sMytmp) == false)
    {
        fso.CreateFolder(sMytmp);
    }
    var oMytmp = fso.GetFolder(sMytmp);
    var file = oMytmp.CreateTextFile("last.txt");
    file.write(entry);
    file.Close();
}
/*------------------------------------------------------------
/ getLastEntry
/
/ Arguemnts: none
/
/ Returns:
/   File content or empty string on error.
/
/ Description:
/   Reads temporary file content.
/-------------------------------------------------------------*/
function getLastEntry()
{
    var oTmp = fso.GetSpecialFolder(2);
    var sMytmp = oTmp + "\\" + module_name;

    var fileName = sMytmp + "\\" + "last.txt";

    var content = readFile(fileName);

    return content;
}
/*------------------------------------------------------------
/ setLastEntry
/
/ Arguemnts: none
/
/ Returns:
/   none
/
/ Description:
/   Sets lastEntry global variable. 
/-------------------------------------------------------------*/
function setLastEntry()
{
    lastEntry = 0;

    if (restoreLastEntry == true)
    {
        var content = getLastEntry();
        if (content != "")
        {
            lastEntry = content.substr(0, content.length-1);
        }
    }
}
/*------------------------------------------------------------
/ setLastEntry
/
/ Arguemnts: userDocument   - object of user input
/
/ Returns:
/   true / false
/
/ Description:
/   Validate user input.
/-------------------------------------------------------------*/
function isValidInput(userDocument)
{
    if (!userDocument.className || !userDocument.path)
    {
        echo("Error: Input invalid! className:[" + userDocument.className + "] path:[" + userDocument.path + "]");
        return false;
    }

    if (userDocument.path.charAt(userDocument.path.length) != '\\')
    {
        userDocument.path += "\\";
    }

    userDocument.sourceFilename = userDocument.path + userDocument.sourceBasename;
    userDocument.headerFilename = userDocument.path + userDocument.headerBasename;

    if (fso.FolderExists(userDocument.path) == false)
    {
        echo("Error: Folder " + path + " does not exists!");
        return false;
    }

    if ((fso.FileExists(userDocument.sourceFilename) || fso.FileExists(userDocument.headerFilename)) && userDocument.overwrite == false)
    {
        echo("Error: File exists...skip!");
        return false;
    }

    return true;
}
/*------------------------------------------------------------
/ ReplaceKeys
/
/ Arguemnts:    content         - content to change
/               userDocument    - object of user input
/
/ Returns:
/   new content
/
/ Description:
/   Creates new content replacing predefined keywords from specified content.
/-------------------------------------------------------------*/
function ReplaceKeys(content, userDocument)
{
    var regex = new Array();

    regex[0] = new RegExp("<sourceBasename>", "ig");
    regex[1] = new RegExp("<headerBasename>", "ig");
    regex[2] = new RegExp("<className>", "ig");
    regex[3] = new RegExp("<baseClassName>", "ig");
    regex[4] = new RegExp("<sourceFilename>", "ig");
    regex[5] = new RegExp("<headerFilename>", "ig");

    var newContent = content;

    newContent = newContent.replace(regex[0], userDocument.sourceBasename);
    newContent = newContent.replace(regex[1], userDocument.headerBasename);
    newContent = newContent.replace(regex[2], userDocument.className);
    newContent = newContent.replace(regex[3], userDocument.baseClassName);
    newContent = newContent.replace(regex[4], userDocument.sourceFilename);
    newContent = newContent.replace(regex[5], userDocument.headerFilename);

    return newContent;
}
/*------------------------------------------------------------
/ Main
/-------------------------------------------------------------*/
function Start()
{
    setLastEntry();

    var templatesPath = GetPSPadTemplatePath();
    var templates = GetFiles(templatesPath, "cpp");

    /* Display form and get user input */
    var ie = createMinIe();

    var newHtmlbody = addHtmlSelectAfter(templates.basenames, lastEntry, htmlbody, 5);
    showPageIe(ie, "Cpp class wizard", newHtmlbody);

    ie.document.all("className").focus();

    var buttonValue;
    var userDocument = new Object();

    while(true)
    {
        try {
            buttonValue = ie.document.all('ButtonHandler').value;
        }
        catch(e)
        {
            return;
        }

        if (buttonValue == "Cancel")
        {
            ie.Quit();
            return;
        }
        else if (buttonValue == "Ok")
        {
            userDocument.className = ie.document.all("className").value;
            userDocument.baseClassName = ie.document.all("baseClassName").value;
            userDocument.path = ie.document.all("path").value;
            userDocument.selected = ie.document.getElementById("myselectid").value;
            userDocument.sourceBasename = ie.document.all("sourceFile").value;
            userDocument.headerBasename = ie.document.all("headerFile").value;
            userDocument.overwrite = false;

            saveLastEntry(userDocument.selected);

            break;
        }
        sleep(100);
    }

    ie.Quit();

    if (!isValidInput(userDocument))
    {
        return;
    }

    /* Create files. */
    var edSrc   = newEditor();
    var edH     = newEditor();

    edSrc.newFile(userDocument.sourceFilename);
    edH.newFile(userDocument.headerFilename);

    try {
    if (edSrc.saveFile() == false || edH.saveFile() == false)
    {
        echo("Error: Saving files! [" + userDocument.sourceFilename + "] [" + userDocument.headerFilename + "] ");
        return;
    }
    }
    catch(e)
    {
        return;
    }

    if (templates.basenames[userDocument.selected] != "[None]")
    {
        var srcTemplateContent = readFile(templates.longnames[userDocument.selected]);
        var headerTemplateContent = readFile(templates.path + templates.namesNoExt[userDocument.selected] + ".h");

        srcTemplateContent = ReplaceKeys(srcTemplateContent, userDocument);
        headerTemplateContent = ReplaceKeys(headerTemplateContent, userDocument);

        edSrc.Text(srcTemplateContent);
        edH.Text(headerTemplateContent);
    }

    edSrc.saveFile();
    edH.saveFile();

    return;
}