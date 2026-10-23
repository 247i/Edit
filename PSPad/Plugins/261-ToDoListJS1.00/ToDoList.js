/**
 *  PSPad Todo script
 *  (c) by Carney
 *
 *  First enhancement by BadBoy 2009.
 *  Version 1.00 by BadBoy 21/3/2010.
 *  Changes:
 *  -Project file count processing was wrong, resulting in missed files.
 *  -Added CVS filtering.
 *  -Improved menus. 
 *  -Added hotkey.
 *  -Added extraction from the current file.
 *  -Removed some unnecessary code.
 *  
 *  NOTE:
 *  Project ToDos are retrieved even if you don't currently have a project open. 
 *  If any project has been opened in PsPad before running this script,
 *  extracting project ToDos will retrieve the ToDo entries for the last
 *  project that was opened. PsPad always knows about the last project,
 *  even if you haven't opened one in the current session and
 *  projectFilesCount()and projectFiles() always return data relating
 *  to that project i.e. PsPad considers it to be open even if you're not
 *  currently showing the Project Manager.  
 */

var module_name = "TodoListJS";
var module_ver = "1.00";
var module_title = "TodoListJS";

function openScript() {
    var obj1 = newEditor();
    obj1.openFile(moduleFileName("TodoListJS"));
}

function findTodos(fs, filename) {
    var txt = "";
    var r, re, line;
    var i = 0;
    var fr;

    re = new RegExp("TODO:");
    fr = fs.GetFile(filename).OpenAsTextStream(1, 0);

    while (!fr.AtEndOfStream) {
        line = fr.ReadLine();
        i++;
        r = line.search(re);
        if (r != -1) {
            txt = txt.concat(filename + ": "+ i +": " + line.substr(r) + "\n");
        }
    }
    fr.Close( );

    return txt;
}

function todoListFile() {
    var fs = new ActiveXObject("Scripting.FileSystemObject")
    var data = "";
    var fr;
    var todos = "";
    
    var editor = newEditor();
    editor.assignActiveEditor();
    fr = editor.fileName();
    if (fs.FileExists(fr)) {
        todos = findTodos(fs, fr);
        if (todos != "") {
            data = data.concat(todos);
        }
    }
    print(data);
}

function todoListAllFiles() {
    var fs = new ActiveXObject("Scripting.FileSystemObject")
    var data = "";
    var fr;
    var todos = "";
    var i = 0;
    
    var editor = newEditor();
    while(i < editorsCount()){
        editor.assignEditorByIndex(i)
        fr = editor.fileName();
        if (fs.FileExists(fr)) {
            todos = findTodos(fs, fr);
            if (todos != "") {
                data = data.concat(todos);
            }
        }
        i++;
    }
    print(data);
}

function todoListProject() {
    var fs = new ActiveXObject("Scripting.FileSystemObject")
    var data = "";
    var fr;
    var i = 0;              // Project file count.
    var j = 0;              // Project file and folder count.
    var todos = "";
    var r, svn, cvs;
    var text = new String(projectFileName());
    
    data = "Project: " + text.substring(text.lastIndexOf('\\') + 1, text.lastIndexOf('.')) + "\n";
    data = data.concat("------------------------------------------------------------------------\n");
    svn = new RegExp(".svn");
    cvs = new RegExp("CVS");
    // projectFilesCount() returns a count of the files in a project, however
    // projectFiles(j) returns folders and files! Seems like a bug.
    while(i < projectFilesCount()) {
        fr = projectFiles(j);
        j++;

        // Skip SVN directories.
        r = fr.search(svn);
        if (r != -1) {
          if (fs.FileExists(fr)) {
            // It's a file, increment the file count.
            i++;
          }
          continue;
        }
        // Skip CVS directories.
        r = fr.search(cvs);
        if (r != -1) {
          if (fs.FileExists(fr)) {
            // It's a file, increment the file count.
            i++;
          }
          continue;
        }

        // It's a file.
        if (fs.FileExists(fr)) {
            todos = findTodos(fs, fr);
            if (todos != "") {
                data = data.concat(todos);
            }
            i++;
        }
    }
    print(data);
}

function print(data) {
    logClear();                 // clearing log window
    var lines = new Array();
    lines = data.split('\n');
    for (var x in lines) {
        logAddLine(lines[x]);   // printing the result to the log
    }
}

function about() {
    /*	Extension description, how to use, license,
    	open source project web address, and developer
    	credits.
    */

    echo(
        "\n" + module_name + " " + module_ver + "\n\n" +
        "PSPad Todo script\n\n" +
        "by Carney,\n" +
        "enhanced by BadBoy.\n\n" +
        "This script looks for TODO statements in the\n" +
        "current project or current file and displays\n" +
        "them in the log.\n\n" +
        "Your statements should start with\n\n" +
        "\"TODO:\"\n\n" +
        "in order to be found."
    );
    return;
}


function Init() {
    addMenuItem("View &current file ToDos", "ToDo list", "todoListFile", "CTRL+ALT+C");
    addMenuItem("View all open file ToDo&s", "ToDo list", "todoListAllFiles", "CTRL+ALT+S");
    addMenuItem("View project To&Dos", "ToDo list", "todoListProject", "CTRL+ALT+D");
    addMenuItem("-","ToDo list","","");														// Menu divider
    addMenuItem("Edit this script", "ToDo list", "openScript");
    addMenuItem("About", "ToDo list", "about");
}