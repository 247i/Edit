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
 *
 *  Version 1.1 by V.Pshentsov <pshentsoff@gmail.com> 22.03.2012 14:55:07
 *  Changes:
 *  - Color syntax highlighting added (Syntax\ToDoListJS.INI).
 *  - Any statements or regular expressions is now supported - you asked 
 *    by input box what statements you need to get. You can press 'Cancel' button
 *    on input box or <Esc> to cancel script execution. 
 *    Examples:
 *      "TODO:" - gets all strings started with "TODO:"
 *      "TODO:|LOOK:" - gets all with "TODO:" or "LOOK:"
 *      & so on...     
 *  - Line numbers added at log output.
 *  - Too long file paths shortened from begining to maximum characters length.
 *  - Some bugs fixed.          
 *  
 */

var module_name = "TodoListJS";
var module_ver = "1.1";
var module_title = "TodoListJS";

function openScript() {
    var obj1 = newEditor();
    obj1.openFile(moduleFileName("TodoListJS"));
}

function findTodos(fs, filename, restr) {
    var txt = "";
    var r, re, line;
    var i = 0;
    var fr;
    var filename_maxlen = 60;
    var filename_str;
    
    if(filename.length > filename_maxlen) { 
      filename_str = '...'+filename.substr(filename.length-filename_maxlen, filename_maxlen);
      } else {
      filename_str = filename;
      } 
    
    re = new RegExp(restr);
    fr = fs.GetFile(filename).OpenAsTextStream(1, 0);

    while (!fr.AtEndOfStream) {
        line = fr.ReadLine();
        i++;
        r = line.search(re);
        if (r != -1) {
            txt = txt.concat(filename_str + "["+ i +"]: " + line.substr(r) + "\n");
        }
    }
    fr.Close( );

    return txt;
}

function todoListFile() {

    var restr = "TODO:";
    restr = inputText("Current file todos. Enter word sentence or RegExp:", restr);
    if(!restr) return;
    
    var fs = new ActiveXObject("Scripting.FileSystemObject")
    var data = "";
    var fr;
    var todos = "";

    var editor = newEditor();
    editor.assignActiveEditor();
    fr = editor.fileName();
    if (fs.FileExists(fr)) {
        todos = findTodos(fs, fr, restr);
        if (todos != "") {
            data = data.concat(todos);
        }
    }
    print(data);
}

function todoListAllFiles() {
    var restr = "TODO:";
    restr = inputText("Opened files todos. Enter word sentence or RegExp:", restr);
    if(!restr) return;
    
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
            todos = findTodos(fs, fr, restr);
            if (todos != "") {
                data = data.concat(todos);
            }
        }
        i++;
    }
    print(data);
}

function todoListProject() {

    var restr = "TODO:";
    restr = inputText("Project files todos. Enter word sentence or RegExp:", restr);
    if(!restr) return;
    
    var fs = new ActiveXObject("Scripting.FileSystemObject")
    var data = "";
    var fr;
    var i = 0;              // Project file count.
    var j = 0;              // Project file and folder count.
    var todos = "";
    var r;
    
    //I'm already know what is this project is =)
    //var text = new String(projectFileName());
    //data = "Project: " + text.substring(text.lastIndexOf('\\') + 1, text.lastIndexOf('.')) + "\n";
    //data = data.concat("------------------------------------------------------------------------\n");
    var svn = new RegExp(".svn");
    var cvs = new RegExp("CVS");
    var git = new RegExp(".git");
    // projectFilesCount() returns a count of the files in a project, however
    // projectFiles(j) returns folders and files! Seems like a bug.
    var projfc = projectFilesCount();
    while(i < projfc) {
        fr = projectFiles(j);
        j++;
        
        //echo(i+" of "+projfc+" files processed.\nCurrent is ["+fr+"]\nNext is ["+projectFiles(j)+"]");
        //Break if end of projectFiles array arrived, but counter 'i' still smaller than bug function result
        if(!fr.length) {
          //echo('Break at #'+j+': ['+fr+'].');
          break;
          }
        //We can simply skip all project folders
        if(!fs.FileExists(fr)) {
          //echo('Skip #'+j+': ['+fr+'] - not a file.');
          continue;
          }

        // Skip SVN directories.
        r = fr.search(svn);
        if (r != -1) {
          if (fr && fs.FileExists(fr)) {
            // It's a file, increment the file count.
            i++;
          }
          continue;
        }
        // Skip CVS directories.
        r = fr.search(cvs);
        if (r != -1) {
          if (fr && fs.FileExists(fr)) {
            // It's a file, increment the file count.
            i++;
          }
          continue;
        }
        // Skip GIT directories.
        r = fr.search(git);
        if (r != -1) {
          if (fr && fs.FileExists(fr)) {
            // It's a file, increment the file count.
            i++;
          }
          continue;
        }

        // It's a file.
        if (fr && fs.FileExists(fr)) {
            todos = findTodos(fs, fr, restr);
            if (todos != "") {
                data = data.concat(todos);
            }
            i++;
        }
        
    }
    
    print(data);
}

function print(data) {
    var lines = new Array();
    var line = "";
    var c = 0;
    
    logClear();                 // clearing log window
    lines = data.split('\n');
    
    for (var x in lines) {
    
        if(lines[x]) {
          c++;
          line = c + ": " + lines[x];
          } else {
          line = "\n";
          }
        
        logAddLine(line);   // printing the result to the log
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
        "enhanced by BadBoy,\n" +
        "enhanced by V.Pshentsov.\n\n" +
        "This script looks for some \"TODO\" statements in the\n" +
        "current project, opened files or current file and displays'em in the log.\n" +
        "Also added syntax highlighting in .\\Syntax\\ToDoListJS.INI\n"+
        "Your statements should start with\n\n" +
        "TODO, DONE, UNDONE, ERROR, BUG, FEATURE, IMPORTANT, LOOK with \':\' or \'!\'\n" +
        "in order to be found.\n"+
        "Example: \"TODO:\", \"ERROR!\" and so on.\n"+
        "You are free and easy to add your own statements at all."
    );
    return;
}

function progressDone(fs, fr) {
    var result = new Array();
    var r, re, line;
    var i = 0;
    var fr;
    //TODO: ~ 23.03.2012 10:56:07 array of "done","todo" word to search for
    re = new RegExp(restr);
    fr = fs.GetFile(filename).OpenAsTextStream(1, 0);

    while (!fr.AtEndOfStream) {
        line = fr.ReadLine();
        r = line.search(re);
        if (r != -1) {
          //TODO: ~ 23.03.2012 10:56:26 count each of 3 word groups to result array  
          }
      }
    fr.Close( );

    return result;
}

function progressDoneFile() {

    var fs = new ActiveXObject("Scripting.FileSystemObject")
    var data = "";
    var fr;
    var todos = "";

    var editor = newEditor();
    editor.assignActiveEditor();
    fr = editor.fileName();
    if (fs.FileExists(fr)) {
        todos = findTodos(fs, fr, restr);
        if (todos != "") {
            data = data.concat(todos);
        }
    }
    print(data);
}

function Init() {
    addMenuItem("View &current file ToDos", "ToDo list", "todoListFile", "CTRL+ALT+C");
    addMenuItem("View all open file ToDo&s", "ToDo list", "todoListAllFiles", "CTRL+ALT+S");
    addMenuItem("View project To&Dos", "ToDo list", "todoListProject", "CTRL+ALT+D");
    addMenuItem("-","ToDo list","","");														// Menu divider
    addMenuItem("Edit this script", "ToDo list", "openScript");
    addMenuItem("About", "ToDo list", "about");
}