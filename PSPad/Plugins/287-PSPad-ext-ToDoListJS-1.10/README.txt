PSPad Todo script
(c) by Carney

Version 1.1 from 22.03.2012 
by V.Pshentsov <pshentsoff@gmail.com>
Whats New:
  - Added Color syntax highlighting INI-file for todos files (Syntax\ToDoListJS.INI).
  - Any statements or regular expressions is now supported - you asked 
   by input box what statements you need to get. You can press 'Cancel' button
   on input box or <Esc> to cancel script execution. 
   Examples:
     "TODO:" - gets all strings started with "TODO:"
     "TODO:|LOOK:" - gets all with "TODO:" or "LOOK:"
     & so on...     
  - Line numbers added at log output.
  - Too long file paths shortened from begining to maximum characters length.
  - Some bugs fixed.          

First enhancement by BadBoy 2009.
Version 1.00 by BadBoy 21/3/2010.
Whats New:
-Project file count processing was wrong, resulting in missed files.
-Added CVS filtering.
-Improved menus. 
-Added hotkey.
-Added extraction from the current file.
-Removed some unnecessary code.

NOTE:
Project ToDos are retrieved even if you don't currently have a project open. 
If any project has been opened in PsPad before running this script,
extracting project ToDos will retrieve the ToDo entries for the last
project that was opened. PsPad always knows about the last project,
even if you haven't opened one in the current session and
projectFilesCount()and projectFiles() always return data relating
to that project i.e. PsPad considers it to be open even if you're not
currently showing the Project Manager.  
