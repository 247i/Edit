'*******************************************************************************
'Filename:     ProjRefresh.vbs                                                 *
'Author:       Holger Biehl                                                    *
'Lastmodified: 21.06.2010 15:17:54                                             *
'VersionInfo:  5.002                                                           *
'Tested with:  PSPad Build 2390                                                *
'*******************************************************************************

'*******************************************************************************
'Notes:                                                                        *
'  -In the following Lines you see the description of possible options.        *
'    To manipulate the behavior or to look how it works,                       *
'    you have to look in the executed script PRefresh.vbs                      *
'  -For giving other Parameters to run the Refresh from PSPad use the          *
'    Public Const myParams in this script (delimit with spaces)                *
'    i.e. Public Const myParams = "Sort=PathExt ActiveOnly=true"               *
'*******************************************************************************

'*******************************************************************************
'Parameter For Refresh (Internal):                                             *
'  PSPad-VBScripting isn't able to give parameters to the script.              *
'  On the other Hand PsPad implements the hole Environment to the Scripting    *
'  So we can get all needed Projectinformation and run the External Script with*
'  the right parameters.                                                       *
'  To run a projectrefresh the File PRefresh.vbs is needed to be in            *
'  PsPad-Installfolder\script (e.g. %programfiles%\PSPad Editor\script)        *
'*******************************************************************************

'*******************************************************************************
'Options For Refresh (Analysed in PRefresh.vbs)                                *
'Everything behaves by priority of given facts:                                *
'1.If no <PathsFilename> and no <IniFilename> is given:                        *
'   Normally Projects DefaultPath is search                                    *
'   The behavior can be changed by Setting Docroot to True                     *
'   If Docroot is true DocumentPath is searched                                *
'*******************************************************************************

'*******************************************************************************
'1.If no <PathsFilename> and no <IniFilename> is given:                        *
'   Normally Projects DefaultPath is search                                    *
'   The behavior can be changed by Setting Docroot to True                     *
'   If Docroot is true DocumentPath is searched                                *
'      "DocRoot="     =>                                                       *
'      omitted: Search the Projectfolder given in Project                      *
'      both: Search the Documentfolder and the Projectfolder given in Project  *
'      true: Search the Documentfolder given in Project                        *
'                                                                              *
'*******************************************************************************

'*******************************************************************************
'2.If a propath.txt is given in projectfiles Path, but no <IniFilename>:       *
'   propath.txt is searched line by line                                       *
'   A new Projectfolder will be created with every line where Path exists      *
'   Extensionfilter and Pathalias is Not available                             *
'*******************************************************************************

'*******************************************************************************
' Simple Startup Example (to Copy in <PathsFilename> without comments):        *
'C:\temp\src
'C:\temp\bin
'C:\temp\lib
'*******************************************************************************

'*******************************************************************************
'3.<IniFilename> is available or Parameters given:                             *
'   with a <IniFilename> in projectfiles path will be searched for:            *
'   ( "Definition=<Text>" => Variable/Array used in Script  (* Remarks)        *
'       description )                                                          *
'                                                                              *
'     "SubToFolder=true"  => SubToFolder                                       *
'       creates ProjectFolder for each SubFolder                               *
'                                                                              *
'     "Propath="          => ProjectPaths                                      *
'       a commaseparated list of Paths to search in                            *
'       Propath=d:\project,c:\windows,c:\program files..                       *
'                                                                              *
'     "Pathalias="        => ProjectPaths         *                            *
'       PathAlias is a commaseparated list of Names                            *
'       Pathalias=project,windows,programs                                     *
'       There should be as much entries in this List as in Propath             *
'     or                                                                       *
'     "CutPathAt="        => CutPathAt            *                            *
'       is a simple string to search in folders fullpath                       *
'       cut what is before                                                     *
'       e.g. CutPathAt=\pa\  and you'll get for                                *
'       D:\Projects\Projectpath\pa\yVersion\src                                *
'       a yVersionsrc as ProjectFolder                                         *
'       CutPathAt will be ignored when defined a PathAlias for equivalent Path *
'       (* only one of both)                                                   *
'                                                                              *
'     "FileFilter="       => FileFilter                                        *
'       a commaseparated extensionlist                                         *
'       FileFilter=.exe,.bat,.html                                             *
'                                                                              *
'     "FileAlias="        => FileFilter                                        *
'       a commaseparated list of Aliases for extensionfilter                   *
'       FileAlias=Executable,Batchfile,Webside                                 *
'       There should be as much entries in this List as in FileFilter          *
'                                                                              *
'     "NoUnknown=true"    =>                                                   *
'       ignores undefined extensions, otherwise all unknown files              *
'       are registered in NoExt-Folder                                         *
'                                                                              *
'     "SubFolders="       => CSubFolders                                       *
'       a commaseparated string with foldernames                               *
'       have to be created as subfolders (logical and Multidimensional)        *
'       Subfolders=bin,src                                                     *
'       Projecttree                                                            *
'         -bin                                                                 *
'           -src                                                               *
'             -File4                                                           *
'           -File1                                                             *
'         -src                                                                 *
'           -File2                                                             *
'         -File3                                                               *
'       File1 is in bin, File2 is in src, File3 is in test, File4 is in bin\src*
'       (only if SubToFolder is not defined)                                   *
'                                                                              *
'     "ExclFolders="      => SExclFolder                                       *
'       a commaseparated string with Folder-Names/Paths                        *
'       not to be searched in                                                  *
'       ExclFolders=bin                                                        *
'       Projecttree                                                            *
'         -src                                                                 *
'           -File2                                                             *
'         -File3                                                               *
'       (same Example as for SubFolders)                                       *
'       Rootfolders can not be excluded                                        *
'                                                                              *
'     "ExclFiles="        => SExclFiles                                        *
'       a commaseparated string with filenames to be ignored                   *
'       entries with Backslash will be searched in absolute path of file       *
'       entries with Asterisk at the end will be searched in file-name/path    *
'         with the InStr-Function                                              *
'       ExclFiles=File4                                                        *
'       Projecttree                                                            *
'         -bin                                                                 *
'           -File1                                                             *
'         -src                                                                 *
'           -File2                                                             *
'         -File3                                                               *
'       ExclFiles=File*                                                        *
'       Projecttree                                                            *
'       (same Example as for SubFolders)                                       *
'                                                                              *
'     "ExclExtens="       => SExclExtens                                       *
'       a commaseparated string with fileextensions to be ignored              *
'       ExclFiles=.exe,.zip                                                    *
'       (not for Extensions of FileFilter)                                     *
'                                                                              *
'     "Sort="  =>                                                              *
'       If omitted no Sort is activated. This means all files are flat listed  *
'       sorted by there Subfolder. If SubtoFolder is true all SubFolders       *
'       are created in one Root. In this case: files are listed                *
'        after they are found ( not sorted by name )                           *
'       others: (Root Folders will be created for Paths and Extensions Aliases)*
'       -PathExt: 1.Sort by Path then Sort by Extension                        *
'       -ExtPath: 1.Sort by Extension then Sort by Path                        *
'       -Path:    1.Sort by Path then by SubFolder                             *
'       -Ext:     1.Sort by Extension then Sort by SubFolder                   *
'                                                                              *
'     "ActiveOnly=true"   =>                                                   *
'       If more than one Folders are searched only the first hit of a relative *
'       Path is saved.                                                         *
'       ProPath=c:\temp;d:\temp                                                *
'       File test.txt is in both Paths: only c:\temp\test.txt will be          *
'       in the Project.                                                        *
'                                                                              *
'*******************************************************************************

'*******************************************************************************
' Simple Startup Example (to Copy in <IniFilename> without comments):          *
''(Pathsorting will be disabled, Extensionsorting depends on Projectsettings)  *
'Propath=c:\temp
'Sort=PathExt
'SubToFolder=true
'CutPathAt=src
'*******************************************************************************

'*******************************************************************************
' Versionhistory:                                                              *
' 1.0 :   first published Version only for flat searching                      *
'         by Paths.txt or found projectpaths                                   *
' 2.0 :   <IniFilename> and Extensionfilter                                    *
' 2.001 : fixed bugs with spaces in foldernames                                *
' 3.0 :   Subfolder-Creation with <IniFilename> CutPaths-Option and PathsAlias *
' 3.001 : restructured script for better sorting and performance               *
'         new ExtRoots and NoUnknown                                           *
'         fixed Bugs with PathsAlias                                           *
' 3.002 : Make Script execution independent from PSPad                         *
'         to enable work while refreshing                                      *
' 3.003 : Subfolder-Option with Names of Folders for which Projectfolders      *
'         should be created                                                    *
' 3.004 : Exclusionfilters for Folders Files and Extensions                    *
' 3.005 : fixed bug when PSPad is minimized to activate Window and close it    *
'         requires correctly registered .ppr extension in Windows              *
'         requires also PSPad is set to open "only one program instance"       *
'         to be set in properties of System Integration                        *
' 3.006 : fixed not to lookup Case Sensitive for Names                         *
'         Original Project to be loaded readonly                               *
' 3.007 : Project reloading not automatically, leaving tmp-File as Backupcopy  *
' 3.008 : fixed Bug in Foldercreation                                          *
' 3.010 : Batch-Mode and new Parameterstructure                                *
' 4.000 : Switches for sortlevelmodifications                                  *
' 4.002 : Improvements for PR-Filenames and Simple Projectfunction             *
'         (like in PSPad 4.5.5 for Creation)                                   *
' 5.000 : Complete new Sortbehaviors, better Performance, more Options,        *
'         direct Parameterinput , no Sortfiles needed using Arrays,            *
'         uses PsPad Extensionfilter if no other defined,                      *
'         Ini/Path-file with projectname & usage of Dictionaries and Variables *
' 5.001   little corrections: sorting subfolders before sorting files          *
' 5.002   dynamic Mergesort: the fastest Way to Sort (more Sortoptions)        *
'*******************************************************************************

Public Const myParams = ""

Const MODULE_NAME  = "PRefresh"
Const MODULE_VER   = "5.002"
Const DIALOG_TITLE = "Refresh Project"

' ------------------------------------------------------------------------------
'                           Default modulesubroutines.
' ------------------------------------------------------------------------------

' "Init" is required
' It is called automatically during initialization to create menu items.
Sub Init
  addMenuItem DIALOG_TITLE, "", MODULE_NAME, "ALT+F5"
End Sub

'*************************************************
' Main Procedure called by PSPad                 *
'*************************************************
Sub PRefresh

  'Initialize
  'SYS: Microsoft (r) Script Runtime
  Set fs = CreateObject("Scripting.FileSystemObject")
  'SYS: Windows Script Host Runtime Library
  Set wshshell = CreateObject("WScript.Shell")

  Dim ProjectFile
  Dim ScriptFolder
  Dim PScript

  'Checking Main Parameter
  If modulePath() <> "" Then
    ScriptFolder = Left(modulePath(),Len(modulePath()) - 1)
    ScriptFolder = Left(ScriptFolder,InstrRev(ScriptFolder,"\"))
    PScript = ScriptFolder & "PRefresh.vbs"
  End If
  If projectFileName() <> "" Then
    ProjectFile = projectFileName()
    projectSave()
  End If
  wshshell.run """" & PScript & """ """ & ProjectFile & """ " & myParams,1,false

End Sub
