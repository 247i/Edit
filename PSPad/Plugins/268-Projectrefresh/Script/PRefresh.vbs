'*******************************************************************************
'Filename:     PRefresh.vbs                                                    *
'Author:       Holger Biehl                                                    *
'Lastmodified: 21.06.2010 15:17:54                                             *
'VersionInfo:  5.002                                                           *
'Tested with:  PSPad Build 2390                                                *
'*******************************************************************************

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!*
'!Very Important!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!*
'!Within PSPad there's no chance to give Parameters over VBScript!!!!!!!!!!!!!!*
'!Do not save this file in the VBScript-Directory of PSPad!!!!!!!!!!!!!!!!!!!!!*
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!*

'*******************************************************************************
' Notes:                                                                       *
'  -For merging more than one Extensions in one Projectfolder                  *
'    give them the same Aliases                                                *
'  -If no ExtensAlias is given,                                                *
'    the Extensions in the list are used themself to create Projectfolders     *
'  -If less than 2 Extensions are given, the Files will be stored flat         *
'    in one folder                                                             *
'  -If less than 2 Paths are given, the Files will be also stored flat         *
'    in one folder                                                             *
'  -If the <IniFilename> File have to be used by other Programs, too           *
'    there's the option to specify a Section called [PSInit]                   *
'    which will be the only section to be searched in <IniFilename>            *
'    for the described Keywords                                                *
'  -<IniFilename> entries can be commented with any Character,                 *
'    because the Keys are only guilty at beginning of lines                    *
'  -In the following Lines you see the description of possible options.        *
'  -To change the name of Initializationfiles modify the lines higher 228      *
'    It is possible to pass the file as parameter too!                         *
'    i.e. Inifile=C:\temp\project.ini                                          *
'  -Parameternames are Casesensitive!!! This is default in VBScript            *
'*******************************************************************************

'*******************************************************************************
'Parameters For Refresh (External):                                            *
'1.PsPad Projectfile                                                           *
'2.All Parameters in Form: Parameter=Values                                    *
'   separated by Space (Values with Spaces should be quoted with Parametername)*
'                                                                              *
'For all cases: Silent=true : No request for replacing the old Project (Batch) *
'               IniGen=true : Writedown settings to a inifile in projects path *
'*******************************************************************************

'*******************************************************************************
'Options For Refresh (Analysed in this File)                                   *
'   With Build 2388 comes the Storage of Extensionfilter within Project.       *
'   This will be recognized in all cases, while no other Filter is defined.    *
'                                                                              *
'Everything behaves by priority of given facts: (3 Cases)                      *
'*******************************************************************************

'*******************************************************************************
'1.If no <PathsFilename> and no <IniFilename> is given:                        *
'   SubToFolder is set automatically to true                                   *
'   Sort is set to Path                                                        *
'   The script behaves almost like the "syncronize with folder" function       *
'   The only differences are the Projectname not based on Foldersname          *
'    and Folder gets a more detailed alias than name.                          *
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
'2.If a <PathsFilename> is given in projectfiles Path, but no <IniFilename>:   *
'   Paths.txt is searched line by line                                         *
'   SubToFolder is set automatically to true                                   *
'   Sort is set to Path                                                        *
'   The script behaves almost like the "syncronize with folder" function       *
'*******************************************************************************

'*******************************************************************************
' Simple Startup Example (to Copy in <PathsFilename> without comments):        *
'C:\temp\src
'C:\temp\bin
'C:\temp\lib
'*******************************************************************************

'*******************************************************************************
'3.<IniFilename> is available, Parameters in Projectfile or Parameters at Call:*
'   with a <IniFilename> in projectfiles path will be searched for:            *
'   ( "Definition=<Text>" => Variable/Array used in Script  (* Remarks)        *
'       description )                                                          *
'                                                                              *
'     "SubToFolder=true"  => SubToFolder                                       *
'       creates ProjectFolder for each SubFolder                               *
'                                                                              *
'     "Paths="          => ProjectPaths                                        *
'       a commaseparated list of Paths to search in                            *
'       Paths=d:\project,c:\windows,c:\program files..                         *
'                                                                              *
'     "PathsAlias="        => ProjectPaths        *                            *
'       PathsAlias is a commaseparated list of Names                           *
'       PathsAlias=project,windows,programs                                    *
'       There should be as much entries in this List as in Paths               *
'     or                                                                       *
'     "CutPathAt="        => CutPathAt            *                            *
'       is a simple string to search in folders fullpath                       *
'       cut what is before                                                     *
'       e.g. CutPathAt=\pa\  and you'll get for                                *
'       D:\Projects\Projectpath\pa\yVersion\src                                *
'       a yVersionsrc as ProjectFolder                                         *
'       CutPathAt will be ignored when defined a PathsAlias for equivalent Path*
'       (* only one of both)                                                   *
'                                                                              *
'     "Extensions="       => Extensions                                        *
'       a commaseparated extensionlist                                         *
'       Extensions=.exe,.bat,.html                                             *
'                                                                              *
'     "ExtensAlias="        => Extensions                                      *
'       a commaseparated list of Aliases for extensionfilter                   *
'       ExtensAlias=Executable,Batchfile,Webside                               *
'       There should be as much entries in this List as in Extensions          *
'                                                                              *
'     "NoUnknown=true"    =>                                                   *
'       ignores undefined extensions, otherwise all unknown files              *
'       are registered in _NoExt-Folder                                        *
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
'       (not for Extensions of Extensions)                                     *
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
'    careful with the following: each Time a Root changes, the Subfoldertree   *
'        will be written again:                                                *
'       -PathExt1: 1.Sort by Path then Sort by Extension; only Folder for Paths*
'       -ExtPath1: 1.Sort by Extension then Sort by Path; only Folder for Ext  *
'       -PathExt0: 1.Sort by Path then Sort by Extension; no Folders           *
'       -ExtPath0: 1.Sort by Extension then Sort by Path; no Folders           *
'                                                                              *
'     "ActiveOnly=true"   =>                                                   *
'       If more than one Folders are searched only the first hit of a relative *
'       Path is saved.                                                         *
'       Paths=c:\temp;d:\temp                                                  *
'       File test.txt is in both Paths: only c:\temp\test.txt will be          *
'       in the Project.                                                        *
'                                                                              *
'*******************************************************************************

'*******************************************************************************
' Simple Startup Example (to Copy in <IniFilename> without comments):          *
''(Pathsorting will be disabled, Extensionsorting depends on Projectsettings)  *
'Paths=c:\temp
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

'Constants
'Set Filenames For Projectinitialization here (should be in every Project equal)
'You can set the Constants to "*" for expecting a
' >Projectfilename(without Extension)<.ini or
' >Projectfilename(without Extension)<.path
Public Const IniFilename   = "*"
Public Const PathsFilename = "*"

Public Const DefProjName   = "Dev"
Public Const Delimiter     = ";"
Public Const WriteBegin    = "[Project tree]"
Public Const WriteEnd      = "["
Public Const DefaultDirDef = "DefaultDir="
Public Const DocDirDef     = "DocumentRoot="
Public Const ExtDef        = "Extensions="
Public Const ProjExt       = ".ppr"
Public Const WorkExt       = ".ppr"
Public Const InitExt       = ".ini"
Public Const PathExt       = ".path"
Public Const TempExt       = ".tmp"
Public Const OldExt        = ".old"
Public Const Section       = "[PSInit]"
Public Const InitEnd       = "["
'Objects
Public WshShell , ProjectFolderObj , TempFileObj , WshFs , WshParam
Public WshExtAlias , WshPathsAlias , ArgsObj

'Variables
Public Projectname , ProjectFileFolder , ProjectFile , ActualPath , SortCases
Public DocRootCases

Dim    Answer

'Arrays
Public Paths() , ProjectPaths() , CSubFolders , Extensions() , FExclFolders
Public ExtAlias() , PathsAlias() , SubFolders() , Folders() , Relatives()
Public FExclFiles , FExclExtens , Parameter , IncludePaths

'Fetch needed System Classes
'SYS: Windows Script Host Runtime Library
Set WshShell = CreateObject( "WScript.Shell" )
'SYS: Microsoft (r) Script Runtime
Set WshFs    = CreateObject( "Scripting.FileSystemObject" )
'SYS: Scripting Dictionaries
Set WshParam     = CreateObject( "Scripting.Dictionary" )
Set WshExtAlias  = CreateObject( "Scripting.Dictionary" )
Set WshPathsAlias = CreateObject( "Scripting.Dictionary" )

'fetch Arguments given after call of script
Set ArgsObj = WScript.Arguments

SortCases    = Split( "pathext,extpath,path,ext,pathext1,extpath1,pathext0,extpath0" , "," )
DocRootCases = Split( "both,true" , "," )

'Preset the Values to search in <IniFilename> file here
Parameter = Split( "Paths,PathsAlias,Extensions,ExtensAlias,SubToFolder," & _
                   "Sort,NoUnknown,CutPathAt,SubFolders,ActiveOnly," & _
                   "ExclFolders,ExclFiles,ExclExtens,Delimiter,DocRoot," & _
                   "Silent,Inifile,Pathfile,IniGen,IncludePaths", _
                   "," )

ReDim Preserve Paths(6,0)
ReDim Preserve SubFolders(2,0)
ReDim Preserve Folders(0)
ReDim Preserve Relatives(0)

'Load Projectsettings
GetProjectData

'Projectfile is not optional
If ProjectFile = "" Then
  WScript.Quit
End If

'Call Sub Refresh
Refresh

If WshParam.Item("IniGen") = "true" Then
  WriteIniFile
End If

'Check if Workspace have been written
If WshParam.Item("Silent") = "true" Then
  If Not WshFs.FileExists(ProjectFile & TempExt) Then
    WScript.Quit
  End If
  WshFs.CopyFile ProjectFile & TempExt , ProjectFile , true
  WshFs.GetFile(ProjectFile & TempExt).Attributes = 2
  WshFs.GetFile(ProjectFile & OldExt).Attributes  = 2
  'wshshell.run ProjectFile
Else
  If Not WshFs.FileExists(ProjectFile & TempExt) Then
    MsgBox "The Project was not refreshed, check your Projectsettings",16,"Refresh failed"
    WScript.Quit
  End If
  'Question for open the new project
  Answer = MsgBox("The Projectrefresh is finished for" & vbCr & ProjectFile & vbCr & "Should the Project been opened?",36,"Project load")
  If Answer = 6 Then
    WshFs.CopyFile ProjectFile & TempExt , ProjectFile , true
    WshFs.GetFile( ProjectFile & TempExt ).Attributes = 2
    WshFs.GetFile( ProjectFile & OldExt ).Attributes  = 2
    Wscript.Sleep 1000
    cWorkspace = Replace( ProjectFile , ProjExt , WorkExt , 1 , -1 , 1 )
    If WshFs.FileExists(cWorkspace) Then
      WshShell.Run cWorkspace
    Else
      WshShell.Run ProjectFile
    End If
  End If
End If

'release Objects
Set ArgsObj  = Nothing
Set WshShell = Nothing
Set WshFs    = Nothing
Set WshParam = Nothing

WScript.Quit

'-------------------------------------------------------------------------------
'Subroutines                                                                   -
'-------------------------------------------------------------------------------

'*******************************************************************************
' Load Project                                                                 *
'*******************************************************************************

Sub GetProjectData

  Dim SearchKey , SearchValue , Editor , GetPaths
  Dim RightSection , DefaultFile , DocPath , ActualLine , ProjectIni , I , J

  'First of all check if Projectfile extension is correctly registered
  SearchKey = "HKEY_CLASSES_ROOT\" & ProjExt & "\"
  If RegExists( SearchKey ) Then
    SearchValue = wshshell.RegRead( SearchKey )
    If RegExists( "HKEY_CLASSES_ROOT\" & SearchValue _
                  & "\Shell\Open\Command\" ) Then
      Editor = wshshell.RegRead( "HKEY_CLASSES_ROOT\" & SearchValue _
                                 & "\Shell\Open\Command\" )
      Editor = Left( Editor , InstrRev( Editor , " " ) - 1 )
    End If
  End If
  'Break if something's wrong
  If Editor = "" Then
    If WshParam.Item("Silent") <> "true" Then
      MsgBox "The Extension " & ProjExt & " is not registered correctly." & vbCr _
           & "Please check your registry settings",16,"Projectfiles not registered"
    End If
    Exit Sub
  End If
  If ArgsObj.Count = 0 Then
    If WshParam.Item("Silent") <> "true" Then
      MsgBox "1.Projectfile for Editor." & vbCr _
           & "(Optional Parameters like in a project.ini, separated with Space, grouped with """")", _
             0,"Parameters for Editor refresh"
    End If
    Exit Sub
  ElseIf Not WshFs.FileExists(ArgsObj(0)) Then
    If WshParam.Item("Silent") <> "true" Then
      MsgBox "The Projectfile" & vbCr & ArgsObj(0) & vbCr _
           & "could not be found!",0,"PRefresh Parameter Error"
    End If
    Exit Sub
  End If

  'Parameter-Dictionary init
  For I = 0 to Ubound( Parameter )
    WshParam.Add Parameter(I) , ""
  Next

  If WshParam.Item("Inifile") = "" Then
    WshParam.Item("Inifile")  = IniFilename
  End If
  If WshParam.Item("Pathfile") = "" Then
    WshParam.Item("Pathfile")  = PathsFilename
  End If
  'projectfile is not optional
  ProjectFile = ArgsObj(0)
  If ProjectFile = "" Then
    Exit Sub
  End If
  'direkte Input-Parameter vorziehen
  For I = 1 To ArgsObj.count - 1
    FillParam ( ArgsObj(I) )
  Next

  'restliche defaultparameter laden
  ProjectFileFolder = WshFs.GetFile( ProjectFile ).ParentFolder
  If Right( ProjectFileFolder , 1 ) <> "\" Then
    ProjectFileFolder = ProjectFileFolder & "\"
  End If
  Projectname = Replace( WshFs.GetFile( ProjectFile ).Name , ProjExt , "" , 1 , -1 , 1 )
  If WshParam.Item("Inifile") = "*" Then
    WshParam.Item("Inifile")  = Projectname & InitExt
  End If
  If Pathsname = "*" Then
    Pathsname = Projectname & PathExt
  End If
  If Projectname = "" Then
    Projectname = DefProjName
  End If
  If WshFs.FileExists( ProjectFileFolder & WshParam.Item("Inifile")) Then
    DefaultFile = ProjectFileFolder & WshParam.Item("Inifile")
    ProjectIni = true
  ElseIf WshFs.FileExists( WshParam.Item("Inifile") ) Then
    DefaultFile = WshParam.Item("Inifile")
    ProjectIni = true
  ElseIf WshFs.FileExists( ProjectFileFolder & WshParam.Item("Pathfile") ) Then
    DefaultFile = ProjectFileFolder & WshParam.Item("Pathfile")
  ElseIf WshFs.FileExists( WshParam.Item("Pathfile") ) Then
    DefaultFile = WshParam.Item("Pathfile")
  End If

  'Preset the Values to search in <IniFilename> file here
  If DefaultFile <> "" _
    And Not WshFs.FileExists( DefaultFile ) Then
    MsgBox "A Settingsfile" & vbCr & DefaultFile & vbCr _
         & "was given but couldn't found","PRefresh Parameter Error"
    Exit Sub
  ElseIf DefaultFile <> "" Then
    RightSection = 0
    Set DefaultSetObj = WshFs.OpenTextFile( DefaultFile )
    Do while Not DefaultSetObj.AtEndOfStream
      ActualLine = DefaultSetObj.ReadLine
      'If Sections are defined in <IniFilename> specify unique Keys For Section PSInit
      'The Script will stop reading when an other section begins
      '(Even If nesessary keys couldn't be found)
      'Otherwise the Reading stops when every Key is found
      If InStr( 1 , ActualLine , Section ) = 1 Then
        RightSection = DefaultSetObj.line
      End If
      If ActualLine <> "" Then
        If ProjectIni = True Then
          FillParam ( ActualLine )
        Else
          If WshParam.Item("Delimiter") = "" Then
            WshParam.Item("Paths") = WshParam.Item("Paths") _
                                       & Delimiter & ActualLine
          Else
            WshParam.Item("Paths") = WshParam.Item("Paths") _
                                       & WshParam.Item("Delimiter") & ActualLine
          End If
        End If
      End If
      'Next Section begins > Exiting
      If RightSection > 0 _
        And RightSection <> DefaultSetObj.line _
        And InStr( 1 , ActualLine , InitEnd ) = 1 Then
        Exit Do
      End If
    Loop
    DefaultSetObj.Close
  End If
  If WshParam.Item("Delimiter") = "" Then
    WshParam.Item("Delimiter") = Delimiter
  End If

  'Loading ProjectFile for specific Informations
  Set ProjectObj = WshFs.OpenTextFile( ProjectFile , 1 , False )
  Do while Not ProjectObj.AtEndOfStream
    ActualLine = ProjectObj.Readline
    'If Documentroot is given as parameter (true/false) search there,
    'else search in default Directory
    'Here the wanted Path is craped from Projectfile
    If WshParam.Item("Paths") = "" Then
      Select Case WshParam.Item("DocRoot")
        Case DocRootCases(0)
          If InStr( 1 , ActualLine , DefaultDirDef ) > 0 Then
            DocPath = DocPath & WshParam.Item("Delimiter") & _
                      Replace( ActualLine , DefaultDirDef , "" , 1 , -1 , 1 )
          End If
          If InStr( 1 , ActualLine , DocDirDef ) > 0 Then
            DocPath = DocPath & WshParam.Item("Delimiter") & _
                      Replace( ActualLine , DocDirDef , "" , 1 , -1 , 1 )
          End If
        Case DocRootCases(1)
          If InStr( 1 , ActualLine , DocDirDef ) > 0 Then
            DocPath = Replace( ActualLine , DocDirDef , "" , 1 , -1 , 1 )
          End If
        Case Else
          If InStr( 1 , ActualLine , DefaultDirDef ) > 0 Then
            DocPath = Replace( ActualLine , DefaultDirDef , "" , 1 , -1 , 1 )
          End If
      End Select
    End If
    If WshParam.Item("Extensions") = "" Then
      If InStr( 1 , ActualLine , ExtDef ) > 0 Then
        WshParam.Item("Extensions") = Replace( Replace( ActualLine , _
                                                        ExtDef , _
                                                        "" , 1 , -1 , 1 ) , _
                                               "," , _
                                               Delimiter , 1 , -1 , 1 )
      End If
    End If
    FillParam ( ActualLine )
    'begin of writing Project Tree: Exit Loop
    If ActualLine = WriteBegin Then
      Exit Do
    End If
  Loop
  ProjectObj.Close

  WshParam.Item("Paths")        = LCase( WshParam.Item("Paths") )
  WshParam.Item("Extensions")   = LCase( WshParam.Item("Extensions") )
  WshParam.Item("SubToFolder")  = LCase( WshParam.Item("SubToFolder") )
  WshParam.Item("Sort")         = LCase( WshParam.Item("Sort") )
  WshParam.Item("NoUnknown")    = LCase( WshParam.Item("NoUnknown") )
  WshParam.Item("CutPathAt")    = LCase( WshParam.Item("CutPathAt") )
  WshParam.Item("SubFolders")   = LCase( WshParam.Item("SubFolders") )
  WshParam.Item("ActiveOnly")   = LCase( WshParam.Item("ActiveOnly") )
  WshParam.Item("ExclFolders")  = LCase( WshParam.Item("ExclFolders") )
  WshParam.Item("ExclFiles")    = LCase( WshParam.Item("ExclFiles") )
  WshParam.Item("ExclExtens")   = LCase( WshParam.Item("ExclExtens") )
  WshParam.Item("DocRoot")      = LCase( WshParam.Item("DocRoot") )
  WshParam.Item("Silent")       = LCase( WshParam.Item("Silent") )
  WshParam.Item("IniGen")       = LCase( WshParam.Item("IniGen") )
  WshParam.Item("IncludePaths") = LCase( WshParam.Item("IncludePaths") )

  If WshParam.Item("Paths") = "" Then
    WshParam.Item("Paths") = DocPath
    If WshParam.Item("SubToFolder") = "" Then
      WshParam.Item("SubToFolder")  = "true"
    End If
    If WshParam.Item("Sort") = "" Then
      WshParam.Item("Sort") = "path"
    End If
  ElseIf ProjectIni = False Then
    If WshParam.Item("SubToFolder") = "" Then
      WshParam.Item("SubToFolder")  = "true"
    End If
    If WshParam.Item("Sort") = "" Then
      WshParam.Item("Sort") = "path"
    End If
  ElseIf InStr( WshParam.Item("Paths") , Delimiter ) = 0 _
    And WshParam.Item("IniGen") <> "true" Then
    WshParam.Item("Sort") = Replace( WshParam.Item("Sort") , SortCases(2) , "" , 1 , -1 , 1 )
  End If
  'Check first For found Paths (and their Alias Optional)
  BuildDoubleArray WshParam.Item("Paths") , _
                   WshParam.Item("PathsAlias") , _
                   ProjectPaths

  'Checking For Filters (and their Names optional)
  BuildDoubleArray WshParam.Item("Extensions") , _
                   WshParam.Item("ExtensAlias") , _
                   Extensions
  'Normally list unknown Extensions, too,
  'except Ignore-Option (NoUnknown) is set true
  If WshParam.Item("NoUnknown") <> "true" Then
    ReDim Preserve Extensions(1,UBound( Extensions , 2 ) + 1)
    Extensions(0,UBound( Extensions , 2 )) = "*"
    Extensions(1,UBound( Extensions , 2 )) = "_NoExt"
  End If

  'AliasArrays needed for Sorting
  BuildAliasArray ProjectPaths , PathsAlias
  BuildAliasArray Extensions   , ExtAlias

  'Aliasdictionaries for Translating the Paths to Aliases
  For I = 0 to UBound( Extensions , 2 )
    WshExtAlias.Add  Extensions(0,I) , Extensions(1,I)
  Next
  For I = 0 to UBound( ProjectPaths , 2 )
    WshPathsAlias.Add ProjectPaths(0,I) , ProjectPaths(1,I)
  Next

  'Writing arrays for subfolders to create and exclusions
  BuildSingleArray WshParam.Item("SubFolders")   , CSubFolders
  BuildSingleArray WshParam.Item("ExclFolders")  , FExclFolders
  BuildSingleArray WshParam.Item("ExclFiles")    , FExclFiles
  BuildSingleArray WshParam.Item("ExclExtens")   , FExclExtens
  BuildSingleArray WshParam.Item("IncludePaths") , IncludePaths

  If Instr( WshParam.Item("Extensions") , Delimiter ) = 0 Then
    WshParam.Item("Sort") = Replace( WshParam.Item("Sort") , SortCases(3) , "" , 1 , -1 , 1 )
  End If
  If Instr( WshParam.Item("Sort") , SortCases(3) ) = 0 _
    And WshParam.Item("Extensions")                > "" _
    And ProjectIni = False Then
    WshParam.Item("NoUnknown") = "true"
  End If

End Sub

'*******************************************************************************
' recursive reading a given folder and its contents                            *
'*******************************************************************************

Sub GetSub ( ReadDir )

  Dim Extension , X , J , SubFolder , RelativePath
  Dim FolderObj , DirObj , FileObj

  If WshFs.FolderExists( ReadDir ) Then
    Set FolderObj = WshFs.GetFolder( ReadDir )
    'First checking given Folderexclusions (surely not for Rootfolder)
    If ReadDir <> ActualPath _
      And Matches( FolderObj.Path , FExclFolders ) Then
      Exit Sub
    End If
    SubFolder = GetRelativePath( ActualPath , LCase( FolderObj.Path ) )
    If ReadDir <> ActualPath _
      And UBound( Filter( Folders , SubFolder ) ) = -1 Then
      If ( WshParam.Item("SubToFolder") = "true" _
        Or Matches( FolderObj.Path , CSubFolders ) ) Then
        If Folders(0)      > "" Then ReDim Preserve Folders(UBound( Folders ) + 1)
        If SubFolders(0,0) > "" Then ReDim Preserve SubFolders(2,UBound( SubFolders , 2 ) + 1)
        Folders(UBound( Folders , 1 )) = SubFolder
        SubFolders(0,UBound( SubFolders , 2 )) = SubFolder
        SubFolders(1,UBound( SubFolders , 2 )) = FolderObj.Name
        SubFolders(2,UBound( SubFolders , 2 )) = False
      End If
    End If
    'Recursive Part
    For each DirObj in FolderObj.SubFolders
      GetSub DirObj.Path
    Next
    'File Part
    For each FileObj in FolderObj.Files
      'First checking given Fileexclusions
      RelativePath = GetRelativePath( ActualPath , LCase( FileObj.Path ) )
      Extension    = Right( LCase( FileObj.Name ), _
                            Len( FileObj.Name ) _
                            - InStrRev( FileObj.Name , "." ) _
                            + 1 )
      If Not Matches( FileObj.Path , FExclFiles ) _
        And Not ( WshParam.Item("ActiveOnly") = "true" _
          And UBound( Filter( Relatives , RelativePath ) ) > -1 ) Then
        'Checking extensionfilters
        If WshExtAlias.Exists(Extension) _
          Or WshParam.Item("NoUnknown") <> "true" Then
          'Save only informations, sort and write to file later
          If Relatives(0) > "" Then Redim Preserve Relatives(UBound( Relatives ) + 1)
          If Paths(0,0)   > "" Then Redim Preserve Paths(6,UBound( Paths , 2 ) + 1)
          Relatives(UBound( Relatives )) = RelativePath
          Paths(0,UBound( Paths , 2 ))   = FileObj.Path
          Paths(1,UBound( Paths , 2 ))   = RelativePath
          If WshExtAlias.Exists(Extension) Then
            Paths(2,UBound( Paths , 2 )) = WshExtAlias.Item(Extension)
          Else
            Paths(2,UBound( Paths , 2 )) = WshExtAlias.Item("*")
          End If
          Paths(3,UBound( Paths , 2 ))   = WshPathsAlias.Item(ActualPath)
          If SubFolder = "." Then
            Paths(4,UBound( Paths , 2 )) = ";"
          Else
            Paths(4,UBound( Paths , 2 )) = SubFolder
          End If
          Paths(5,UBound( Paths , 2 ))   = ActualPath
          Paths(6,UBound( Paths , 2 ))   = FileObj.Name
        End If
      End If
    Next
  End If

End Sub

'*******************************************************************************
' Mergesort                                                                    *
'*******************************************************************************

Public Sub MergeSort ( ArrayObj , MirrorObj , Dim1 , Dim2 , Dim3 , Dim4 , First , Last )

  Dim iLeft , iRight , iMid , Diff

  If Dim1 > -1 Then
    If Last = 0 Then
      First = LBound( ArrayObj , 2 )
      Last  = UBound( ArrayObj , 2 )
      ReDim MirrorObj(UBound( ArrayObj , 1 ) , UBound( ArrayObj , 2 ))
    End If
  Else
    If Last = 0 Then
      First = LBound( ArrayObj )
      Last  = UBound( ArrayObj )
      ReDim MirrorObj(UBound( ArrayObj ))
    End If
  End If
  Diff = Last - First
  Select Case Diff
    Case 0
      'Do Nothing, cause it's the same Element
    Case 1
      If SortCheck( ArrayObj , Dim1 , Dim2 , Dim3 , Dim4 , First , Last ) Then
        SortSwitch ArrayObj , Dim1 , First , Last
      End If
    Case Else
      Diff = Diff \ 2 + First
      MergeSort ArrayObj , MirrorObj , Dim1 , Dim2 , Dim3 , Dim4 , First    , Diff
      MergeSort ArrayObj , MirrorObj , Dim1 , Dim2 , Dim3 , Dim4 , Diff + 1 , Last
      ' Merge the resulting halves
      iLeft  = First    ' start of first (left) half
      iRight = Diff + 1 ' start of second (right) half
      iMid   = First    ' start of output (mirror array)
      Do
        If SortCheck( ArrayObj , Dim1 , Dim2 , Dim3 , Dim4 , iLeft , iRight ) Then
          SortCopy ArrayObj , MirrorObj , Dim1 , iRight , iMid
          iRight = iRight + 1
          If iRight > Last Then
            For iLeft = iLeft To Diff
              iMid = iMid + 1
              SortCopy ArrayObj , MirrorObj , Dim1 , iLeft , iMid
            Next
            Exit Do
          End If
        Else
          SortCopy ArrayObj , MirrorObj , Dim1 , iLeft , iMid
          iLeft = iLeft + 1
          If iLeft > Diff Then
            For iRight = iRight To Last
              iMid = iMid + 1
              SortCopy ArrayObj , MirrorObj , Dim1 , iRight , iMid
            Next
            Exit Do
          End If
        End If
        iMid = iMid + 1
      Loop
      For iMid = First To Last
        SortCopy MirrorObj , ArrayObj , Dim1 , iMid , iMid
      Next
  End Select

End Sub

'*******************************************************************************
' Main Procedure                                                               *
'*******************************************************************************

Sub Refresh

  Dim ActualLine , Writing

  'Writing ProjectFile Line per Line to Workfile till Project tree begins
  'After building new Project tree writing rest of ProjectFile'
  'Loading ProjectFile and Create a temporary Workfile
  If WshFs.FileExists( ProjectFile & OldExt ) Then
    WshFs.DeleteFile ProjectFile & OldExt , true
  End If
  WshFs.CopyFile ProjectFile , ProjectFile & OldExt , true
  Set ProjectObj = WshFs.OpenTextFile( ProjectFile & OldExt , 1 , false )
  'Deleting Workfile from previous run
  If WshFs.FileExists( ProjectFile & TempExt ) Then
    WshFs.DeleteFile ProjectFile & TempExt , true
  End If
  Set TempFileObj = WshFs.CreateTextFile( ProjectFile & TempExt , true )
  'Fetching the Projectfolder

  Do while Not ProjectObj.AtEndOfStream
    ActualLine = ProjectObj.Readline
    If Writing = False Then
      If ProjExt       = ".ppr" Then
      If InStr( 1 , ActualLine , "Relative path=" ) > 0 Then
        TempFileObj.Writeline "Relative path=0"
      ElseIf InStr( 1 , ActualLine , "AbsolutePath=" ) > 0 Then
        TempFileObj.Writeline "AbsolutePath=1"
      Else
          TempFileObj.Writeline ActualLine
        End If
      Else
        'If Documentroot is given as parameter (true/false) search there,
        'else search in default Directory
        'Here the wanted Path is craped from Projectfile
        TempFileObj.Writeline ActualLine
      End If
      'Set the begin of writing Project Tree
      If ActualLine = WriteBegin Then
        Writing = true
      End If
    Else
      'Check For the End of Old Project Tree'
      If Left( ActualLine , 1 ) = WriteEnd Then
        Writing = false
        TempFileObj.Writeline ActualLine
      End If
      'On first enter building the Project Tree and Set wrote to true
      'further getting to the end of ProjectFile
      If wrote = false Then
        'Init complete, search begins
        For I = 0 To UBound( ProjectPaths , 2 )
          ActualPath = LCase(ProjectPaths(0,I))
          If WshFs.FolderExists( ActualPath ) Then
            GetSub ActualPath
          End If
        Next
        Sorting
        WriteToProject
        WriteIncludePaths
        wrote = true
      End If
    End If
  Loop
  'release Files
  If Not TempFileObj is Nothing Then
    TempFileObj.Close
  End If
  If Not ProjectObj is Nothing Then
    ProjectObj.Close
  End If

End Sub

'*******************************************************************************
' Copy one entry iun another of Array                                          *
'*******************************************************************************

Sub SortCopy ( ArrayObj1 , ArrayObj2 , Dim1 , Index1 , Index2 )

  If Dim1 > -1 Then
    Dim I
    For I = 0 To UBound( ArrayObj1 , 1 )
      ArrayObj2(I,Index2) = ArrayObj1(I,Index1)
    Next
  Else
    ArrayObj2(Index2) = ArrayObj1(Index1)
  End If

End Sub

'*******************************************************************************
' Read in temp entry from Array                                                *
'*******************************************************************************

Sub SortPop ( ArrayObj , Dim1 , Index1 , HelpArray )

  If Dim1 > -1 Then
    Dim I
    For I = 0 To UBound( ArrayObj , 1 )
      HelpArray(I) = ArrayObj(I,Index1)
    Next
  Else
    HelpArray = ArrayObj(Index1)
  End If

End Sub

'*******************************************************************************
' Write temp entry in Array                                                    *
'*******************************************************************************

Sub SortPush ( ArrayObj , Dim1 , Index1 , HelpArray )

  If Dim1 > -1 Then
    Dim I
    For I = 0 To UBound( ArrayObj , 1 )
      ArrayObj(I,Index1) = HelpArray(I)
    Next
  Else
    ArrayObj(Index1) = HelpArray
  End If

End Sub

'*******************************************************************************
' Switch two entries to Sort an Array                                          *
'*******************************************************************************

Sub SortSwitch ( ArrayObj , Dim1 , Index1 , Index2 )

  If Dim1 > -1 Then
    Dim HelpArray()
    ReDim Preserve HelpArray(Ubound( ArrayObj , 1 ))
  End If
  SortPop  ArrayObj , Dim1 , Index1 , HelpArray
  SortCopy ArrayObj , ArrayObj , Dim1 , Index2 , Index1
  SortPush ArrayObj , Dim1 , Index2 , HelpArray

End Sub

'*******************************************************************************
' Sorting the choosen Case                                                     *
'*******************************************************************************

Sub Sorting

  Dim MirrorArray()

  'Sort the Paths-Array
  'Sorting with MergeSort for best Performance
  Select Case WshParam.Item("Sort")

    'Pathsorting and Extensionsorting
    Case SortCases(0),SortCases(4),SortCases(6)
      MergeSort Paths , MirrorArray , 3 , 2 , 4 , 6 , 0 , 0
    'Extensionsorting and Pathsorting
    Case SortCases(1),SortCases(5),SortCases(7)
      MergeSort Paths , MirrorArray , 2 , 3 , 4 , 6 , 0 , 0
    'Pathsorting grouped by Subfolder (needed in case ExtensionFilter was defined)
    Case SortCases(2)
      MergeSort Paths , MirrorArray , 3 , 4 , 6 , -1 , 0 , 0
    'Extensionsorting grouped by Subfolder (needed in case more than one Paths are searched)
    Case SortCases(3)
      MergeSort Paths , MirrorArray , 2 , 4 , 6 , -1 , 0 , 0
    'sorting, even with Extensionfilter and searching more than one Path
    Case Else
      MergeSort Paths , MirrorArray , 4 , 6 , -1 , -1 , 0 , 0

  End Select

End Sub

'*******************************************************************************
' Write File path to File                                                      *
'*******************************************************************************

Sub WriteFileName ( Stream , Filename , Tabbing )

  If Filename <> "" Then
    Stream.WriteLine Tabbing & Filename
  End If

End Sub

'****************************************************
' Write Folder Close                                *
'****************************************************

Sub WriteFolderClose ( Stream , Foldername )

  'Without extensionfilter no Folder for Unknown Extensions
  'or every extension is unknown in this case
  'If ( UBound( Extensions , 2 ) > 0 _
  '  Or UBound( Extensions , 2 ) = 0 _
  '   And Foldername <> "_NoExt" ) Then
  '  Stream.Writeline "# End Group"
  '  Stream.Writeline ""
  'End If

End Sub

'*******************************************************************************
' Write Folder to File                                                         *
'*******************************************************************************

Sub WriteFolderName ( Stream , FolderName , Tabbing )

  'Without extensionfilter no Folder for Unknown Extensions
  'or every extension is unknown in this case
  If FolderName                   <> "" _
    And ( UBound( Extensions , 2 ) > 0 _
       Or UBound( Extensions , 2 ) = 0 _
        And Foldername            <> "_NoExt" ) Then
    Stream.Writeline Tabbing & "-" & FolderName
  End If

End Sub

'*******************************************************************************
' Write Inifile                                                                *
'*******************************************************************************

Sub WriteIniFile

  Dim InifileObj
  Dim I

  If WshParam.Item("Inifile") = "" Then
    WshParam.Item("Inifile") = Projectname & InitExt
  End If
  If WshFs.FileExists( ProjectFileFolder & WshParam.Item("Inifile") ) Then
    If WshParam.Item("Silent") <> "true" Then
      Answer = MsgBox("The IniGen-Parameter was specified, but" & vbCr _
                      & WshParam.Item("Inifile") & vbCr & "exists!" _
                      & " Overwrite?",36,"Generate Inifile")
      If Answer <> 6 Then
        Exit Sub
      End If
    End If
    WshFs.DeleteFile ProjectFileFolder & WshParam.Item("Inifile") , true
  End If
  Set InifileObj = WshFs.CreateTextFile( ProjectFileFolder _
                                         & WshParam.Item("Inifile") , true )
  InifileObj.Writeline Section
  For I = 0 To UBound( Parameter )
    If Parameter(I)                   <> "IniGen" _
      And Parameter(I)                <> "Inifile" _
      And Parameter(I)                <> "Pathfile" _
      And WshParam.Item(Parameter(I)) <> "" Then
      InifileObj.Writeline Parameter(I) & "=" & WshParam.Item(Parameter(I))
    End If
  Next
  InifileObj.Close

End Sub

'****************************************************
' Write Included Paths                              *
'****************************************************

Sub WriteIncludePaths

  'TempFile.writeline "# Begin Group ""Include Search Paths"",2"
  'For I = 0 To UBound(IncludePaths,1)
  '  TempFile.Writeline "# Begin Group """ & IncludePaths(I) & """,4"
  '  TempFile.Writeline "# Begin Source File,536870920"
  '  TempFile.Writeline "SOURCE=" & IncludePaths(I)
  '  TempFile.Writeline "# End Source File"
  '  TempFile.Writeline "# End Group"
  '  TempFile.Writeline ""
  'Next
  'TempFile.Writeline "# End Group"
  'TempFile.Writeline ""

End Sub

'*******************************************************************************
' Write SortArray to Projectfile                                               *
'*******************************************************************************

Sub WriteToProject

  Dim CurExtAlias , CurPathsAlias , Tabbing , I , J , Creation , Foldertabs , NextRoot

  Select Case ProjExt
    Case ".ppr"
      TempFileObj.Writeline Projectname
    Case ".edp"
      WriteFolderName TempFileObj , Projectname , ""
  End Select

  Select Case WshParam.Item("Sort")
    Case SortCases(0),SortCases(1)
      Tabbing  = "		"
    Case SortCases(2),SortCases(3),SortCases(4),SortCases(5)
      Tabbing  = "	"
    Case Else
      Tabbing  = ""
  End Select

  For I = 0 to UBound( Paths , 2 )

    Select Case WshParam.Item("Sort")

      Case SortCases(0)
        If Paths(3,I) <> CurPathsAlias Then
          WriteFolderClose TempFileObj , CurPathsAlias
          CurPathsAlias = Paths(3,I)
          NextRoot = True
          WriteFolderName TempFileObj , CurPathsAlias , "	"
        End If
        If Paths(2,I) <> CurExtAlias Or NextRoot = True Then
          WriteFolderClose TempFileObj , CurExtAlias
          CurExtAlias = Paths(2,I)
          NextRoot = True
          WriteFolderName TempFileObj , CurExtAlias , "		"
        End If
      Case SortCases(1)
        If Paths(2,I) <> CurExtAlias Then
          WriteFolderClose TempFileObj , CurExtAlias
          CurExtAlias = Paths(2,I)
          NextRoot = True
          WriteFolderName TempFileObj , CurExtAlias , "	"
        End If
        If Paths(3,I) <> CurPathsAlias Or NextRoot = True Then
          WriteFolderClose TempFileObj , CurPathsAlias
          CurPathsAlias = Paths(3,I)
          NextRoot = True
          WriteFolderName TempFileObj , CurPathsAlias , "		"
        End If
      Case SortCases(2),SortCases(4)
        If Paths(3,I) <> CurPathsAlias Then
          WriteFolderClose TempFileObj , CurPathsAlias
          CurPathsAlias = Paths(3,I)
          NextRoot = True
          WriteFolderName TempFileObj , CurPathsAlias , "	"
        End If
        If WshParam.Item("Sort") = SortCases(4) _
          And Paths(2,I) <> CurExtAlias Or NextRoot = True Then
          CurExtAlias = Paths(2,I)
          NextRoot    = True
        End If
      Case SortCases(3),SortCases(5)
        If Paths(2,I) <> CurExtAlias Then
          WriteFolderClose TempFileObj , CurExtAlias
          CurExtAlias = Paths(2,I)
          NextRoot = True
          WriteFolderName TempFileObj , CurExtAlias , "	"
        End If
        If WshParam.Item("Sort") = SortCases(5) _
          And Paths(3,I) <> CurPathsAlias Or NextRoot = True Then
          CurPathsAlias = Paths(3,I)
          NextRoot    = True
        End If

    End Select

    'Closing all finished Subfolders
    For J = UBound( SubFolders , 2 ) To 0
      If InStr( Paths(4,I) , SubFolders(0,J) ) = 0 _
        And SubFolders(2,J) = True Then
        Foldertabs = Left( FolderTabs , Len( FolderTabs ) - 1 )
        WriteFolderClose TempFileObj , SubFolders(1,J) , Tabbing & FolderTabs
        SubFolders(2,J) = False
      End If
    Next

    If NextRoot Then
      For J = 0 To UBound( SubFolders , 2 )
        SubFolders(2,J) = False
      Next
    End If

    NextRoot = False

    For J = 0 To UBound( SubFolders , 2 )
      If SubFolders(0,J) > "" _
        And InStr( Paths(4,I) , SubFolders(0,J) ) = 1 Then
        Foldertabs = FolderTabs & "	"
        If SubFolders(2,J) = False Then
          WriteFolderName TempFileObj , SubFolders(1,J) , Tabbing & FolderTabs
          SubFolders(2,J) = True
        End If
      End If
    Next

    WriteFileName TempFileObj , Paths(0,I) , Tabbing & FolderTabs & "	"

    Foldertabs = ""

  Next

End Sub

'-------------------------------------------------------------------------------
' Functions                                                                    -
'-------------------------------------------------------------------------------

'*******************************************************************************
' Build a 2 dimensional Array for Aliases (Reverse to) from Input Array        *
'*******************************************************************************

Function BuildAliasArray ( Arrayname , AliasArray )

  Dim I , J , Found
  Dim MirrorArray()

  ReDim Preserve AliasArray(1,0)

  For I = 0 To UBound( Arrayname , 2 )
    Found = False
    For J = 0 To UBound( AliasArray , 2 )
      If InStr( AliasArray(0,J) , Arrayname(1,I) ) > 0 Then
        AliasArray(1,J) = AliasArray(1,J) _
                          & WshParam.Item("Delimiter") _
                          & Arrayname(0,I)
        Found = True
        Exit For
      End If
    Next
    If Found = False Then
      ReDim Preserve AliasArray(1,UBound( AliasArray , 2 ) + 1)
      AliasArray(0,UBound( AliasArray , 2 )) = Arrayname(1,I)
      AliasArray(1,UBound( AliasArray , 2 )) = Arrayname(0,I)
    End If
  Next
  MergeSort AliasArray , MirrorArray , 0 , -1 , -1 , -1 , 0 , 0

End Function

'*******************************************************************************
' Build a 2 dimensional Array from 2 Strings                                   *
'*******************************************************************************

Function BuildDoubleArray ( String1 , String2 , Arrayname )

  Dim Temp1 , Temp2 , Temp3 , I

  String1 = Replace( String1 , "*" , "" , 1 , -1 , 1 )
  String1 = Replace( Replace( String1 , "," , ";" , 1 , -1 , 1 ) , ";;" , ";" , 1 , -1 , 1 )
  String2 = Replace( Replace( String2 , "," , ";" , 1 , -1 , 1 ) , ";;" , ";" , 1 , -1 , 1 )
  Temp1   = Split( LCase( String1 ) , ";" )
  Temp2   = Split( String2 , ";" )

  ReDim Preserve Arrayname(1,UBound( Temp1 ))

  For I = 0 To UBound(Temp1)
    If String2 > "" And UBound(Temp2) >= I Then
      If Temp2(I) <> "" Then
        Temp3 = Temp2(I)
      End If
    ElseIf InStr( Temp1(I) , "\" ) > 0 Then
      'Create Pathsalias
      Temp3 = CutPath( Temp1(I) )
    ElseIf InStr( Temp1(I) , "." ) > 0 Then
      'Take Extension without Dot
      Temp3 = Replace( Temp1(I) , "." , "" , 1 , -1 , 1 )
    End If
    Arrayname(0,I) = Temp1(I)
    Arrayname(1,I) = Temp3
  Next

End Function

'*******************************************************************************
' Build a 2 dimensional Array from 2 Strings                                   *
'*******************************************************************************

Function BuildSingleArray ( String1 , Arrayname )

  Dim Delimiter1

  Delimiter1 = WshParam.Item("Delimiter")

  String1 = Replace( Replace( LCase( String1 ) , _
                              "," , _
                              Delimiter1 , 1 , -1 , 1 ) , _
                     Delimiter1 + Delimiter1 , _
                     Delimiter1 , 1 , -1 , 1 )
  Arrayname = Split( String1 , Delimiter1 )

End Function

'*******************************************************************************
' Cutting Path for Projectfoldername                                           *
'*******************************************************************************

Function CutPath ( Path )

  Dim Cutter

  If WshParam.Item("CutPathAt") = "" Then
    Cutter = LCase(Projectname)
  Else
    Cutter = WshParam.Item("CutPathAt")
  End If
  If InStr( Path , Cutter ) > 0 Then
    CutPath = Right( Path , _
                     Len( Path ) _
                     - Instr( Path , Cutter ) _
                     - Len( Cutter ) _
                     + 1 )
    If CutPath = "" Then
      CutPath = Left( Path , _
                       Len( Path ) _
                       - InstrRev( Path , Cutter ) )
    End If
  End If
  If CutPath = "" Then
    CutPath = Path
  End If

  CutPath = Replace( Replace( CutPath , "\" , "" , 1 , -1 , 1 ) , _
                     ":" , "" , 1 , -1 , 1 )

End Function

'*******************************************************************************
' Fill Dictionary with the needed Parameters                                   *
'*******************************************************************************

Function FillParam ( Argument )

  Const Delimiter1 = "="

  Dim Temp1 , Temp2

  If Argument > "" Then

    If InStr( 1 , Argument , Delimiter1 ) > 0 Then
      Temp1 = Left( Argument , InStr( 1 , Argument , Delimiter1 ) - 1 )
      Temp2 = Right( Argument , Len( Argument ) _
                                - InStr( 1 , Argument , Delimiter1 ) )
    Else
      Temp1 = Argument
      Temp2 = ""
    End If

    If WshParam.Exists(Temp1) _
      And WshParam.Item(Temp1) = "" Then
      If WshParam.Item(Temp1) = "" Then
        WshParam.Item(Temp1) = Temp2
      Else
        WshParam.Item(Temp1) = WshParam.Item(Temp1) _
                               & WshParam.Item("Delimiter") _
                               & Temp2
      End If
    End If

  End If

End Function

'*******************************************************************************
' returns a Path relative to a given Root                                      *
'*******************************************************************************

Function GetRelativePath ( RootPath , FullPath )

  'Only needed to get relative Paths to the chosen one
  Dim TempPath

  'Drive is equal
  If Left( RootPath , 3 ) = Left( FullPath , 3 ) Then
    If InStr( 1 , FullPath , RootPath ) = 1 Then
      'Fullpath is a Subfolder of Path
      GetRelativePath = Replace( FullPath , RootPath , "." , 1 , -1 , 1 )
    Else
      'Cutting Driveletter and repeat adding ..\ to Rootfolder of Drive
      GetRelativePath = Right( FullPath , Len( FullPath ) - 3 )
      TempPath = RootPath
      Do While TempPath > Left( RootPath , 3 )
        TempPath = Left( TempPath , InStrRev( TempPath , "\" ) )
        GetRelativePath = "..\" & GetRelativePath
      Loop
    End If
  Else
    ' Other Drive: have to be returned absolutely'
    GetRelativePath = FullPath
  End If

End Function

'*******************************************************************************
' Checking if Path matches a given List                                        *
'*******************************************************************************

Function Matches ( Path , Excludelist )

  Dim Exclusion , Name

  Name = Right( Path , Len( Path ) - InStrRev( Path , "\" ) )

  For Each Exclusion In Excludelist
    If Exclusion = Name _
      Or ( Left( Exclusion , 1 ) = "*" _
       And InStr( Name , Replace( Exclusion , "*" , "" , 1 , -1 , 1) ) > 1 ) _
      Or ( InStr( Exclusion , "\" ) > 0 _
       And Exclusion = Path ) _
      Or ( InStr( Exclusion , "\" ) > 0 _
       And Left( Exclusion , 1 ) = "*" _
       And InStr( Path , Replace( Exclusion , "*" , "" , 1 , -1 , 1 ) ) > 1 ) Then
      Matches = True
      Exit For
    End If
  Next

End Function

'*******************************************************************************
' Checking whether Registrykey exists or not                                   *
'*******************************************************************************

Function RegExists( Wert )

  Dim Test

  On Error Resume Next
  Test = wshshell.RegRead( Wert )
  RegExists = ( Err.number = 0 )
  On Error Goto 0

End Function

'*******************************************************************************
' Check if 2 Elements have to be Sorted                                        *
'*******************************************************************************

Function SortCheck ( ArrayObj , Dim1 , Dim2 , Dim3 , Dim4 , Index1 , Index2 )

  SortCheck = False

  If Dim1 > -1 Then
    If ArrayObj(Dim1,Index1) > ArrayObj(Dim1,Index2) Then
      SortCheck = True
    End If
    If Dim2 > -1 Then
      If    ArrayObj(Dim1,Index1) = ArrayObj(Dim1,Index2) _
        And ArrayObj(Dim2,Index1) > ArrayObj(Dim2,Index2) Then
        SortCheck = True
      End If
    End If
    If Dim3 > -1 Then
      If    ArrayObj(Dim1,Index1) = ArrayObj(Dim1,Index2)   _
        And ArrayObj(Dim2,Index1) = ArrayObj(Dim2,Index2)   _
        And ArrayObj(Dim3,Index1) > ArrayObj(Dim3,Index2) Then
        SortCheck = True
      End If
    End If
    If Dim4 > -1 Then
      If    ArrayObj(Dim1,Index1) = ArrayObj(Dim1,Index2)   _
        And ArrayObj(Dim2,Index1) = ArrayObj(Dim2,Index2)   _
        And ArrayObj(Dim3,Index1) = ArrayObj(Dim3,Index2)   _
        And ArrayObj(Dim4,Index1) > ArrayObj(Dim4,Index2) Then
        SortCheck = True
       End If
    End If
  Else
    If ArrayObj(Index1) > ArrayObj(Index2) Then
      SortCheck = True
    End If
  End If

End Function