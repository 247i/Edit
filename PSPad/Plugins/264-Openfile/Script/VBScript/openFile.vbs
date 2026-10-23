'*******************************************************************************
'Filename    : openfile.vbs                                                    *
'Lastmodified: 11.06.2010 14:43:44                                             *
'Description : Gets the HREF|SRC|ACTION|FILE|VIRTUAL attribute of the tag      *
'                that the cursor is resting in and opens the file. If a file   *
'                path or a part thereof is selected then this will be treated  *
'                as a relative path on the local machine/network               *
'                In the case of virtual paths the user is presented with a     *
'                browse folder dialog. Cannot open remote files.               *
'Created     : 21 September 2005                                               *
'Created by  : James Swan                                                      *
'Modified by : Holger Biehl                                                    *
'Tested with : PSPad Build 2390                                                *
'*******************************************************************************

'*******************************************************************************
'Note        : Since this script uses Regular Expressions, VBScript version 5  *
'            is needed to run it. This is the version of the VBScript          *
'            scripting engine released with Internet Explorer 5.               *
'*******************************************************************************

'*******************************************************************************
'  Version History  :                                                          *
'  1.000     Added Functions to search in a propath.txt,                       *
'            a editor.ini or the opened projectfile for a marked File          *
'            Cursorposition is enough to grab the Filename from Text           *
'            (marking not nesessary)                                           *
'            Tested with Progressinclude files                                 *
'  1.001     Bug fixed to not open files case sensitve, thanks to Tonino       *
'  1.010     Openfile uses another editor.ini entry IncludePaths=              *
'            (to act indepentant from Projectrefresh)                          *
'  1.011     searching for <project>.ini/.path like in projectrefresh          *
'*******************************************************************************

Public Const PathSearch  = "Propath="
Public Const SDelimiter  = "Delimiter="

Public PathInclude
Public objEditor
Public IncludePaths()
Public Delimiter

Const MODULE_NAME   = "OpenInclude"
Const MODULE_VER    = "1.010"
Const DIALOG_TITLE  = "Open Include"
Const MODULE_NAME2  = "OpenInclude2"
Const DIALOG_TITLE2 = "Open Include2"

'Delimiter is needed to split Entries in editor.ini (default)
'can be overwritten by Delimiter= in editor.ini
Delimiter   = ";"
PathInclude = "IncludePaths="

'*******************************************************************************
' Create the Menue                                                             *
'*******************************************************************************

Sub Init
    addMenuItem DIALOG_TITLE, "", MODULE_NAME, "ALT+O"
    addMenuItem DIALOG_TITLE2, "", MODULE_NAME2, "CTRL+ALT+O"
End Sub

'*******************************************************************************
' Mainprocedure                                                                *
'*******************************************************************************
Sub OpenInclude2

  PathInclude = "IncludePaths2="
  OpenInclude

End Sub

Sub OpenInclude
  Dim fs, objLinkEditor
  Dim strFileName, strActivePath, strFilePath, strTag, strFileContent
  Dim BACKSLASH, SLASH
  Dim i, intSelStart, intSelLength
  Dim blnVirtual
  Dim arrPath
  Dim Opened
  Const ICON_STOP = 16
  BACKSLASH = Chr(92)
  SLASH = Chr(47)
  blnVirtual = true

  Set objEditor = newEditor()
  Set fs = CreateObject("Scripting.FileSystemObject")
  With objEditor
    .assignActiveEditor()
    '-------------------------------------------------------------------
    ' get the name of the file to open
    '-------------------------------------------------------------------
    ' get selected text
    strFileName = .selText()
    ' get the contents of the "source" attribute of the tag that the cursor is resting in
    If Len(strFileName) = 0 Then
      strTag = GetTagAtCursor()
      strFileName = GetHTMLAttribute(strTag)
      If Len(strTag) = 0 then
        MsgBox "Aborting..." & vbCrLf & vbCrLf & _
        "Actual Line is empty"
        Exit Sub
      Else
        If len(strFileName) = 0 then
          strFileName = strTag
        End If
      End If
    End if
    If InStr(strFileName, ":") = 0 and InStr(strFileName, "\\") = 0 and InStr(strFileName, ".\") = 0 Then
      blnVirtual = True
    End if
    '-------------------------------------------------------------------
    If Len(strFileName) > 0 Then
      If Left(strFileName, 7) = "mailto:" Then
        MsgBox "Aborting..." & vbCrLf & vbCrLf &_
        "Can't open a ""mailto:"" link." & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
        Exit Sub
      End if
      If InStr(strFileName, "http://") > 0 Then
        MsgBox "Can't open server files" & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
      Elseif blnVirtual = true then
        ProPathTxt = Left(projectFileName(),InStrRev(projectFileName(),BACKSLASH)) & "propath.txt"
        ProjectIni = Left(projectFileName(),InStrRev(projectFileName(),BACKSLASH)) & "editor.ini"
        Opened = false
        If Not fs.FileExists(ProjectIni) Then
          ProjectIni = Replace(projectFileName(),".ppr",".ini")
        End If
        If Not fs.FileExists(ProPathTxt) Then
          ProPathTxt = Replace(projectFileName(),".ppr",".path")
        End If
        If fs.FileExists(ProjectIni) Then
          set PathFile = fs.OpenTextFile(ProjectIni,1,true)
          If GetProject( PathFile ) = True Then
            If UBound(IncludePaths,1) > 0 Then
              For I = 0 to UBound(IncludePaths,1)
                strActivePath = IncludePaths(I) & BACKSLASH
                if fs.FileExists(strActivePath & strFileName) then
                  Opened = true
                  Exit For
                End If
              Next
            End If
          End If
          PathFile.Close
        Elseif fs.FileExists(ProPathTxt) Then
          set PathFile = fs.OpenTextFile(ProPathTxt,1,true)
          do while not PathFile.AtEndOfStream
            strActivePath = PathFile.Readline & BACKSLASH
            if fs.FileExists(strActivePath & strFileName) then
              Opened = true
              Exit Do
            End If
          loop
          PathFile.Close
        Elseif fs.FileExists(projectFileName()) Then
          Set PathFile = fs.OpenTextFile(projectFileName(),1,true)
          Do While Not PathFile.AtEndOfStream
            Searchline = PathFile.Readline
            if InStr(1,Searchline,Replace(strFileName,"/","\"),vbTextCompare) > 0 Then
              strFileName = Replace(Searchline,"	","")
              strActivePath = Left(projectFileName(),InStrRev(projectFileName(),BACKSLASH))
              if fs.FileExists(strActivePath & strFileName) then
                Opened = true
                Exit Do
              End If
            End If  
          Loop
          PathFile.Close
        End If
        If not Opened Then
          strActivePath = Null
          'strActivePath = BrowseForFolder("Cannot currently resolve virtual paths..." & _
          'vbCrLf & "Please select the include file folder manual.")
          If IsNull(strActivePath) Then
            Exit Sub
          Else
            strActivePath = strActivePath & BACKSLASH
          End if
          If InStr(strFileName, SLASH) > 0 Then
            ' grab filename from path
            strFileName = Mid(strFileName, InStrRev(strFileName, SLASH) + 1)
          End if
        End if
      Elseif blnVirtual = false then
        ' get the file path of the active document
        strActivePath = ExtractFilePath(objEditor.fileName())
        If InStr(strFileName, "../") > 0 Then
          ' trim trailing backslash from active document path
          strActivePath = Left(strActivePath,Len(strActivePath) - 1)
          ' convert the path into an array of drive/directory names
          arrPath = Split(strActivePath, BACKSLASH)
          strActivePath = ""
          ' go one level up in active document path for each "../" in relative file path
          For i = 0 To UBound(arrPath) - MatchCount("\.\.\/", strFileName)
            strActivePath = strActivePath & arrPath(i) & BACKSLASH
          Next
          ' strip leading "../":s from relative file path
          strFileName = Replace(strFileName, "../", "")
        End If
      End If
      ' replace slashes with backslashes in relative file path
      strFileName = Replace(strFileName, SLASH, BACKSLASH)
      ' set absolute path for selected file
      strFilePath = strActivePath & strFileName
      ' strip querystring
      If InStr(strFilePath, "?") > 0 Then
        strFilePath = Mid(strFilePath, 1, InStr(strFilePath, "?")-1)
      End if
      ' strip anchor identifier
      If InStr(strFilePath, "#") > 0 Then
        strFilePath = Mid(strFilePath, 1, InStr(strFilePath, "#")-1)
      End if
      If fs.FileExists(strFilePath) Then
        Set objLinkEditor = newEditor()
        objLinkEditor.openFile(strFilePath)
        Set objLinkEditor = Nothing
      Else
        msgbox strActivePath & blnVirtual
        MsgBox "Aborting..." & vbCrLf & vbCrLf &_
        "Could not find the file." & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
        Exit Sub
      End if
      Set fs = Nothing
    Else
      MsgBox "Aborting..." & vbCrLf & vbCrLf &_
      "No file name selected." & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
      Exit Sub
    End if
  End With
  Set objEditor = Nothing

End Sub

'*******************************************************************************
' Counts Matches of a String within another
'*******************************************************************************

Function MatchCount ( patrn , str )

  Dim objRegEx, objMatch, objMatches

  Set objRegEx = New RegExp
    With objRegEx
      .Pattern = patrn
      .IgnoreCase = True
      .Global = True
      Set objMatches = .Execute(str)
    End With
  Set objRegEx = Nothing
  MatchCount = objMatches.Count

End Function

'*******************************************************************************
' Bring up a Dialog for Folderselection                                        *
'*******************************************************************************

Function BrowseForFolder ( strPrompt )

  Const BIF_NONEWFOLDERBUTTON = &H0200   ' Do not include the New Folder button in the browse dialog box.
  Const BSF_DESKTOP = 0           ' Desktop is the root directory. With BIF_returnonlyfsdirs circumvents problem with OK-button
  Dim objShell, objFolder, intColonPos, objWshShell

  On Error Resume Next
  Set objShell = CreateObject("Shell.Application")
  Set objWshShell = CreateObject("WScript.Shell")
  Err.Clear
  Set objFolder = objShell.BrowseForFolder(&H0, strPrompt, BIF_NONEWFOLDERBUTTON, BSF_DESKTOP)
  If Err.Number > 0 Then
    MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
  End If
  If objFolder is Nothing Then
    BrowseForFolder = Null
  Else
    BrowseForFolder = objFolder.ParentFolder.ParseName(objFolder.Title).Path
    If Err.Number > 0 Then
      BrowseForFolder = Null
      If objFolder.Title = "Desktop" Then
        If NOT objWshShell is Nothing Then
          BrowseForFolder = objWshShell.SpecialFolders("Desktop")
        End If
      End If
      intColonPos = InStr(objFolder.Title, ":")
      If intColonPos > 0 Then
        BrowseForFolder = Mid(objFolder.Title, intColonPos - 1, 2) & "\"
      End If
    End If
  End If

End Function

'*******************************************************************************
' Search for HTML-Tags                                                         *
'*******************************************************************************

Function GetHTMLAttribute ( strTag )

  Dim objRegExp, attribute, sourceAttributes
  Dim strFileName

  Set objRegExp = new RegExp
  With objRegExp
    .Global = True
    .IgnoreCase = True
    .Pattern = "(HREF|SRC|ACTION|FILE|VIRTUAL)=""*([^"">\s]*)"
    Set sourceAttributes = .Execute(strTag)
    For Each attribute in sourceAttributes
      strFileName = objRegExp.Replace(attribute.Value,"$2")
      If InStr(strFileName, "?") > 0 Then
        strFileName = Left(strFileName, InStr(strFileName, "?")-1)    ' strip any querystring
      End If
      If InStr(strFileName, "<%") > 0 Then
        strFileName = Left(strFileName, InStr(strFileName, "<%")-1)   ' strip any inline asp
      End If
    Next
  End With
  Set sourceAttributes = nothing
  Set objRegExp = nothing
  GetHTMLAttribute = strFileName

End Function

'*******************************************************************************
' Gets the Filename even when not marked before                                *
'*******************************************************************************

Function GetTagAtCursor ()

  Dim strLine, strChar, strReturn
  Dim i, intPosX, intStartPos, intEndPos
  Dim objEditor

  Set objEditor = newEditor()
  With objEditor
    .assignActiveEditor()
    strLine = .lineText()
    intPosX = .caretX()
  End With
  intStartPos = 1
  intEndPos = Len(strLine)
  For i = intPosX To Len(strLine) Step 1
    strChar = (Mid(strLine, i, 1))
    If (strChar = ">" AND strLastChar <> "%") or strChar = " " or strChar = "}" or strChar = """" or strChar = "'" Then
      if strLastChar = "}" then
        intEndPos = i - 2
      elseif strChar = "}" or strChar = " " or strChar = """" or strChar = "'" then
        intEndPos = i - 1
      else
        intEndPos = i
      End If
      Exit For
    End If
    strLastChar = strChar
  Next
  For i = intEndPos To 1 Step -1
    strChar = (Mid(strLine, i, 1))
    If (strChar = "<" AND strLastChar <> "%") or strChar = " " or strChar = "{" or strChar = """" or strChar = "'" Then
      if strLastChar = "{" then
        intStartPos = i + 2
      elseif strChar = "{" or strChar = " " or strChar = """" or strChar = "'" then
        intStartPos = i + 1
      else
        intStartPos = i
      End If
      Exit For
    End If
    strLastChar = strChar
  Next
  GetTagAtCursor = Mid(strLine, intStartPos, intEndPos - intStartPos + 1)

End Function

'*******************************************************************************
' Gets ParentFolder?                                                           *
'*******************************************************************************

Function ExtractFilePath ( strPath )

    If Len(strPath) = 0 Then
      Exit Function                                    ' input string is empty
    Else
      strPath = Replace(strPath, Chr(47), Chr(92))     ' convert backslashes to forward slashes
      If InStr(1, strPath, Chr(92)) = 0 Then
        Exit Function                                  ' string contains no forward slashes
      End If
    End If
    ExtractFilePath = Left(strPath, InStrRev(strPath, Chr(92)))

End Function

'*******************************************************************************
' Reading ProjectFile entries to Arrays                                        *
'*******************************************************************************

Function GetProject ( ProjectFile )

  'A given ProjectFile is searched For the searchparameters here
  Dim PathFound
  Dim InclFound

  'Preset the Values to search in editor.ini file here
  RightSection = 0

  'Searching Line by Line to End or Found all needed'
  Do while Not ProjectFile.AtEndOfStream
    Line = ProjectFile.ReadLine
    'If Sections are defined in editor.ini specify unique Keys For Section PSInit
    'The Script will stop reading when an other section begins
    '(Even If nesessary keys couldn't be found)
    'Otherwise the Reading stops when every Key is found
    If InStr(1,Line,Section,1) = 1 Then
      RightSection = ProjectFile.line
    end If
    PathFound   = CheckInitEntry ( Line , PathSearch , PathFound )
    Delimiter   = CheckInitEntry ( Line , SDelimiter , Delimiter )
    InclFound   = CheckInitEntry ( Line , PathInclude , InclFound )

    'Next Section begins > Exiting
    If RightSection > 0 And RightSection <> ProjectFile.line And InStr(1,Line,"[",1) = 1 Then
      Exit Do
    End If
    'All Keys found > Exiting
    If PathFound       <> "" _
      And Delimiter    <> "" _
      And InclFound    <> "" Then
      Exit Do
    End If
  Loop

  'Check first For found Paths (and their Alias Optional)
  'Writing Array for Included Paths to search in
  If InclFound <> "" Then
    GetProject = BuildSingleArray ( IncludePaths , InclFound )
  Else
    GetProject = BuildSingleArray ( IncludePaths , PathFound )
  End If

End Function

'*******************************************************************************
' Build a 1 dimensional Array from 1 Strings                                   *
'*******************************************************************************

Function BuildSingleArray ( Arrayname , String1 )

  Dim Temp1

  ReDim Preserve Arrayname(0)
  Do While String1 <> ""
    If InStr(1,String1,Delimiter) > 0 Then
      Temp1 = Left(String1,InStr(1,String1,Delimiter) - 1)
    Else
      Temp1 = ""
    End If
    'Not the last entry: cutting
    If Temp1 <> "" Then
      String1 = Replace(String1,Temp1 & Delimiter,"",1,1)
    Else
      Temp1 = String1
      String1 = ""
    End If
    'Write in the Array for later use
    If Temp1 <> "" Then
      Arrayname(UBound(Arrayname,1)) = Temp1
    End If
    If String1 = "" Then
      Exit Do
    End If
    ReDim Preserve Arrayname(UBound(Arrayname,1) + 1)
  Loop
  BuildSingleArray = true

End Function

'*******************************************************************************
' Checking a line of editor.ini for a given option                             *
'*******************************************************************************

Function CheckInitEntry ( String1 , InitEntry , Stored )

  If InStr(1,String1,InitEntry,1) = 1 Then
    CheckInitEntry = Right(String1,len(String1) - len(InitEntry))
  ElseIf Stored <> "" Then
    CheckInitEntry = Stored
  Else
    CheckInitEntry = ""
  End If

End Function

