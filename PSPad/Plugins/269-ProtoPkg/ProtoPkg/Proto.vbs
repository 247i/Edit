'===============================================================================
'	Filename		  : Proto.vbs
'	DESCRIPTION		: Copies a code snippet from a file in the "Proto" subfolder
'					        directly under PsPad's root directory.
'                 The file must be named "x.proto", where "x" is the file
'                 extension of the file being edited when Proto is invoked.
'                 Proto will use the word just typed (if any), or will ask for
'                 a keyword.

'	Created			  : 07 July 2010
'	Created by		: Russ Williams

'	Comments		  :
'===============================================================================
Option Explicit

Const MODULE_NAME  = "InsertProto"  'Name of subroutine to be called
Const MODULE_VER   = "0.1"          'Version number
Const MODULE_TITLE = "Proto"        'Name to show in script menu
Const TEST = False

' Subroutine "Init" Is required 
' It is called automatically during initialization To create menu items.
Sub Init
  addMenuItem MODULE_TITLE, "RWAddons", MODULE_NAME, "CTRL+ALT+P"   'Add in "RWAddons" submenu
End Sub

' Subroutine to insert prototype snippet
Sub InsertProto()
	Dim objEditor, objProtoEditor, objFSO
	Dim Proto, ProtoPath, ProtoFile, strFilePath
  Dim ProtoLine, ProtoKw, ProtoBlock, index
  Dim fso, f1, ts, s, x, found
  Const ForReading = 1  'Code for FSO OpenTextFile

	Set objEditor = newEditor() 'Create new edit object
  objEditor.assignActiveEditor()  'activate it
  strFilePath = objEditor.fileName  'Get edit file name

	ProtoPath = AppPath() & "Proto" 'Path to Proto folder
  ProtoFile = ProtoPath & "\" & GetExtensionName(strFilePath) & ".proto"
  found = False

	Set objFSO = CreateObject("Scripting.FileSystemObject") 'Create a file system object
	If objFSO.FolderExists(ProtoPath) Then 'Folder exists
    If objFSO.FileExists(ProtoFile) Then 'File exists
      strFilePath = ProtoFile
    Else
   		strFilePath = BrowseForFile("Please select the Prototype  file to use:", ProtoPath) 'Prompt for Proto choice
    End if
		Set objFSO = Nothing
	Else                             'Give up
		MsgBox "Aborting..." & vbNewLine & vbNewLine & "Could not locate the Prototypes file:" & ProtoPath, 16, MODULE_TITLE
		Set objFSO = Nothing
		Exit Sub
	End if

	If (NOT IsNull(strFilePath)) AND Len(strFilePath) Then
  ' Proto file name is in strFilePath, look up snippet name
    x = objEditor.caretX  'preset to cursor x-coord
    ProtoBlock = "" 'preset to null string'
    Set objFSO = CreateObject("Scripting.FileSystemObject") 'create fso
    Set ts = objFSO.OpenTextFile(strFilePath, ForReading) 'create textstring
    If x > 1 Then
      objEditor.Command("ecSelWordLeft")  'select word to left of cursor
      ProtoKw = UCase(objEditor.SelText())  'copy to ProtoKw
      If objEditor.SelLength > 0 then 'selection is not null
        objEditor.SelText("") 'remove selected text
        x = objEditor.caretX 'save cursor now
      End If
    Else
      ProtoKw = UCase(InputBox("Enter a keyword:"))
    End If
    ' Search for desired keyword (preceded by ":")
    Do Until Left(ProtoLine,Len(ProtoKw)+1) = ":" & ProtoKw
      If ts.AtEndOfStream Then Exit Do
      ProtoLine = ts.ReadLine
      If TEST Then Echo "Searching..." & ProtoLine
    Loop
    If NOT ts.AtEndOfStream Then
      found = True
      Index = InStr(ProtoLine, " ")
      If Mid(ProtoLine,Index+1,1) = "=" Then Index = Index + 2
      ProtoLine = Right(ProtoLine, Len(ProtoLine) - index)
      Do Until Left(ProtoLine,1) = ":" OR ts.AtEndOfStream
        If TEST Then Echo "Reading..." & ProtoLine
        ProtoBlock = ProtoBlock & ProtoLine & vbNewLine
        ProtoLine = ts.ReadLine  ' Read next line'
      Loop
		End If
    If x > Len(objEditor.lineText) Then x = Len(objEditor.lineText) 'Having problems here'
    If x < 1 Then x = 1 'just in case
    objEditor.lineText Left(objEditor.lineText,x - 1) & ProtoBlock & Right(objEditor.lineText,Len(objEditor.lineText)+1-x)
    ts.Close          ' Close the file'
    set ts = nothing
    set objFSO = nothing
    If NOT found Then MsgBox "Keyword not found: " & UCase(ProtoKw), 16, MODULE_TITLE
	Else                              'Give up
		MsgBox "Aborting..." & vbNewLine & vbNewLine & "No prototype chosen.", 16, MODULE_TITLE
	End If

	Set objProtoEditor = nothing
	Set objEditor = nothing

End Sub

Function BrowseForFolder(strPrompt, strRootPath)
	On Error Resume Next
	Const BIF_BROWSEINCLUDEFILES = &H4000 	' The browse dialog box will display files as well as folders.
	Const BIF_NONEWFOLDERBUTTON = &H0200 	' Do not include the New Folder button in the browse dialog box.
	Dim objShell, objFolder, intColonPos
	Set objShell = CreateObject("Shell.Application")

	Set objFolder = objShell.BrowseForFolder(&H0, strPrompt, BIF_NONEWFOLDERBUTTON + BIF_BROWSEINCLUDEFILES, strRootPath)

	If Err.Number > 0 Then
		MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
	End If

	If objFolder is Nothing Then
		' User clicked Cancel
		BrowseForFolder = Null
	Else
		BrowseForFolder = objFolder.ParentFolder.ParseName(objFolder.Title).Path

		If Err.Number > 0 Then
			BrowseForFolder = Null
		End If
	End If
	If Err.Number > 0 Then
		MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
	End If

	If objFolder is Nothing Then
		' User clicked Cancel
		BrowseForFolder = Null
	Else
		BrowseForFolder = objFolder.ParentFolder.ParseName(objFolder.Title).Path

		If Err.Number > 0 Then

			BrowseForFolder = Null 'will be null if no special case applies

			'If selected folder is a drive, it will have a colon e.g. C:\
			intColonPos = InStr(objFolder.Title, ":")

			If intColonPos > 0 Then
				BrowseForFolder = Mid(objFolder.Title, intColonPos - 1, 2) & "\"
			End If
		End If
	End If
End Function


Function BrowseForFile(strPrompt, strRootPath)
	On Error Resume Next
	Dim objDialog
    Set objDialog = CreateObject("UserAccounts.CommonDialog")

	' Only Windows XP can create this object, otherwise call BrowseForFolder
	If Err.Number > 0 Then
		BrowseForFile = BrowseForFolder(strPrompt, strRootPath)
		Exit Function
	End If

	objDialog.Filter = "Proto Files|*.proto|All Files|*.*"
    objDialog.FilterIndex = 1
    objDialog.InitialDir = strRootPath
    objDialog.ShowOpen

	BrowseForFile = objDialog.FileName
	If Err.Number > 0 Then
		MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		BrowseForFile = Null
	End If
End Function

' GetExtensionName extracts the extension part of a filespec
Function GetExtensionName(strFileName)
  Dim count
  count = InStr(StrReverse(strFileName),".") - 1
  GetExtensionName = Right(strFileName,count)
End Function

' AppPath gets the path for the editor module
Function AppPath()
	Dim strPath
	strPath = modulePath()
	If Len(strPath) = 0 Then Exit Function 'Couldn't get a path
	AppPath = Left(strPath, InStrRev(strPath, "\Script\VBScript\"))
End Function
