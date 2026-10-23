'*******************************************************************************
'	FILENAME		: save_snippet.vbs
'	DESCRIPTION		: Inserts a code snippet from a directory named "Snippets"
'					  directly under PsPad's root directory.

'	CREATED			: 26 September 2005
'	CREATED BY		: James Swan

'	NOTE			: This script has only been tested on Windows 2000 and XP

'	You may distribute this script freely, but please keep this header intact.
'*******************************************************************************
Option Explicit

Const MODULE_NAME  = "InsertSnippet"
Const MODULE_VER   = "0.1"
Const MODULE_TITLE = "Insert Snippet"

Sub Init
    addMenuItem MODULE_TITLE, "Utilities", MODULE_NAME, "CTRL+ALT+INS"
End Sub

Sub InsertSnippet()
	Dim objEditor, objSnippetEditor, objFSO
	Dim strSnippet, strSnippetPath, strFilePath, strSnippetName

	Set objEditor = newEditor()

	With objEditor
		.assignActiveEditor()
		strSnippetPath = AppPath() & "Snippets"
	End With

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If NOT objFSO.FolderExists(strSnippetPath) Then
		On error Resume next
			objFSO.CreateFolder(strSnippetPath)
		On error GoTo 0
		sleep(500)
	End if

	If objFSO.FolderExists(strSnippetPath) Then
		strFilePath = BrowseForFile("Please select the snippet to insert?", strSnippetPath)
		Set objFSO = Nothing
	Else
		MsgBox "Aborting..." & vbCrLf & vbCrLf & "Could not locate the snippets folder:" & strSnippetPath, 16, MODULE_TITLE
		Set objFSO = Nothing
		Exit Sub
	End if

	If (NOT IsNull(strFilePath)) AND Len(strFilePath) Then
			Set objSnippetEditor = newEditor()
			With objSnippetEditor
				.openFile(strFilePath)
				strSnippet = .text()
				.closeFile()
			End With
			objEditor.lineText strSnippet & objEditor.lineText
	Else
		MsgBox "Aborting..." & vbCrLf & vbCrLf & "No snippet chosen.", 16, MODULE_TITLE
	End If

	Set objSnippetEditor = nothing
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

	objDialog.Filter = "Snippet Files|*.txt|All Files|*.*"
    objDialog.FilterIndex = 1
    objDialog.InitialDir = strRootPath
    objDialog.ShowOpen

	BrowseForFile = objDialog.FileName
	If Err.Number > 0 Then
		MsgBox ("Error # " & CStr(Err.Number) & " " & Err.Description)
		BrowseForFile = Null
	End If
End Function

Function AppPath()
	Dim strPath
	strPath = modulePath()
	If Len(strPath) = 0 Then Exit Function
	AppPath = Left(strPath, InStrRev(strPath, "\Script\VBScript\"))
End Function
