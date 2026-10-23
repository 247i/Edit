'*******************************************************************************
'	Filename		: open_linked_file.vbs

'	Description		: Gets the HREF|SRC|ACTION|FILE|VIRTUAL attribute of the tag
'					  that the cursor is resting in and opens the file. If a file
'					  path or a part thereof is selected then this will be treated
'					  as a relative path on the local machine/network
'					  In the case of virtual paths the user is presented with a 
'					  browse folder dialog. Cannot open remote files.

'	Created			: 21 September 2005
'	Created by		: James Swan

'	Note			: Since this script uses Regular Expressions, VBScript version 5 
'					  is needed to run it. This is the version of the VBScript 
'					  scripting engine released with Internet Explorer 5.
'					  This script has only been tested on Windows 2000 and XP

'	You may distribute this script freely, but please keep this header intact.

'	Version History	: 0.2 - fixed bug in GetTagAtCursor() function to allow for inline ASP
'					  0.3 - fixed bug in GetHTMLAttribute() function to strip any querystring or inline ASP from file name
'*******************************************************************************
	Dim objEditor

	Const MODULE_NAME  = "OpenLinkedFile"
	Const MODULE_VER   = "0.3"
	Const DIALOG_TITLE = "Open Linked File"

	Sub Init
	    addMenuItem DIALOG_TITLE, "Utilities", MODULE_NAME, "CTRL+ALT+O"
	End Sub

	Sub OpenLinkedFile()
		Dim objFSO, objLinkEditor
		Dim strFileName, strActivePath, strFilePath, strTag, strFileContent
		Dim BACKSLASH, SLASH
		Dim i, intSelStart, intSelLength
		Dim blnVirtual
		Dim arrPath

		Const ICON_STOP = 16
		BACKSLASH = Chr(92)
		SLASH = Chr(47)
		blnVirtual = False

		Set objEditor = newEditor()

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
				If InStr(UCase(strTag), "INCLUDE VIRTUAL") > 0 Then blnVirtual = True End if
			End if
			'-------------------------------------------------------------------

			If Len(strFileName) Then

				If Left(strFileName, 7) = "mailto:" Then
					MsgBox "Aborting..." & vbCrLf & vbCrLf &_
					"Can't open a ""mailto:"" link." & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
					Exit Sub
	            End if

				If InStr(strFileName, "http://") > 0 Then

					MsgBox "Can't open remote files" & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE

				Else

					' get the file path of the active document
					strActivePath = ExtractFilePath(objEditor.fileName())

					If InStr(strFileName, "../") > 0 Then

						' trim trailing backslash from active document path
	                    strActivePath = Left(strActivePath,Len(strActivePath)-1)

						' convert the path into an array of drive/directory names
						arrPath = Split(strActivePath, BACKSLASH)
						strActivePath = ""

						' go one level up in active document path for each "../" in relative file path
						For i = 0 To UBound(arrPath) - MatchCount("\.\.\/", strFileName)
							strActivePath = strActivePath & arrPath(i) & BACKSLASH
					    Next

						' strip leading "../":s from relative file path
						strFileName = Replace(strFileName, "../", "")

					ElseIf blnVirtual Then

						strActivePath = BrowseForFolder("Cannot currently resolve virtual paths..." &_
						vbCrLf & "Please select the include file folder.")

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

					Set objFSO = CreateObject("Scripting.FileSystemObject")
						If objFSO.FileExists(strFilePath) Then
							Set objLinkEditor = newEditor()
							objLinkEditor.openFile(strFilePath)
							Set objLinkEditor = Nothing
						Else
							MsgBox "Aborting..." & vbCrLf & vbCrLf &_
							"Could not find the file." & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
							Exit Sub
						End if
					Set objFSO = Nothing
				End if
			Else
				MsgBox "Aborting..." & vbCrLf & vbCrLf &_
				"No file name selected." & vbTab & vbCrLf, ICON_STOP, DIALOG_TITLE
				Exit Sub
			End if

		End With

		Set objEditor = Nothing

	End Sub

	Function MatchCount(patrn, str)
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

	Function RegExpTest(patrn, str)
		Dim objRegEx

		Set objRegEx = New RegExp
			With objRegEx
				.Pattern = patrn
				.IgnoreCase = True
				RegExpTest = .Test(str)
			End With
		Set objRegEx = Nothing
	End Function

	Function GetFileExt(strFileName)
		Dim strReturn
		If Len(strFileName) = 0 Then Exit Function			' input string is empty
		If InStr(strFileName, ".") = 0 Then Exit Function	' string contains no file ext
		strReturn = Right(strFileName, Len(strFileName) - InStrRev(strFileName, "."))
		If InStr(strReturn, "?") > 0 Then strReturn = Left(strReturn, InStr(strReturn, "?")-1) End if
		GetFileExt = strReturn
	End Function

	Function BrowseForFolder(strPrompt)
		Const BIF_NONEWFOLDERBUTTON = &H0200 	' Do not include the New Folder button in the browse dialog box.
		Const BSF_DESKTOP = 0 					' Desktop is the root directory. With BIF_returnonlyfsdirs circumvents problem with OK-button
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

	Function GetHTMLAttribute (strTag)
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
						strFileName = Left(strFileName, InStr(strFileName, "?")-1)		' strip any querystring
					End If
					If InStr(strFileName, "<%") > 0 Then
						strFileName = Left(strFileName, InStr(strFileName, "<%")-1)		' strip any inline asp
					End If
				Next
			End With
		Set sourceAttributes = nothing
		Set objRegExp = nothing

		GetHTMLAttribute = strFileName
	End Function

	Function GetTagAtCursor()
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
			If strChar = ">" AND strLastChar <> "%" Then
				intEndPos = i
				Exit For
			End If
			strLastChar = strChar
		Next

		For i = intEndPos To 1 Step -1
			strChar = (Mid(strLine, i, 1))
			If strChar = "<" AND strLastChar <> "%" Then
				intStartPos = i
				Exit For
			End If
			strLastChar = strChar
		Next

		GetTagAtCursor = Mid(strLine, intStartPos, intEndPos - intStartPos + 1)
	End Function

	Function ExtractFilePath(strPath)
    	If Len(strPath) = 0 Then Exit Function				 ' input string is empty
    	strPath = Replace(strPath, Chr(47), Chr(92))		 ' convert backslashes to forward slashes
    	If InStr(1, strPath, Chr(92)) = 0 Then Exit Function ' string contains no forward slashes
    	ExtractFilePath = Left(strPath, InStrRev(strPath, Chr(92)))
    End Function
