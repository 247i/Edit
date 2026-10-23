'*******************************************************************************
'	Filename		: save_backup_file.vbs

'	Created			: 23 January 2006
'	Created by		: James Swan

'	Comments		: Saves a backup copy of the active document with the 
'					  date attached as a prefix to the file name.
'					  eg: 060123_save_backup_file.vbs.

'					  If a backup of the file has already been created for that 
'					  date then a version number is added to the date prefix.
'					  eg: 060123_01_save_backup_file.vbs

'					  In the event of an existing backup file with a version
'					  number the script will increment the version number in the
'					  prefix by one.
'					  eg: 060123_02_save_backup_file.vbs

'	Note			: If an alternative root backup path is specified on line 38
'					  then the full path to the original file will be recreated
'					  in this root folder and the backup file will be saved there.
'					  This script has only been tested on Windows 2000 and XP

'	You may distribute this script freely, but please keep this header intact.

'	Version History	: 0.2 - 7 August 2006 - added ability to specify an alternative 
'							backup folder to that of the original file
'*******************************************************************************
	Option Explicit
	
	Const MODULE_NAME  = "SaveBackUpFile"
	Const MODULE_VER   = "0.2"
	Const MODULE_TITLE = "Save BackUp File"
	
	'**************************** CONFIGURE ************************************
	' an alternative root backup folder path
	Const BACKUP_ROOT  = ""	' eg: "C:\backup\"
	'***************************************************************************
	
	Dim objFSO
	Dim g_strNewFilePath
	
	Sub Init
	    addMenuItem MODULE_TITLE, "Utilities", MODULE_NAME, "SHIFT+CTRL+ALT+B"
	End Sub
	
	Sub SaveBackUpFile()
		Dim objEditor, objSaveEditor
		Dim strFileContents, strFilePath, strFolderPath, strFolderName
		Dim strRoot, strFileName
		Dim i
		Dim arrFolders
	
		Set objEditor = newEditor()
		Set objSaveEditor = newEditor()
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		With objEditor
			.assignActiveEditor()
			strFileContents = .text()		' grab the file contents of the active document
			strFilePath 	= .fileName()	' grab the file path of the active document
		End With
		
		'-----------------------------------------------------------------------
		'	create a file path to a new file with its name prefixed with the
		'	current date - eg: 060123_save_backup_file.vbs
		'-----------------------------------------------------------------------
		strFolderPath	 = GetDirectory(strFilePath)
		strFileName		 = GetFileName(strFilePath)
		
		' in the event of an alternative backup folder to that of the original file
		If Len(BACKUP_ROOT) Then
			
			strFolderPath = BACKUP_ROOT & Replace(strFolderPath, ":", "")
			
			' split path tree into array of folder names
			arrFolders = Split(strFolderPath, "\")
			
			' loop through nested folders and create if necessary
			For i = 1 To UBound(arrFolders)
				strRoot = strRoot & arrFolders(i-1) & "\"
				strFolderName = strRoot & arrFolders(i)
				
				If NOT objFSO.FolderExists(strFolderName) Then
					objFSO.CreateFolder strFolderName
				End if
			Next
			
		End If
		
		g_strNewFilePath = strFolderPath & FormatDate(Date()) & "_" & strFileName
		
		'-----------------------------------------------------------------------
		'	if a backup of this file has already been created today then add a
		'	version number to the date prefix - eg: 060123_01_save_backup_file.vbs
		'-----------------------------------------------------------------------
		
		If objFSO.FileExists(strFolderPath & FormatDate(Date()) & "_01_" & strFileName) Then
			g_strNewFilePath = strFolderPath & FormatDate(Date()) & "_01_" & strFileName
		End if
		
		'-----------------------------------------------------------------------
		'	check if a backup file with this prefix already exists - in the event
		'	of an existing file the function will recurse and increment the version 
		'	number in the prefix by one - sets a global variable
		'-----------------------------------------------------------------------
		Call CheckForExisting(g_strNewFilePath, strFileName)
		
		With objSaveEditor
			.newFile()						' open a new document
			.text(strFileContents)			' copy the contents of the original file
			.saveFileAs(g_strNewFilePath) 	' save the document with the prefixed file name
			.closeFile()					' close the document
		End With
		
		Set objFSO = Nothing	
		Set objEditor = Nothing
		Set objSaveEditor = Nothing
	
	End Sub
	
	Function GetDirectory(strPath)
		If Len(strPath) = 0 Then Exit Function				 ' input string is empty
		strPath = Replace(strPath, Chr(47), Chr(92))		 ' convert backslashes to forward slashes
		If InStr(1, strPath, Chr(92)) = 0 Then Exit Function ' string contains no forward slashes
		GetDirectory = Left(strPath, InStrRev(strPath, Chr(92)))
	End Function
	
	Function GetFileName(strPath)
		If Len(strPath) = 0 Then Exit Function				' input string is empty
		strPath = Replace(strPath, Chr(47), Chr(92))		' convert backslashes to forward slashes
		If InStr(strPath, Chr(92)) = 0 Then Exit Function	' string contains no forward slashes
		If Right(strPath, 1) = Chr(92) Then Exit Function	' path contains no file name
		GetFileName = Right(strPath, Len(strPath) - InStrRev(strPath, Chr(92)))
	End Function
	
	Function FormatDate(datDate)
	
		Dim intDay, intMonth, intYear
		
		intDay = Day(datDate)
		intMonth = Month(datDate)
		intYear = Year(datDate)
		
		If intDay < 10 Then intDay = "0" & intDay End If
		If intMonth < 10 Then intMonth = "0" & intMonth End If
		
		FormatDate = Right(intYear,2) & intMonth & intDay
	
	End Function
	
	Sub CheckForExisting(strFilePath, strOrignalFileName)
		Dim objFile
		Dim strFileName, strVersionPrefix, strDatePrefix
		Dim intVersion
		
		If objFSO.FileExists(strFilePath) Then
			
			' get filename without extension
			strFileName = GetFileName(strFilePath)
			strFileName = Mid(strFileName, 1, InStrRev(strFileName,".")-1)
			
			' get the date prefix and version number
			strDatePrefix = Left(strFileName, 7)
			intVersion = Mid(strFileName,8,2)
			
			If IsNumeric(intVersion) AND Mid(strFileName,10,1) = "_" Then
				intVersion = CInt(intVersion) + 1
				intVersion = "0" & intVersion
				intVersion = Right(intVersion, 2)
				strVersionPrefix = intVersion & "_"
				g_strNewFilePath = Replace(strFilePath, strFileName, strDatePrefix & strVersionPrefix & Mid(strFileName, 11))
			Else
				' rename the file by adding a version number to the date prefix
				Set objFile = objFSO.GetFile(strFilePath)
					objFile.Name = strDatePrefix & "01_" & strOrignalFileName
				Set objFile = Nothing
				g_strNewFilePath = Replace(strFilePath, strFileName, strDatePrefix & "02_" & Mid(strFileName, 8))
			End If
			
			Call CheckForExisting(g_strNewFilePath, strOrignalFileName)
			
		Else
			g_strNewFilePath = strFilePath
		End if
	End Sub
