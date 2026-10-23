'*******************************************************************************
'	Filename		: text_to_url.vbs

'	Description		: Converts the selected text into an URL

'	Created			: 04 September 2007
'	Created by		: Marco Torskyj
 
'	Note			: This script has only been tested on Windows 2000 and XP

'	You may distribute this script freely, but please keep this header intact.
'*******************************************************************************
Option Explicit

Const MODULE_NAME = "TextToURL"
Const MODULE_VER = "0.1"
Const MODULE_TITLE = "Text To URL"

Sub Init
    addMenuItem MODULE_TITLE, "Convert", MODULE_NAME
End Sub

Sub TextToURL()
	Dim editor
	Dim i, intCreateOrdered
	Dim strInput, strLine, strList, strListType
	Dim arrLines

	Set editor = newEditor()

	With editor
		.assignActiveEditor()
		strInput = .selText()
	End With

	If Len(strInput) > 0 Then

		arrLines = Split(strInput, vbCrLf)

		For i = 0 To UBound(arrLines)
  			strLine = Trim(arrLines(i))
			strList = strList & "<a href=""" & strLine & """>" & strLine &  "</a>" & vbCrLf
	  	Next

		editor.selText strList

	Else
		MsgBox "Please select some text...  ", 0, MODULE_TITLE
	End If

	Set editor = Nothing

End Sub
