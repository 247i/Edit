'*******************************************************************************
'	Filename		: text_to_list.vbs

'	Description		: Converts the selected text into an HTML list

'	Created			: 28 September 2005
'	Created by		: James Swan
 
'	Note			: This script has only been tested on Windows 2000 and XP

'	You may distribute this script freely, but please keep this header intact.
'*******************************************************************************
Option Explicit

Const MODULE_NAME = "TextToList"
Const MODULE_VER = "0.1"
Const MODULE_TITLE = "Text To List"

Sub Init
    addMenuItem MODULE_TITLE, "Convert", MODULE_NAME
End Sub

Sub TextToList()
	Dim editor
	Dim i, intCreateOrdered
	Dim strInput, strLine, strList, strListType
	Dim arrLines
	Const vbQuestion = 32
	Const vbYesNo = 4
	Const vbYes = 6 

	Set editor = newEditor()

	With editor
		.assignActiveEditor()
		strInput = .selText()
	End With

	If Len(strInput) > 0 Then
		
		intCreateOrdered = MsgBox ("Do you want to create an ordered list  " & vbCrLf & "(No = create an unordered list)", vbQuestion + vbYesNo, MODULE_TITLE)
		
		If intCreateOrdered = vbYes Then
			strListType = "ol"
		Else
			strListType = "ul"
		End If
		
		arrLines = Split(strInput, vbCrLf)
		strList = "<" & strListType & ">" & vbCrLf

		'-----------------------------------------------------------------------
		'	concatenate table string
		'-----------------------------------------------------------------------
		For i = 0 To UBound(arrLines)
  			strLine = Trim(arrLines(i))
			strList = strList & vbTab & "<li>" & strLine & "</li>" & vbCrLf
	  	Next

		strList = strList & "</" & strListType & ">"
		editor.selText strList

	Else
		MsgBox "Please select some text...  ", 0, MODULE_TITLE
	End If

	Set editor = Nothing

End Sub
