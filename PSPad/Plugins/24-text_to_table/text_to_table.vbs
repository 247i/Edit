'*******************************************************************************
'	Filename		: text_to_table.vbs

'	Description		: Converts the selected text into an HTML table with the
'					  border, cellpadding and cellspacing attributes set to zero.
'					  Each line of selected text is converted to a table row and
'					  tabs within the text are treated as column delimiters
'					  within the row. The last column of any row with less than
'					  the maximum number of columns for the table will take up
'					  the difference in column span required. Blank lines are
'					  converted to empty rows. (containing "&nbsp;")

'	Created			: 22 September 2005
'	Created by		: James Swan 

'	Last Updated	: 09 July 2006

'	Note			: The default column delimiter, as well as the row and the 
'					  column indent used when formatting the returned HTML code 
' 					  can be changed to your preference by changing the values 
' 					  of the 3 global variables starting on LINE 47.

' 					  If using the default column delimiter - Tab (Chr(9)) - then
'					  you need to configure PSPad correctly: 
'					  Go to >Settings >Program Settings >Editor Settings and make 
'					  sure that the box Real Tabs is checked.

'					  This script has only been tested on Windows 2000 and XP

'	You may distribute this script freely, but please keep this header intact.

'	Version History	: 0.2 - fixed bug to avoid inserting colspan if less than 2
'							fixed bug to correct maximum colspan if rows contain tabs
'					  0.3 - added possibility to easily personalise the desired 
'							column delimiter, as well as the row and column indent
'							used when formatting the returned HTML code
'*******************************************************************************
Option Explicit

Const MODULE_NAME = "TextToTable"
Const MODULE_VER = "0.3"
Const MODULE_TITLE = "Text To Table"
Dim COLUMN_DELIMITER, COLUMN_INDENT, ROW_INDENT

'**************************** CONFIGURE ****************************************
' the string to use as the column delimiter for each row in the selected text
COLUMN_DELIMITER = Chr(9) 	' Tab

' the string to use as the row indent when formatting the returned HTML code
ROW_INDENT       = Chr(32) 	' Single Blank Space

' the string to use as the column indent when formatting the returned HTML code
COLUMN_INDENT    = Chr(9)	' Tab
'*******************************************************************************

Sub Init
    addMenuItem MODULE_TITLE, "Convert", MODULE_NAME
End Sub

Sub TextToTable()
	Dim editor
	Dim i, j, intRowCells, intMaxCells, intSpanCount
	Dim strInput, strLine, strTable, strCell
	Dim arrLines, arrCells

	Set editor = newEditor()

	With editor
		.assignActiveEditor()
		strInput = .selText()
	End With

	If Len(strInput) > 0 Then

		arrLines = Split(strInput, vbCrLf)
		strTable = "<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf

		'-----------------------------------------------------------------------
		'	establish the maximum number of cells per row
		'-----------------------------------------------------------------------
		intMaxCells = 0
		If InStr(strInput, COLUMN_DELIMITER) > 0 Then
			For i = 0 To UBound(arrLines)
	  			strLine = arrLines(i)
				arrCells = Split(strLine, COLUMN_DELIMITER)
				intRowCells = UBound(arrCells) + 1
				If intRowCells > intMaxCells Then intMaxCells = intRowCells End If
		  	Next
	  	End If
		'-----------------------------------------------------------------------
		'	concatenate table string
		'-----------------------------------------------------------------------
		For i = 0 To UBound(arrLines)
  			strLine = Trim(arrLines(i))
  			If Len(strLine) = 0 Then strLine = "&nbsp;" End if

			arrCells = Split(strLine, COLUMN_DELIMITER)
			intRowCells = UBound(arrCells)

			strTable = strTable & ROW_INDENT & "<tr>" & vbCrLf

			For j = 0 To intRowCells
				strCell = arrCells(j)
				strTable = strTable & COLUMN_INDENT & "<td"

				If j = intRowCells Then
					intSpanCount = intMaxCells - intRowCells
				Else
					intSpanCount = 1
				End If

				If intRowCells = 0 AND intMaxCells > 1 Then
						strTable = strTable & " colspan=""" & intMaxCells & """"
				ElseIf intRowCells+1 < intMaxCells AND intSpanCount > 1 Then
						strTable = strTable & " colspan=""" & intSpanCount & """"
				End if

				strTable = strTable & ">" & strCell & "</td>" & vbCrLf

			Next

			strTable = strTable & ROW_INDENT & "</tr>" & vbCrLf
	  	Next

		strTable = strTable & "</table>"
		editor.selText strTable

	Else
		MsgBox "Please select some text...  ", 0, MODULE_TITLE
	End If

	Set editor = Nothing

End Sub
