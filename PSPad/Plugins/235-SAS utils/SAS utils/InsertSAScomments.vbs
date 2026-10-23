' *******************************************************************************
' 
' PROGRAM NAME:
' InsertSAScomments.vbs
' 
' DESCRIPTION:
' Insert or remove SAS comments /* */
' Unfortunately, CTRL+/ and CTRL+SHIFT+/ are not available as shortcut keys
' 
' PROGRAMMER:
' John Hendrickx <John2.Hendrickx@gmail.com>
' 
' *******************************************************************************

' VB Script Document
' Based on InsertComments-ENG.js
' by maleval@gawab.com
' option explicit
Const MODULE_NAME  = "InsertSAScomments"
Const MODULE_VER   = "1.0"

Sub Init
  addMenuItem "Insert /* */", "SAS", "insertSASComments","CTRL+SHIFT+C"
  addMenuItem "Remove /* */", "SAS", "removeSASComments","CTRL+ALT+C"
End Sub

Sub insertSASComments()
	dim currFile, s, currline, startline, i
	set currFile = newEditor()
	currFile.assignActiveEditor
	s = currFile.selText()
	currline=currFile.caretY()
	
	if (len(s)=0) then 
		startline=currline
	else	
		currFile.command "ecCut"
		startline=currFile.caretY()
		currFile.command "ecUndo"
	end if
	
	currFile.setCaretPos 1, startline 

	Set delLeadTabs = New RegExp
	With delLeadTabs
		.Global     = True
		.IgnoreCase = True
		.Pattern    = "^\t+(.*)"
	End With

	for i=1 to currline-startline+1
		s = currFile.lineText()
		if (len(trim(s))>0) then
			s = delLeadTabs.Replace(s,"$1")
			s = "/*"+s+"*/"
			currFile.lineText(s)
		end if
		currFile.command "ecDown"
	next
	
	currFile.setCaretPos 1, currline
	currFile.command "ecLineEnd"
End Sub

Sub removeSASComments()
	dim currFile, s, currline, startline, i, regReplace, delLeadTabs, res
	set currFile = newEditor()
	currFile.assignActiveEditor
	s = currFile.selText()
	currline=currFile.caretY()
	
	if (len(s)=0) then 
		startline=currline
	else	
		currFile.command "ecCut"
		startline=currFile.caretY()
		currFile.command "ecUndo"
	end if
	
	Set regReplace = New RegExp
	With regReplace
		.Global     = True
		.IgnoreCase = True
		.Pattern    = "(.*)\/\*(.*)\*\/(.*)"
	End With

	Set delLeadTabs = New RegExp
	With delLeadTabs
		.Global     = True
		.IgnoreCase = True
		.Pattern    = "^\t+(.*)"
	End With

	currFile.setCaretPos 1, startline 
	for i=1 to currline-startline+1
		s = currFile.lineText()
		s = delLeadTabs.Replace(s,"$1")
		res = regReplace.Replace(s, "$1$2$3")
		currFile.lineText(res)
		currFile.command "ecDown"
	next
	
	currFile.setCaretPos 1, currline
	currFile.command "ecLineEnd"
End Sub
