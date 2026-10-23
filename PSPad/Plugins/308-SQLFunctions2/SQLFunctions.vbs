option explicit
const module_name  = "SQLFunctions"       'this name must be unique !!!
const module_ver   = "2.0"            'version
 
const bDEBUG=false
const bTIMER=false
' Folder for created files - MUST be a valid directory'
Const TEMP_DIR = "%TEMP%"

' Change the connection string to match your database
' Put %CN=NAME% w/out CONN_ on first line of file to connect

Const CONN_DBF = "DSN=DBF;"
Const CONN_DBF2 = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=c:\temp;"
Const CONN_IssueXLS = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=C:\temp\Issues.xls;"
Const CONN_IssueXLS2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\temp\issues2.xls;Extended Properties=""Excel 8.0;HDR=YES"";"
Const CONN_Northwind = "Driver={SQL Server};Server=MyServer;Database=Northwind;Uid=myuser;Pwd=mypassword;"
Const CONN_SQLITE = "DRIVER=SQLite3 ODBC Driver;Database=C:\db\my.sqlite;"
Const CONN_LocalSQLServer = "Provider=SQLNCLI11;Server=(localdb)\v11.0;Database=localdb;Integrated Security=SSPI"

Dim CONN_DEFAULT
Dim CONN_TYPE

Const CLEAR_LOG = True 			' Set to False to not clear the log window.  
Const MAX_RECORDS = 300 	' Max records returned for GetData and CreateInserts.  Override with %MAX_RECORDS=20%
Const MAX_FIELD_LENGTH = 300 	' Max field length displayed for GetData.   Override with %MAX_FIELD_LENGTH=50%
Const CMD_TIMEOUT = 30 			' Max seconds for a SQL statement to execute.   Override with %CMD_TIMEOUT=60%
Const RESULTS_TO_FILE = False ' If True, then 'GetData' will write the results in a new document.  Override with %RESULTS_TO_FILE=true|false% 
Const FILE_EXECUTION = False	' Set to True to allow entire files to be executed Override with %ALLOW_FILE_EXECUTION=true|false%

Const REUSE_BROWSER = true		' Reuse the current browser window.  Override with %REUSE_BROWSER%
Const DETERMINE_SELECT_STATEMENT = true 	' If true, then select statements can be executed w/out selecting text.'
Const CREATE_STATEMENTS_SEP = "::"	' Separator between the SQL Statement and the text statement to repeat
Const ENTIRE_FILE_DENIED_MSG = "Entire File Execution Disabled"

CONN_DEFAULT = CONN_Northwind ' Needs to be set to one of the connections above.  Override with %CN=Name% without the CONN_'
								'Override db with %DB=DatabaseName%
CONN_TYPE = "MS"	' If = O then use Oracle settings  Override with CN_TYPE.  Oracle support is experimental


'-------------------------------------'
' Output to browser template'
'-------------------------------------'
CONST HTML_HEAD = "<html><head><title>PSPad Output</title>" 
DIM HTML_STYLE
DIM HTML_HEADER_SCRIPT
DIM HTML_HEADER

HTML_STYLE = "<style type='text/css'>" _
& ".evenRow { background-color:white;font-family:verdana;font-size:smaller;font-weight:normal;} " _
& ".oddRow { background-color:#F3EEEA;font-family:verdana;font-size:smaller;font-weight:normal;} " _
& ".headingRow{ background-color:white;font-family:verdana;font-size:smaller;color:black;font-weight:bold;}" _
& ".hideSpanData { display:none; } " _
& ".showSpanData { display:inline; } " _
& ".divBatch { background-color:blue;color:white;font-family:verdana;font-size:normal;font-weight:bold;cursor:pointer}" _
& "</style>" 

HTML_HEADER_SCRIPT = vbcrlf & "<script language='javascript'>" & vbcrlf _
& "function ShowHideToggle(pSpanID,pMinimizeOnly) { " & vbcrlf _
& " var o = document.getElementById(pSpanID); " & vbcrlf _
& " if (o) { " & vbcrlf _
& " 	if ( (o.className == 'hideSpanData') && (!pMinimizeOnly) )" & vbcrlf _
& "		o.className = 'showSpanData' " & vbcrlf _
& " 	else " & vbcrlf _
& "		o.className = 'hideSpanData'; }" & vbcrlf _
& "}" & vbcrlf _
& "function RemoveItem(pSpanID) { " & vbcrlf _
& "	var oDiv = document.getElementById('divData');" & vbcrlf _
& "	var oChild = document.getElementById(pSpanID);" & vbcrlf _
& "	oDiv.removeChild(oChild);" & vbcrlf _
& "}" & vbcrlf _
& "function AddItem(pDivID, pText) { " & vbcrlf _
& "	var oDiv = document.getElementById('divData');" & vbcrlf _
& " var oChild = document.createElement('div');" & vbcrlf _
& " oChild.id = pDivID;" & vbcrlf _
& " oChild.innerHTML = pText;" & vbcrlf _
& " oDiv.appendChild(oChild);" & vbcrlf _
& " oDiv.appendChild(document.createElement('br'));" & vbcrlf _
& "}" & vbcrlf _
& "</script>"

HTML_HEADER = HTML_HEAD & HTML_STYLE & HTML_HEADER_SCRIPT _
& "</head><body><div id='divData' style='width:100%;'>"

Const HTML_FOOTER = "</div></body></html>"

Const HTML_BROWSER_WIDTH = 1000
Const HTML_BROWSER_HEIGHT = 800
Const HTML_MAX_CHAR_WIDTH = 50
Const HTML_BROWSER_TOOLBAR = 1

Const HTML_BATCH_HEADER = "Statement #"

'--------------------------------------------------
'---------- C# Class template
'-------------------------------------------------'
Dim CSHARP_CLASS_START_TEMPLATE 
Dim CSHARP_PROPERTY_TEMPLATE
Dim CSHARP_CLASS_END_TEMPLATE 

CSHARP_CLASS_START_TEMPLATE = "public class <classname> : BaseModel" & vbcrlf & "{" & vbcrlf
CSHARP_PROPERTY_TEMPLATE = vbtab & "public <fieldtype> <fieldname> { get; set;}" & vbcrlf & vbcrlf 
CSHARP_CLASS_END_TEMPLATE = "}" & vbcrlf & vbcrlf
'------------------------------------------------------------------------------'

'-------------------------------------------------'
' SQL Server queries'
Dim IDENTITY_QUERY
IDENTITY_QUERY = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' " _
	& "AND TABLE_NAME = '<TABLENAME>' AND COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 1"

'--------------------------'

'--------------------------- Main script starts here --------------------------'

dim Document, SQLLog
dim gBrowser
dim gLine0	' used for configuration
Dim gDataCnt
Dim gbNewLogFunctions 
gDataCnt = 0
set gBrowser = nothing
set SQLLog = NewEditor()
gbNewLogFunctions = (SQLLog.AssignLog = 0)

dim test, glTimerStart, glTimerEnd, goHelper
set goHelper = nothing
CONST gV1 = FALSE

' name "Init" is required, its called automatically during initialization to create menu items
sub Init
	addMenuItem "Execute SQL", "SQL Server", "ExecSQL","ALT+X"
	addMenuItem "Get Data", "SQL Server", "GetData", "ALT+G"
	addMenuItem "Get Data In Browser", "SQL Server", "GetDataInBrowser", "ALT+B"
	addMenuItem "Select * from <selected>", "SQL Server", "GetAllData", "CTRL+ALT+G"
	addMenuItem "Select * from <selected> in Browser", "SQL Server", "GetAllDataInBrowser", "CTRL+ALT+B"
	addMenuItem "-", "SQL Server", ""  
	addMenuItem "Table List", "SQL Server", "TableList"
	addMenuItem "Stored Procedure List", "SQL Server", "SPList"
	addMenuItem "View List", "SQL Server", "ViewList"
	addMenuItem "UDF List", "SQL Server", "UDFList"
	addMenuItem "Database List", "SQL Server", "DatabaseList"
	addMenuItem "-", "SQL Server", ""  
	addMenuItem "Get SQL Code", "SQL Server", "GetCode", "CTRL+ALT+C"
	addMenuItem "Get Table Definition","SQL Server","TableDef", ""
	addMenuItem "-", "SQL Server", ""  
	addMenuItem "Create Inserts", "SQL Server", "GetInserts"
	addMenuItem "Create Inserts With Checks", "SQL Server", "GetInsertsAdvanced"
	addMenuItem "Create Statements For Each Record - {field}", "SQL Server", "CreateForEachRecord"
	addMenuItem "Create Statements For Text Record - {%1}", "SQL Server", "CreateForEachText"

	addMenuItem "-", "SQL Server", ""  
	addMenuItem "Generate C# Classes for Tables", "SQL Server", "TableClasses"
   addMenuItem "-", "SQL Server", ""
   addMenuItem "Edit Script", "SQL Server", "EditMe"
	
'	addMenuItem "Test Oracle", "Oracle", "TestOracle", "Alt+D"
end sub

Function TimerStart()
	if bTIMER then
		if goHelper is nothing then
			set goHelper = CreateObject("PSPadHelper.Oracle") 
		end if
	
		glTimerStart = goHelper.Milliseconds()
		TimerStart = glTimerStart
	end If	
End Function

Function TimerEnd()
	if bTIMER then
		if goHelper is nothing then
			set goHelper = CreateObject("PSPadHelper.Oracle") 
		end if
	
		glTimerEnd = goHelper.Milliseconds()
		TimerEnd = glTimerEnd
	End If
End Function

Function WriteTimeElapsed
	dim lDiff
	if bTIMER then
		
		lDiff = glTimerEnd - glTimerStart
		WriteInfo "Time Elapsed: " & lDiff
		WriteTimeElapsed = lDiff
	End If
End Function


Sub TestOracle()
	call AssignDocument()
	if LoadLibrary() then
		dim OraSession, OraDB, OraDy
		dim i, connStr
		dim sDB, sUID, sPWD
		dim sSQL, sResults
		
		sSQL = GetSQLStatement()
		
		If sSQL = "" then 
			Exit Sub
		End If
		dim times(9)
		dim time
		for i = 0 to 9

		call TimerStart
		MaxFieldLength = clng(GetConfigValueDefault("MAX_FIELD_LENGTH",MAX_FIELD_LENGTH))
		MaxRecords = clng(GetConfigValueDefault("MAX_RECORDS", MAX_RECORDS))
		CONN_TYPE = GetConfigValueDefault("CN_TYPE",CONN_TYPE)

		connStr = GetConnectionString()
		sDB = ParseConnectionString(connStr, "Data Source")
		sUID = ParseConnectionString(connStr, "User ID")
		sPWD = ParseConnectionString(connStr, "Password")

		Set OraSession = CreateObject("OracleInProcServer.XOraSession")
		Set OraDB = OraSession.DbOpenDatabase(sDB, sUID & "/" & sPWD, 0)
		
		Set OraDy = OraDB.DbCreateDynaset(sSQL, 0)
		sResults =  RecordsetToString(OraDy)
		OraDy.Close()
		OraDB.Close()

		Call TimerEnd
		time = WriteTimeElapsed()
		times(i) = time
		next
		dim iTot 
		iTot = 0
		for i = 0 to 9
			iTot = iTot + times(i)
		next
		WriteInfo "Total time: " & iTot
		WriteInfo "Avg: " & iTot / 10		
		msgbox "Done!"
		WriteInfo sResults
	End If
End Sub

Function ParseConnectionString(byval vConnStr, byval vKey)
	dim sValue, iStart, iFinish, iKeyLen
	
	sValue = ""
	
	If instr(vConnStr,vKey & "=") > 0 then
		iStart = instr(vConnStr, vKey)
		iKeyLen = len(vKey) + 1
		iFinish =  instr(mid(vConnStr,iStart+iKeyLen),";")

		If iFinish > 0 then
			sValue = mid(vConnStr, iStart+iKeyLen, iFinish-1)
		End If
	
	End If

	ParseConnectionString = sValue

End Function

Sub EditMe()
 newEditor.openFile moduleFileName(module_name)
End Sub


'----------------------------------
' SQL Retrieve Tables Script for PSPad
'	Purpose
'		The purpose of this script is to retrieve a list of Tables
'		and dump it to a file directly from PSPad.  
'----------------------------------
Sub TableList
	call AssignDocument()
	if LoadLibrary() then
		dim txt, sResults, docname
		
		txt = trim("" & Document.selText)
		
		sResults = GetTableList(txt)
		
		if isOracle() then
			docname = TEMP_DIR & "oratbls.sql"
		else
			docname = TEMP_DIR & "tables.sql"
		end if
		
		if sResults <> "" then
			call NewDocument(docname, sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteInfo "SQL Failed 1"
		end if
	End If
End Sub
'EOF TableList

'----------------------------------
' SQL Retrieve User Defined Function Script for PSPad
'	Purpose
'		The purpose of this script is to retrieve a list of UDF
'		and dump it to a file directly from PSPad.  
'----------------------------------
Sub UDFList
	call AssignDocument()
	if LoadLibrary() then
		dim txt, sResults
		txt = trim("" & Document.selText)
		sResults = GetUDFList(txt)
		
		if sResults <> "" then
			call NewDocument(TEMP_DIR & "udf.sql", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 2"
		end if
	End If
End Sub
'EOF UDFList

'----------------------------------
' SQL Retrieve Stored Procedures Script for PSPad
'	Purpose
'		The purpose of this script is to retrieve a list of Stored Procedures
'		and dump it to a file directly from PSPad.  
'----------------------------------
Sub SPList
	call AssignDocument()
	if LoadLibrary() then
		dim txt, sResults
		txt = trim("" & Document.selText)
	
		sResults = GetSPList(txt, false)
		
		if sResults <> "" then
			call NewDocument(TEMP_DIR & "sp.sql", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 3"
		end if
	End If
End Sub
'EOF SPList

'----------------------------------
' SQL Retrieve Views Script for PSPad
'	Purpose
'		The purpose of this script is to retrieve a list of Views
'		and dump it to a file directly from PSPad.  
'----------------------------------
Sub ViewList
	call AssignDocument()
	if LoadLibrary() then
		dim txt, sResults
		txt = trim("" & Document.selText)
	
		sResults = GetViewList(txt)
		if sResults <> "" then
			call NewDocument(TEMP_DIR & "view.sql", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 4"
		end if
	End If
End Sub
'EOF ViewList

'----------------------------------
' SQL Retrieve Databases Script for PSPad
'	Purpose
'		The purpose of this script is to retrieve a list of Databases
'		and dump it to a file directly from PSPad.  
'----------------------------------
Sub DatabaseList

	call AssignDocument()
	if LoadLibrary() then
		dim sResults
		
		sResults = GetDatabaseList()
		
		if sResults <> "" then
			call NewDocument(TEMP_DIR & "databases.sql", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteInfo "SQL Failed 5"
		end if
	End If
End Sub
'EOF DatabaseList

'----------------------------------
' SQL Retrieve Table Definition Script for PSPad
'	Purpose
'		The purpose of this script is to retrieve the definition of one or 
'		more tables and dump it to a file directly from PSPad.  
'----------------------------------
Sub TableDef
	
	call AssignDocument()
	if LoadLibrary() then
		dim txt, sResults

		txt = trim("" & Document.selText)
		
		if txt = "" then
			WriteInfo "Executing entire file."
			txt = Document.Text
		else
			WriteInfo "Executing selected text."
		end if
		
		sResults = GetTableDef(txt) 
		if sResults <> "" then
			call NewDocument(TEMP_DIR & trim(txt) & ".sql", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 6"
		end if
	End If
End Sub
'EOF DatabaseList

'----------------------------------
' Execute SQL Script for PSPad
'	Purpose
'		The purpose of this script is to execute one or more SQL statements
'		against a database.
'----------------------------------
Sub ExecSQL
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL
		
		sSQL = trim("" & Document.selText)
		
		if sSQL = "" then
			if AllowEntireFile() then
				WriteInfo "Executing entire file."
				sSQL = Document.Text
			else
				WriteError ENTIRE_FILE_DENIED_MSG
				Exit Sub
			end if
			
		else
			WriteInfo "Executing selected text."
		end if
		if ExecuteSQL(sSQL) then
			WriteInfo "SQL Executed Successfully"
		else
			WriteInfo "SQL Failed 7"
		end if
	end if
End Sub
'EOF ExecSQL

'----------------------------------
' Retrieve SQL Script for PSPad
'	Purpose
'		The purpose of this script is to either a View, UDF, or Stored Procedure
'		definition and dump it to a file directly from PSPad.  
'----------------------------------
Sub GetCode
'Oracle'
'SELECT line, text
'FROM all_source
'WHERE owner = ?
'  AND name = ?
''  AND type = 'PROCEDURE'
'ORDER BY line
	call AssignDocument()
	if LoadLibrary() then
		dim txt,sSQL, sResults, sFileName
		
		txt = trim("" & Document.selText())
		
		if txt = "" then
			if AllowEntireFile() then
				WriteInfo "Executing entire file."
				txt = Document.Text
			else
				txt = InputBox("Stored Procedure, View, or UDF:","Entire File Execution Disabled","")
				txt = trim("" & txt)

				if txt = "" then
					WriteError "Cancelled by user."
					Exit Sub
				end if
			end if
		else
			WriteInfo "Executing selected text."
		end if
		sResults = GetHelpText(txt,true)
		if sResults > "" then
			if instr(txt,vbcrlf) > 0 then 
				sFileName = ""
			else
				sFileName = TEMP_DIR & txt & ".sql"
			end if
			call NewDocument(sFileName, sResults)
			
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 8"
		end if
	end if
End Sub
'EOF GetCode

'----------------------------------
' Retrieve SQL Data for PSPad
'	Purpose
'		The purpose of this script is to retrieve the results of a 
'		SQL statement and dump it to a file directly from PSPad.  
'----------------------------------
Sub GetData
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sResults
		
		sSQL = GetSQLStatement()
		
		If sSQL = "" then 
			Exit Sub
		End If
		sResults = ExecuteSQLRetRSString(sSQL) 
		if sResults <> "" then
			if ResultsToFile() then
				call NewDocument("", sResults)
			else
				WriteInfo sResults
			end if
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 9"
		end if
	end if
End Sub
'EOF GetData


'----------------------------------
' Retrieve SQL Data for PSPad
'	Purpose
'		The purpose of this script is to retrieve the results of a 
'		SQL statement and dump it to a file directly from PSPad.  
'----------------------------------
Sub GetDataInBrowser
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sResults, sFileName
		
		sSQL = GetSQLStatement()
		
		If sSQL = "" then 
			Exit Sub
		End If
	
		sResults = ExecuteSQLDisplayBrowser(sSQL) 
		if sResults <> "" then
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 11"
		end if
	end if
End Sub
'EOF GetData

'----------------------------------
' Retrieve All SQL Data for PSPad
'	Purpose
'		The purpose of this script is to build a 'select * from <selected> 
'		SQL statement and dump it to a file directly from PSPad.  
'----------------------------------
Sub GetAllData
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sResults, sFileName
		
		sSQL	= trim("" & Document.selText)
		
		if sSQL = "" then

			sSQL = InputBox("Table name:","Table Name","")
			sSQL = trim("" & sSQL)

			if sSQL = "" then
				WriteError "Cancelled by user"
				Exit Sub
			end if

		end if

		WriteInfo "Select * from " & sSQL & "."
	
		sResults = ExecuteSQLRetRSString("select * from " & sSQL) 
		if sResults <> "" then
			if ResultsToFile() then
				call NewDocument("", sResults)
			else
				WriteInfo sResults
			end if
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 12"
		end if
	end if
End Sub
'EOF GetAllData

'----------------------------------
' Retrieve All SQL Data for PSPad and write to Browser
'	Purpose
'		The purpose of this script is to build a 'select * from <selected> 
'		SQL statement and dump it to a browser directly from PSPad.  
'----------------------------------
Sub GetAllDataInBrowser
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sResults, sFileName
		
		sSQL	= trim("" & Document.selText)
		
		if sSQL = "" then

			sSQL = InputBox("Table name:","Table Name","")
			sSQL = trim("" & sSQL)

			if sSQL = "" then
				WriteError "Cancelled by user"
				Exit Sub
			end if

		end if

		WriteInfo "Select * from " & sSQL & "."

		sResults = ExecuteSQLDisplayBrowser("select * from " & sSQL) 
		if sResults <> "" then
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 13"
		end if

	end if
End Sub
'EOF GetAllDataInBrowser

'----------------------------------
' Create Insert Statements for PSPad
'	Purpose
'		The purpose of this script is to retrieve the results of a 
'		SQL statement and create an Insert Statement for each record.  
'----------------------------------
Sub GetInserts
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sTableName, sResults, sFileName
		
		sSQL = GetSQLStatement()
		
		If sSQL = "" then 
			Exit Sub
		End If

		sSQL = replace(sSQL,vbcrlf, " ")
		sSQL = replace(sSQL,vbtab, " " )
			
		IF instr(ucase(sSQL),"SELECT ") = 0 then
			sSQL = "SELECT * FROM " & sSQL
		End IF

		sTableName = GetSQLSelectTable(sSQL)
		
		sTableName = InputBox("Insert table name:","Table Name",sTableName)
		if trim("" & sTableName) = "" then
			msgbox "Table name is required",0, "Create Inserts"
			exit sub
		end if
		
		sTableName = trim("" & sTableName)
		
		WriteInfo "Executing selected text."
	
		sResults = ExecuteSQLRetInserts(sSQL,sTableName) 
	
		if sResults <> "" then
			call NewDocument("", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 14"
		end if
	end if
End Sub
'EOF GetInserts


'----------------------------------
' Create Insert Statements for PSPad
'	Purpose
'		The purpose of this script is to retrieve the results of a 
'		SQL statement and create an Insert Statement for each record.  
'----------------------------------
Sub GetInsertsAdvanced
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sTableName, sResults, sFileName
		dim sKeyFields, bCreateUpdates
		dim sIdentityField
		
		sSQL = GetSQLStatement()
		
		If sSQL = "" then 
			Exit Sub
		End If
		
		sTableName = GetSQLSelectTable(sSQL)
		
		sTableName = InputBox("Insert table name:","Table Name", sTableName)
		if trim("" & sTableName) = "" then
			msgbox "Table name is required", 0, "Create Inserts"
			exit sub
		end if
		sTableName = trim("" & sTableName)
		
		sKeyFields = InputBox("Key Field(s), separated by commas:","Create Inserts")
		if trim("" & sKeyFields) = "" then
			msgbox "Key Field is required"
			exit sub
		end if
		sKeyFields = trim(sKeyFields)
		
		sIdentityField = ExecuteSQLRetCol0(REPLACE(IDENTITY_QUERY, "<TABLENAME>", sTableName),false)

		sIdentityField = trim("" & InputBox("Enter identity field.  Leave blank to perform identity insert", "Identity Field", sIdentityField))

		bCreateUpdates = (msgbox("Create Update Statement if exists?",4, "Create Inserts") = 6) 'vbYes
				
		WriteInfo "Executing selected text."
	
		sResults = ExecuteSQLRetInsertsCheck(sSQL,sTableName, sKeyFields, bCreateUpdates, sIdentityField) 
	
		if sResults <> "" then
			call NewDocument("", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 16"
		end if
	end if
End Sub
'EOF GetInsertsAdvanced


'----------------------------------
' Create C# Class from Table Name for PSPad
'	Purpose
'		The purpose of this script is to generate a C# class for the selected table or view.
'----------------------------------
Sub TableClasses
	call AssignDocument()
	if LoadLibrary() then
		dim txt, sResults, sFileName
		
		txt = trim("" & Document.selText())
		
		if txt = "" then
			txt = InputBox("Table or View:","Enter a table or view name","")
			txt = trim("" & txt)

			if txt = "" then
				WriteError "Cancelled by user."
				Exit Sub
			end if
		else
			WriteInfo "Executing selected text."
		end if
		txt = trim(txt)
		
		sResults = GetTableClasses(txt)

		if sResults > "" then
			if instr(txt,vbcrlf) > 0 then 
				sFileName = ""
			else
				sFileName = TEMP_DIR & txt & ".cs"
			end if
			call NewDocument(sFileName, sResults)
			
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 15"
		end if
		
	end if
End Sub
'EOF TableClasses


'----------------------------------
' Create Statements over a recordset for PSPad
'	Purpose
'		The purpose of this script is to retrieve the results of a 
'		SQL statement and Create the entered Statement for each record.
'		Replaces are performed based upon field name.  
'		The 
'----------------------------------
Sub CreateForEachRecord
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sStatement, sResults, sFileName
		dim sKeyField, bCreateUpdates, bEscapeQuotes
		
		sSQL	= trim("" & Document.selText)
		
		if sSQL = "" then
			msgbox "Usage: Select the SQL Statement to Execute and the string to repeat, separated by " & CREATE_STATEMENTS_SEP & _
			vbcrlf & "To perform field substitution surround the fields with curly brackets {field}.", 0, "Create Statements For Each Record"
			Exit Sub
		end if
		
		dim asString
		
		asString = split(sSQL,CREATE_STATEMENTS_SEP)
		if UBound(asString) = 0 then
			sStatement = InputBox("Enter Statement to repeat.  Use curly braces {} to surround field names:","Enter Statement")
			if trim("" & sStatement) = "" then
				msgbox "Statement is required"
				exit sub
			end if
			sStatement = trim("" & sStatement)
		else
			sStatement = trim("" & asString(1))
			sSQL = trim("" & asString(0))
		end if
		
		bEscapeQuotes = (msgbox("Wrap substitutions with single quotes (if generating SQL Statements, answer Yes)?",4, "Create Statements") = 6) 'vbYes
				
		WriteInfo "Executing selected text."
	
		sResults = ExecuteSQLRetStatements(sSQL,sStatement, bEscapeQuotes) 
	
		if sResults <> "" then
			call NewDocument("", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 17"
		end if
	end if
End Sub
'EOF CreateForEachRecord

'----------------------------------
' Create Statements from selected text for PSPad
'	Purpose
'		The purpose of this script is to take selected text 
'		and entered separators for each field and record and Create the entered Statement for each record.
'		{%1}, {%2}, etc is used as the substitution.  
'		 
'----------------------------------
Sub CreateForEachText
	call AssignDocument()
	if LoadLibrary() then
		dim sSQL, sStatement, sResults, sFileName
		dim sFldSep, sRecSep
		
		sSQL	= trim("" & Document.selText)
		
		if sSQL = "" then
			msgbox "Usage: Select the text and the string to repeat, separated by " & CREATE_STATEMENTS_SEP & _
			vbcrlf & "To perform substitutions use {%1}, {%2}, etc.", 0, "Create Statements For Text Record"
			Exit Sub
		end if
		
		dim asString, bEscapeQuotes
		
		asString = split(sSQL,CREATE_STATEMENTS_SEP)
		if UBound(asString) = 0 then
			sStatement = InputBox("Enter Statement to repeat.  Use {%1}, {%2}, etc for substitutions:","Enter Statement")
			if trim("" & sStatement) = "" then
				msgbox "Statement is required"
				exit sub
			end if
			sStatement = trim("" & sStatement)
		else
			sStatement = trim("" & asString(1))
			sSQL = trim("" & asString(0))
		end if
		
		sFldSep = InputBox("Enter field separator, leave blank if only one field per record.  Use \n and \t for CRLF and tab","Field Separator")
		if sFldSep = "" then
			sFldSep = " "
		ElseIF sFldSep = "\n" then
			sFldSep = vbCrLf
		ElseIF sFldSep = "\t" then
			sFldSep = vbTab
		End IF

		sRecSep = InputBox("Enter record separator.  Use \n and \t for CRLF and tab", "Record Separator")
		if sRecSep = "" then
			msgbox "Record separator is required."
			Exit Sub
		End If
		
		IF sRecSep = "\n" then
			sRecSep = vbCrLf
		ElseIF sRecSep = "\t" then
			sRecSep = vbTab
		End IF

		bEscapeQuotes = (msgbox("Wrap substitutions with single quotes (if generating SQL Statements, answer Yes)?",4, "Create Statements") = 6) 'vbYes

				
		WriteInfo "Executing selected text."
		
		dim sStatementPost, st, iCnt, sField
		dim asRecords, sRecord, asFields, iFld
		dim iMax
		
		asRecords = split(asString(0),sRecSep)
		
		set st = new FastString
	
		st() = vbcrlf
	
		iMax = UBound(asRecords)
		if trim(asRecords(iMax)) = "" then
			iMax = iMax - 1
		end if
	
		FOR iCnt = 0 to iMax
			sRecord = trim(asRecords(iCnt))
			sRecord = replace(sRecord,vbcrlf,"")
			asFields = split(sRecord, sFldSep)
			
			sStatementPost = sStatement
			FOR iFld = UBound(asFields) to 0 Step -1
				sField = trim(asFields(iFld))
				sField = replace(sField,vbcrlf,"")
				if bEscapeQuotes then
					sField = replace(sField,"'","''")
				end if
				
				sStatementPost = replace(sStatementPost,"{%" & iFld+1 & "}",sField)
			Next
			st() = sStatementPost
			st() = vbcrlf
		Next
		WriteInfo iCnt & " Record(s) Returned "
		sResults = st()
	
		if sResults <> "" then
			call NewDocument("", sResults)
			WriteInfo "SQL Executed Successfully"
		else
			WriteError "SQL Failed 18"
		end if
	end if
End Sub
'EOF CreateForEachText


'------------ end of main script --------------------
Function ClearLog()
	dim bValue
	
	bValue = GetConfigValueDefault("CLEAR_LOG", CLEAR_LOG)

	err.clear
	'xxOn Error resume next
	ClearLog = cBool(bValue)
	if err.number <> 0 then
		ClearLog = CLEAR_LOG
		err.clear
	end if

End Function


Function ResultsToFile()
	dim bValue
	bValue = GetConfigValueDefault("RESULTS_TO_FILE", RESULTS_TO_FILE)

	err.clear
	'xxOn Error resume next
	ResultsToFile = cBool(bValue)
	if err.number <> 0 then
		ResultsToFile = RESULTS_TO_FILE
		err.clear
	end if

End Function

Function AllowEntireFile() 
	dim bValue
	bValue = GetConfigValueDefault("ALLOW_FILE_EXECUTION", FILE_EXECUTION)

	err.clear
	'xxOn Error resume next
	AllowEntireFile = cBool(bValue)
	if err.number <> 0 then
		AllowEntireFile = false
		err.clear
	end if
	
End Function



Function AssignDocument()
	logSetTypeMemo
	if CLEAR_LOG then 
		Call ClearLog
	end if
	gLine0 = ""
	err.clear
	set Document = newEditor()
	'xxOn Error resume next
	call Document.AssignActiveEditor()
	if err.number <> 0 then
		WriteInfo "Opening new document"
		call Document.newFile()
		err.clear
	end if
	AssignDocument = true
End Function


Function GetSQLSelectTable(byval vSQL)
	dim i

	vSQL = replace(vSQL,vbcrlf, " ")
	vSQL = replace(vSQL,vbtab, " ")
	i = Instr(ucase(vSQL)," FROM ")
	if i > 0 then
		vSQL = trim(mid(vSQL, i + 6))
		i = instr(vSQL," ")
		if i > 0 then
			GetSQLSelectTable = left(vSQL,i - 1)
		else
			GetSQLSelectTable = vSQL
		end if
		Exit Function
	end if
	GetSQLSelectTable = ""
End Function

Function LoadLibrary()
' Code combined to one file
	LoadLibrary = true
End Function

'-----------------------------------------------------------------------
' library code
'------------------------------------------------------------------------

Const SQL_LIBRARY_LOADED = true

dim MaxFieldLength
Dim MaxRecords



Function isOracle()
	CONN_TYPE = GetConfigValueDefault("CN_TYPE",CONN_TYPE)
	isOracle =  left("" & CONN_TYPE,1) = "O" 
End Function

Function isOracleNew()
	CONN_TYPE = GetConfigValueDefault("CN_TYPE",CONN_TYPE)
	isOracleNew =  left("" & CONN_TYPE,2) = "O2" 

End Function


'----------------------------------
' Function GetConnection
'	This function determines which connection string 
' should be used, creates an ADO connection
' and opens it.  If a database was specified on the 
' first line with %DB=<database>%, then the specified
' database is selected.
' 	Parameters 
'		none
'	Returns
'		An ADO Connection object
'----------------------------------
Function GetConnection()
	dim sConn, sDB, oConn
		
	sConn = GetConnectionString()
	MaxFieldLength = clng(GetConfigValueDefault("MAX_FIELD_LENGTH",MAX_FIELD_LENGTH))
	MaxRecords = clng(GetConfigValueDefault("MAX_RECORDS", MAX_RECORDS))
	CONN_TYPE = GetConfigValueDefault("CN_TYPE",CONN_TYPE)

	On Error Resume Next
	If sConn > "" then 
		set oConn = createobject("adodb.Connection")
		oConn.ConnectionString = sConn
		oConn.Open
		if Err.Number <> 0 then
			WriteError "Unable to open connection: " & Err.Description
			set GetConnection = nothing
			Exit Function
		end if
		
		sDB = GetDatabase()

		if Err.Number <> 0 then
			WriteError Err.Number & ":" & Err.Description
			err.clear
			Set GetConnection = nothing
			Exit Function
		end if

		if sDB > "" then
			oConn.Execute "use " & sDB
			WriteInfo "Using Database " & sDB
			if Err.Number <> 0 then
				WriteError Err.Number & ":" & Err.Description
				err.clear
				Set GetConnection = nothing
				Exit Function
			end if
		end if
		Set GetConnection = oConn
	else
		Set GetConnection = nothing
	end if
End Function
' EOF GetConnection

'----------------------------------
' Function GetConnectionString
'	Calls GetConfigValue to see if the default connection
'	was changed by specifying %CN=<connection>% 
'	The connection value should match a CONN_CONNECTION, 
'	without the CONN_.
' 	Parameters 
'		none
'	Returns
'		The connection specified as a string, otherwise the CONN_DEFAULT value
'----------------------------------
Function GetConnectionString()
	dim sTmp, sConn
	sConn = ""
	On Error resume next

	sTmp = GetConfigValue("CN")
	if sTmp > "" then
		Execute "sConn = trim("""" & CONN_" & sTmp & ")"
		if err.number = 0 and sConn > "" then
			WriteInfo "Using Connection " & sTmp
		else
			WriteError "Invalid connection string configuration."
			WriteError Err.Number & ":" & Err.Description
			Err.Clear
			Exit Function
		end if
	end if
	
	if sConn = "" then
		sConn = CONN_DEFAULT
		WriteInfo "Using Default Connection"
	end if

	err.clear
	GetConnectionString = sConn
End Function
'EOF GetConnectionString'

'----------------------------------
' Function GetDatabase
'	Calls GetConfigValue to see if the database
'	was changed by specifying %DB=<database>% 
' 	Parameters 
'		none
'	Returns
'		The datatbase specified as a string, otherwise an empty string
'----------------------------------
Function GetDatabase()
	GetDatabase = GetConfigValue("DB")
End Function
'EOF GetDatabase'

'----------------------------------
' Function GetConfigValue
'	Checks the first line of the active file for 
'	%vName=<value>% and returns <value>'
' 	Parameters
'		vName - the name to search 
'	Returns
'		The value specified
'----------------------------------
Function GetConfigValue(byval vName)
	dim sLine0, iLoc
	dim x,y, bReadLine
	
	On Error resume next
	
	GetConfigValue = 	getVarValue("%" & vName & "%")

	bReadLine = false
	if gLine0 = "" then
	
		x = Document.caretX
		y = Document.caretY
		
		call Document.setCaretPos(0,0)
		bReadLine = true
	
		gLine0 = Document.LineText
	end if	

	sLine0 = gLine0
	
	iLoc = instr(sLine0,"%" & vName & "=")
	
	if iLoc > 0 then
		sLine0 = mid(sLine0,iLoc + len(vName) + 2)
		iLoc = instr(sLine0,"%")
		if iLoc > 0 then
			 GetConfigValue  = trim(left(sLine0,iLoc-1))
		end if
	end if
	
	if bReadLine then
		call Document.setCaretPos(x,y)
	end if
	
End Function
'EOF GetConfigValue'

'----------------------------------
' Function GetConfigValueDefault
'	Checks the first line of the active file for 
'	%vName=<value>% and returns <value>'
' 	Parameters
'		vName - the name to search 
'	Returns
'		The value specified or vDefault if not found
'----------------------------------
Function GetConfigValueDefault(byval vName, byval vDefault)
	dim sLine0, iLoc
	dim x,y, bReadLine
	'on error resume next
	err.clear
	dim sTmp
	
	sTmp = getVarValue("%" & vName & "%")
	
	if sTmp = "" then
		sTmp = vDefault
	end if
	
	GetConfigValueDefault = sTmp

	bReadLine = false
	if gLine0 = "" then
	
		x = Document.caretX
		y = Document.caretY
		
		call Document.setCaretPos(0,0)
		bReadLine = true
		
		gLine0 = Document.LineText
	end if
	
	sLine0 = gLine0
	iLoc = instr(sLine0,"%" & vName & "=")
	
	if iLoc > 0 then
		sLine0 = mid(sLine0,iLoc + len(vName) + 2)
		iLoc = instr(sLine0,"%")
		if iLoc > 0 then
			 GetConfigValueDefault  = trim(left(sLine0,iLoc-1))
		end if
	end if
	if bReadLine then
		call Document.setCaretPos(x,y)
	end if
End Function
'EOF GetConfigValue'

'----------------------------------
' Function GetTableList
'	Retrieves the list of tables on the current server
'	
' 	Parameters
'		vTxt - optional filter text
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetTableList(byval vTxt)

	if not isOracle() then
		GetTableList = GetSysObjects("U", vTxt, false)
	else
		GetTableList = ExecuteSQLRetCol0("SELECT table_name from user_tables order by table_name", true) 	
	end if	
End Function
'EOF GetTableList'

'----------------------------------
' Function GetViewList
'	Retrieves the list of views on the current server
'	
' 	Parameters
'		vTxt - optional filter text
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetViewList(byval vTxt)
	if not isOracle() then
		GetViewList = GetSysObjects("V", vTxt, false)
	else
		GetViewList = ExecuteSQLRetCol0("SELECT view_name from user_views", true) 	
	end if
End Function
'EOF GetViewList'

'----------------------------------
' Function GetSPList
'	Retrieves the list of stored procedures on the current server
'	
' 	Parameters
'		vTxt - optional filter text
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetSPList(byval vTxt, byval vBeginsWith)
'Oracle:'
'SELECT owner, object_name
'FROM all_objects 
'WHERE object_type = 'PROCEDURE'
if not isOracle() then
	GetSPList = GetSysObjects("P", vTxt, vBeginsWith)
else
	GetSPList = ExecuteSQLRetCol0("select object_name || '.' || Procedure_Name AS Object_Procedure FROM user_procedures WHERE object_type in ('PACKAGE', 'PROCEDURE') ORDER BY Object_type, object_name, procedure_name", true)
end if

End Function
'EOF GetSPList'

'----------------------------------
' Function GetUDFList
'	Retrieves the list of user defined functions on the current server
'	
' 	Parameters
'		vTxt - optional filter text
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetUDFList(byval vTxt)
' Oracle:'
' select object_name
' from   all_objects
' where  object_type = 'FUNCTION'

if not isOracle() then
	GetUDFList = GetSysObjects("FN:IF:TF", vTxt, false)
else
	GetUDFList = ExecuteSQLRetCol0("select object_name FROM user_procedures WHERE object_type = 'FUNCTION' ", true)
End If
End Function
'EOF GetUDFList'

'----------------------------------
' Function GetSysObjects
'	Creates a sql statement based upon the given sysobject types
'	and returns the result as a string.
'	Multiple items can be passed, separated by a colon
'	
' 	Parameters
'		vTypes - one or more types, separated by a colon
'		vFilter - If entered, will filter the results to include
'		only entries beginning with the vFilter value
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetSysObjects(byval vTypes, byval vFilter, byval vFilterBeginsWith)
	dim sSelect, sSQL, sWhere
	dim sConn, i, iUBound
	dim asTypes
	sSelect = "select name from sysobjects " 
	
	if instr(vTypes,":") > 0 then
		asTypes = split(vTypes,":")
		iUbound =  ubound(asTypes)
		sWhere = " where xtype in ("

		for i = 0 to iUBound
			sWhere = sWhere & "'" & asTypes(i) & "',"
		next
		sWhere = left(sWhere,len(sWhere)-1) & ") "
	else
		sWhere = "where xtype = '" & vTypes & "'"
	end if 
	
	if vFilter > "" then
		if vFilterBeginsWith then
			sWhere = sWhere & " and name like '" & vFilter & "%' "
		Else
			sWhere = sWhere & " and name like '%" & vFilter & "%' "
		End If
	end if
	
	sSQL = sSelect & sWhere & " Order by name"
	GetSysObjects = ExecuteSQLRetCol0(sSQL, true) 
End Function
'EOF GetSysObjects'

'----------------------------------
' Function GetDatabaseList
'	Retrieves the list of databases on the current server
'	
' 	Parameters
'		None
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetDatabaseList()
	dim sSelect, sSQL, sWhere
	dim sConn, i, iUBound
	dim asTypes
	sSQL = "use master; select name from sysdatabases order by name"
	
	if isOracle() then
		WriteError "Not Supported"
		Exit Function
	End if	

	GetDatabaseList = ExecuteSQLRetCol0(sSQL, true) 
End Function
'EOF GetDatabaseList'

'----------------------------------
' Function GetTableDef
'	Retrieves the list of field definitions for a table
'	
' 	Parameters
'		None
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetTableDef(byval vTable)
	dim sSelect, sSQL, sWhere
	dim sConn, i, iUBound
	dim asTypes
	'Sysobjects.name AS TABLE_NAME, " & _

	
	if not isOracle()  then
		WriteDebug "In GetTableDef CONN_TYPE =" & CONN_TYPE & "." 
		sSQL = "SELECT " & _ 
	    "syscolumns.name AS COLUMN_NAME, " & _
	    "systypes.name AS DATA_TYPE, syscolumns.length" & _
	    " as CHARACTER_MAXIMUM_LENGTH, " & _
	    "syscomments.text as COLUMN_DEFAULT, " & _
	    "syscolumns.isnullable as IS_NULLABLE " & _
	    "FROM syscolumns INNER JOIN systypes " & _
	    "ON syscolumns.xusertype = systypes.xusertype " & _
	    "JOIN sysobjects ON syscolumns.id = sysobjects.id " & _
	    "LEFT OUTER JOIN syscomments ON syscolumns.cdefault = syscomments.id " & _
	    "WHERE Sysobjects.name = '" & vTable & "' " & _
		 "ORDER BY syscolumns.colid"
	Else
		sSQL = "select dbms_metadata.get_ddl('TABLE','" & ucase(vTable) & "') from dual"
	End If

	GetTableDef = ExecuteSQLRetRSString(sSQL) 
	WriteCode sSQL
End Function
'EOF GetTableDef'


'----------------------------------
' Function GetTableClasses
'	Parses a list of tables, and generates .net Classes
'	
' 	Parameters
'		vTables - list of tables separated by cr/lf
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetTableClasses(byval vTables)
	dim asData, sTmp, stData, iUBound, i
	dim oConn, oCmd
	
	asData = split(vTables,vbCrLf)
	if not isArray(asData) then
		WriteError "Invalid text selection"
		Exit Function
	end if

	set oConn = nothing
	set oCmd  = nothing
	if not CreateConnCmdObjects(oConn, oCmd) then
		Exit Function
	End If

	set stData = new FastString
	iUBound = ubound(asData)
	On Error resume next
	for i = 0 to iUBound
		sTmp = trim(asData(i))
		WriteInfo "Processing table " & sTmp
		stData() = ExecTableToClass(oConn, oCmd, sTmp)
	next

	sTmp = stData()
	if sTmp = "" then
		WriteError "SQL Failed 19"
	end if
	GetTableClasses = sTmp
	set oCmd = nothing
	oConn.close
	set oConn = nothing


End Function
'EOF GetTableClasses'

'----------------------------------
' Function GetHelpText
'	Parses a list of objects, and retrieves the definitions, with drop statements
'	
' 	Parameters
'		vObjects - list of objects separated by cr/lf
'		vCreateDrop - T/F for creating a drop statement
'	Returns
'		The list as a CRLF string
'----------------------------------
Function GetHelpText(byval vObjects, byval vCreateDrop)
	dim asData
	dim stDrops, stData, i, iUBound, sTmp, sTmp2
	dim oConn, oCmd
	
	asData = split(vObjects,vbCrLf)
	if not isArray(asData) then
		WriteError "Invalid text selection"
		Exit Function
	end if

	set oConn = nothing
	set oCmd  = nothing
	if not CreateConnCmdObjects(oConn, oCmd) then
		Exit Function
	End If
	
	if vCreateDrop then
		set stDrops = new FastString
	end if

	set stData = new FastString
	iUBound = ubound(asData)
	On Error resume next
	for i = 0 to iUBound
		sTmp = trim(asData(i))
		if sTmp > "" and left(sTmp,1) <> "-"  then
			stData() = ExecuteSQLRetCol0_2(oConn, oCmd, "sp_helptext " & sTmp, false)
			if err.number = 0 then 
				stData() = vbcrlf & "GO" & vbcrlf & vbcrlf & vbcrlf

				if vCreateDrop then	' determine type and create drop statement
					sTmp2 = GetDropStatement(oConn, oCmd, sTmp)
					if Err.Number = 0 then
						if sTmp2 > "" then
							stDrops() = sTmp2 & vbcrlf & "GO" & vbcrlf
						end if
					else
						WriteError "GetDropStatement failed for :" & sTmp & " Err: " & Err.Description
						Err.Clear
					end if
				end if
			else
				WriteError "sp_helptext failed for :" & sTmp & " Err: " & Err.Description
				Err.Clear
			end if
			
		elseif sTmp > "" then
			if left(sTmp,2) = "-%" then
				stData() = mid(sTmp,3) & vbcrlf
			else
				stData() = sTmp & vbcrlf
			end if
		end if
	next
	if vCreateDrop then 
		sTmp = stDrops() & vbcrlf & vbcrlf & stData()
	else
		sTmp = stData()
	end if
	
	if sTmp = "" then
		WriteError "SQL Failed 20"
	end if
	GetHelpText = sTmp
	set oCmd = nothing
	oConn.close
	set oConn = nothing
End Function
'EOF GetHelpText

Function CreateConnCmdObjects(byref rConn, byref rCmd)
	On Error resume next
	set rConn = GetConnection()
	if rConn is Nothing then
	CreateConnCmdObjects = false
		Exit Function
	end if
	set rCmd = createobject("adodb.Command")
	If Err.Number <> 0 then
		WriteError "Unable to create ADO objects:" & Err.Description
		rConn.close
		set rConn = nothing
		CreateConnCmdObjects = false
		Exit Function
	End If
	
	set rCmd.ActiveConnection = rConn
	rCmd.CommandType = 1
	rCmd.CommandTimeout = cint(GetConfigValueDefault("CMD_TIMEOUT",CMD_TIMEOUT))
	
	CreateConnCmdObjects = true

End Function
'EOF CreateConnCmdObjects


Function GetDropStatement(byval vConn, byval vCmd, byval vObject)
	dim sType, sStmt
	sType = GetObjectType(vConn, vCmd, vObject)
	select case sType

	case "V" :
		sStmt =  "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" _
		& vObject & "]') and OBJECTPROPERTY(id, N'IsView') = 1) " _
		& "drop view [dbo].[" & vObject & "]"
		
	case "P" :
		sStmt =  "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" _
		& vObject & "]') and OBJECTPROPERTY(id, N'IsProcedure') = 1) " _
		& "drop procedure [dbo].[" & vObject & "]"

	case "FN", "IF", "TF" :
		sStmt =  "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" _
		& vObject & "]') and xtype in (N'FN', N'IF', N'TF')) " _
		& "drop function [dbo].[" & vObject & "] "
	end select
	GetDropStatement = sStmt
	
End Function
'EOF GetDropStatement

Function GetObjectType(byval vConn, byval vCmd, byval vObject)
	dim sSQL
	sSQL = "select xtype from dbo.sysobjects where id = object_id(N'[dbo].[" & vObject & "]')"
	sSQL = ExecuteSQLRetCol0_2(vConn, vCmd, sSQL, false)
	GetObjectType = trim(sSQL)
End Function
'EOF GetObjectType


'----------------------------------
' GetSQLStatement for PSPad
'	Purpose
'		Retrieves the selected SQL statement 
'----------------------------------
Function GetSQLStatement()
		dim sSQL
		sSQL	= trim("" & Document.selText)
		
		if sSQL = "" then
			if AllowEntireFile() then
				WriteInfo "Executing entire file."
				sSQL = Document.Text
				
			ElseIf DETERMINE_SELECT_STATEMENT then
				sSQL = DetermineSQLStatement()
				If sSQL = "" then
					sSQL = InputBox("SQL Statement:","Entire File Execution Disabled","")
					sSQL = trim("" & sSQL)
					If sSQL = "" then
						WriteError "Cancelled by user"
						Exit Function
					End if
				End If
					
			Else
				sSQL = InputBox("SQL Statement:","Entire File Execution Disabled","")
				sSQL = trim("" & sSQL)
				if sSQL = "" then
					WriteError "Cancelled by user"
					Exit Function
				end if
			end if
		else
			WriteInfo "Executing selected text."
		end if

		GetSQLStatement = sSQL
End Function

'----------------------------------
' DetermineSQLStatement for PSPad
'	Purpose
'		Retrieves the SQL statement under the cursor 
'----------------------------------
Function DetermineSQLStatement()
	Dim x, y, yStart, sText, sLine, sTrim, bDone
	Dim iCount 

	x = Document.caretX
	yStart = Document.caretY
	
	y = yStart
	sLine = Document.lineText()
	sTrim = trim(sLine)
	
	If sTrim = "" then
		Exit Function
	End If

	if instr(ucase(sTrim),"SELECT ") = 1 then
		DetermineSQLStatement = sLine
		Exit Function
	end if
	
	sText = sLine
	bDone = false
	iCount = 0
	While not bDone
		y = y - 1
		call Document.setCaretPos(x, y)
		
		sLine = Document.lineText()
		sText = sLine & vbcrlf & sText

		If instr(ucase(sLine),"SELECT ") = 1 then
			bDone = true
		End If
		
		if y = 0 then
			bDone = true
		end if
	Wend
	
	call Document.setCaretPos(x, yStart)
	DetermineSQLStatement = sText

End Function



Function ExecuteSQLRetCol0_2(byval vConn, byval vCmd, byval vSQL, byval vCRLF)
	dim oRS, st, cr

	set oRS = ExecuteSQLRetRS(vConn, vCmd, vSQL)
	if oRS is nothing then
		ExecuteSQLRetCol0_2 = ""
		Exit Function
	end if

	set st = new FastString
	if vCRLF then
		cr = vbcrlf
	else
		cr = ""	
	end if

	while not oRS.EOF
		st() = oRS(0) & cr
		oRS.MoveNext
	wend
	oRS.close
	set oRS = nothing
	
	ExecuteSQLRetCol0_2 = st()
End Function
'EOF ExecuteSQLRetCol0_2

'----------------------------------
' Function ExecuteSQL
'	Executes multiple SQL statements against the server/database specified
'	If a 'hard' error occurs, the script is stopped. 
'	If a 'soft' error occurs, it is displayed in the console window and the script continues.
'	
' 	Parameters
'		vText - SQL Statements to execute
'	Returns
'		True if no hard errors were encountered
'----------------------------------
Function ExecuteSQL(byval vSQL)
	dim asSQL, i, iUBound, oErr
	dim oConn, oCmd
	const adExecuteNoRecords = 128
	
	On Error resume next
	asSQL = ParseText(vSQL)
	if isArray(asSQL) then
		iUBound = ubound(asSQL)
		
		set oConn = GetConnection()
		if oConn is Nothing then
			Exit Function
		end if
		set oCmd = createobject("adodb.Command")
		If Err.Number <> 0 then
			WriteError "Unable to create ADO objects:" & Err.Description
			oConn.close
			set oConn = nothing
			Exit Function
		End If
		
		set oCmd.ActiveConnection = oConn
		oCmd.CommandTimeout = cint(GetConfigValueDefault("CMD_TIMEOUT",CMD_TIMEOUT))
		oCmd.CommandType = 1
		for i = 0 to iUBound
			oCmd.commandtext = asSQL(i)
			WriteDebug "Executing " & asSQL(i)
	
			oConn.Errors.Clear
			oCmd.Execute ,,adExecuteNoRecords
			if Err.Number <> 0 then
				WriteError "Unable to execute sql: "
				WriteCode asSQL(i) 
				WriteError Err.Description
				set oCmd = nothing
				oConn.close
				set oConn = nothing
				Exit Function
			else
				for each oErr in oConn.Errors
					WriteWarning "SQL Msg: " & oErr.Number & ":" & oErr.Description
				next
			end if
		next
		ExecuteSQL = true
		
		set oCmd = nothing
		oConn.close
		set oConn = nothing
	else
		WriteError "No SQL Found"
	end if
End Function 	
'EOF ExecuteSQL

'----------------------------------
' Function ExecuteSQLReturnCol0
'	Executes SQL against the server/database specified and returns the first column
'	
' 	Parameters
'		vSQL - SQL Statement to execute
'		vCRLF - True = add a CR/LF between each record
'	Returns
'		the recordset returned as a string
'----------------------------------
Function ExecuteSQLRetCol0(byval vSQL, byval vCRLF)
	dim i, iUBound, oErr, st, cr
	dim oConn, oCmd, oRS

	set oConn = nothing
	set oCmd = nothing
	set oRS = ExecuteSQLRetRS(oConn, oCmd, vSQL)
	if oRS is nothing then
		ExecuteSQLRetCol0 = ""
		Exit Function
	end if

	set st = new FastString
	if vCRLF then
		cr = vbcrlf
	end if

	while not oRS.EOF
		st() = oRS(0) & cr
		oRS.MoveNext
	wend
	
	ExecuteSQLRetCol0 = st()
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing
End Function 	
'EOF ExecuteSQLRetCol0

'----------------------------------
' Function ExecuteSQLRetRSString(
'	Executes SQL against the server/database specified and returns all columns as a string
'	
' 	Parameters
'		vSQL - SQL Statement to execute
'	Returns
'		the recordset returned as a string
'
'----------------------------------
Function ExecuteSQLRetRSString(byval vSQL)
	dim oErr
	dim oConn, oCmd, oRS
	dim sResults, bDone
	dim st

	set st = new FastString
	set oConn = nothing
	set oCmd = nothing

	set oRS = ExecuteSQLRetRS(oConn, oCmd, vSQL)
	if oRS is nothing then
		ExecuteSQLRetRSString = ""
		Exit Function
	end if
	if oRS.State = 0 then 'closed'
		ExecuteSQLRetRSString = ""
		Exit Function
	end if

	if not oRS.EOF then
		st() = RecordsetToString(oRS)
		bDone = false
		'WriteInfo oRS.source
		while not bDone
			On Error resume next
			set oRS = oRS.NextRecordSet()
			if err.number <> 0 or oRS is Nothing then 
				bDone = true
			elseif oRS.state = 0 then 'closed'
				bDone = true
			elseif not oRS.eof then
				st() = vbcrlf & "-----------------------------------------------" & vbcrlf
				st() = RecordsetToString(oRS)
				'WriteInfo oRS.source
			else
				bDone = true
			end if
		Wend
		ExecuteSQLRetRSString = st()
	end if
	On Error resume next
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing
	err.clear
End Function 	
'EOF ExecuteSQLRetRSString

'----------------------------------
' Function ExecuteSQLDisplayBrowser(
'	Executes SQL against the server/database specified and displays results in a browser
'	
' 	Parameters
'		vSQL - SQL Statement to execute
'	Returns
'		the recordset returned as a string
'
'----------------------------------
Function ExecuteSQLDisplayBrowser(byval vSQL)
	dim oErr
	dim oConn, oCmd, oRS
	dim sResults, bDone
	dim st

	set st = new FastString
	set oConn = nothing
	set oCmd = nothing

	set oRS = ExecuteSQLRetRS(oConn, oCmd, vSQL)
	if oRS is nothing then
		ExecuteSQLDisplayBrowser = ""
		Exit Function
	end if
	if oRS.State = 0 then 'closed'
		ExecuteSQLDisplayBrowser = ""
		Exit Function
	end if

	dim objExplorer
	Set objExplorer = GetBrowserWindow()

	if not oRS.EOF then
		bDone = false
		
		call addResultsToBrowser (objExplorer, RecordsetToHTMLTable(oRS),true, vSQL)
		while not bDone
			On Error resume next
			set oRS = oRS.NextRecordSet()
			if err.number <> 0 or oRS is Nothing then 
				bDone = true
			elseif oRS.state = 0 then 'closed'
				bDone = true
			elseif not oRS.eof then
				call addResultsToBrowser (objExplorer, RecordsetToHTMLTable(oRS),false, vSQL)
			else
				bDone = true
			end if
		Wend
	end if
	
	ExecuteSQLDisplayBrowser = "success"
	
	On Error resume next
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing
	err.clear
End Function 	
'EOF ExecuteSQLDisplayBrowser



Sub addResultsToBrowser(byref rExplorer, byval vString, byval vMinimize, byval vSQL)

	err.clear
	on error resume next
	dim objDocument, oDiv
	dim sResults, sSpanID, sSpanSQLID, sPrevSpanID

	dim objExplorer
	Set objExplorer = rExplorer 'GetBrowserWindow()

	Set objDocument = objExplorer.Document
	set oDiv = objDocument.getElementById("divData")
	if oDiv is nothing then
		dim st
		SET st = new FastString
		st() = HTML_HEADER & HTML_FOOTER
		objDocument.Open
		objDocument.Writeln st()
		
		'Wscript.Sleep 2000
		objDocument.Close
		set oDiv = objDocument.getElementById("divData")
	end if

	if oDiv is nothing then
		WriteError "Unable to reference browser"
		SET objDocument = Nothing
		Exit Sub
	end if

	
	gDataCnt = gDataCnt + 1
	sPrevSpanID = "span" & gDataCnt - 1
	sSpanID = "span" & gDataCnt
	sSpanSQLID = "spansql" & gDataCnt
	
	sResults = " <div class='divBatch' >" _
	& " <span onClick='ShowHideToggle(""" & sSpanID & """,false)'" _
	& "title='Click to toggle show/hide' >" _
	& HTML_BATCH_HEADER & gDataCnt & " : " & Now _
	& "</span>&nbsp;&nbsp;&nbsp;&nbsp;<span onClick='ShowHideToggle(""" &  sSpanSQLID & """,false)'" _
	& "title='Click to toggle SQL Display' >[SQL]</span>" _
	& "&nbsp;&nbsp;&nbsp;&nbsp;<span onClick='RemoveItem(""div" &  sSpanID & """)'" _
	& "title='Click to remove results' >[X]</span>" _
	& "</div>" _
	& "<span id='" & sSpanSQLID & "' class='hideSpanData' >" _
	& vSQL _
	& "</span>" _
	& "<span id='" & sSpanID & "' class='showSpanData' >" _
	& vString & "</span>"

	if vMinimize and gDataCnt > 1 then
		call objDocument.parentWindow.ShowHideToggle(sPrevSpanID,true)
	end if

	call objDocument.parentWindow.AddItem("div" & sSpanID, sResults)
	WriteDebug sResults

	set oDiv = nothing
	set objDocument = nothing

end sub
'EOS addResultsToBrowser

'----------------------------------
' Function GetBrowserWindow(
'	Executes SQL against the server/database specified and returns insert statements for each row
'	
'	Returns
'		Browser object
'
'----------------------------------
Function GetBrowserWindow()
	dim i
	err.clear
	
	if (not GetConfigValueDefault("REUSE_BROWSER", REUSE_BROWSER)) or (gBrowser is nothing) then
		Set gBrowser = WScript.CreateObject ("InternetExplorer.Application", "IE_")
		gBrowser.Navigate "about:blank"
	end if

	on error resume next
	i = gBrowser.ToolBar		
	if err.number <> 0 then
		Set gBrowser = WScript.CreateObject ("InternetExplorer.Application", "IE_")
		gBrowser.Navigate "about:blank"
		err.clear
	end if
	on error goto 0
	gBrowser.ToolBar = HTML_BROWSER_TOOLBAR
	gBrowser.StatusBar = 1
	gBrowser.Width = HTML_BROWSER_WIDTH
	gBrowser.Height = HTML_BROWSER_HEIGHT
	'objExplorer.Left = 0
	'objExplorer.Top = 0
	gBrowser.Visible = 1

	set GetBrowserWindow = gBrowser

End Function


'----------------------------------
' Function ExecuteSQLRetInserts(
'	Executes SQL against the server/database specified and returns insert statements for each row
'	
' 	Parameters
'		vSQL - SQL Statement to execute
'		vTableName - name of the table to receive the insert statements
'	Returns
'		the recordset returned as a string
'
'----------------------------------
Function ExecuteSQLRetInserts(byval vSQL, byval vTableName)
	dim oConn, oCmd, oRS
	set oConn = nothing
	set oCmd = nothing
	
	set oRS = ExecuteSQLRetRS(oConn, oCmd, vSQL)
	if oRS is nothing then
		Exit Function
	end if
	if not oRS.EOF then
		ExecuteSQLRetInserts = RecordsetToInsertSQL(oRS, vTableName, "", false, "")
	end if
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing
End Function
'EOF ExecuteSQLRetInserts


'----------------------------------
' Function ExecuteSQLRetInsertsCheck(
'	Executes SQL against the server/database specified and returns insert statements for each row
'  This version also adds a SQL check to either not add if exists, or perform an update
'	
' 	Parameters
'		vSQL - SQL Statement to execute
'		vTableName - name of the table to receive the insert statements
'	Returns
'		the recordset returned as a string
'
'----------------------------------
Function ExecuteSQLRetInsertsCheck(byval vSQL, byval vTableName, byval vKeyField, byval vUpdateIfExists, byval vIdentityField)
	dim oConn, oCmd, oRS
	set oConn = nothing
	set oCmd = nothing
	
	set oRS = ExecuteSQLRetRS(oConn, oCmd, vSQL)
	if oRS is nothing then
		Exit Function
	end if
	if not oRS.EOF then
		if isOracle() then
			ExecuteSQLRetInsertsCheck = RecordsetToInsertSQLOracle(oRS, vTableName, vKeyField, vUpdateIfExists)
		else
			ExecuteSQLRetInsertsCheck = RecordsetToInsertSQL(oRS, vTableName, vKeyField, vUpdateIfExists, vIdentityField)
		end if
	end if
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing
End Function
'EOF ExecuteSQLRetInserts


'----------------------------------
' Function ExecuteSQLRetStatements(
'	Executes SQL against the server/database specified and returns vStatement for each record.
'	Substitutions are performed for field names.
'	
' 	Parameters
'		vSQL - SQL Statement to execute
'		vStatement - Statement to repeat
'	Returns
'		string
'
'----------------------------------
Function ExecuteSQLRetStatements(byval vSQL, byval vStatement, byval vEscapeQuotes)
	dim oConn, oCmd, oRS
	set oConn = nothing
	set oCmd = nothing
	
	WriteDebug("ExecuteSQLRetStatements: SQL:" & vSQL & vbcrlf & "Statement:" & vStatement)
	 
	set oRS = ExecuteSQLRetRS(oConn, oCmd, vSQL)
	if oRS is nothing then
		Exit Function
	end if
	if not oRS.EOF then
		WriteDebug("ExecuteSQLRetStatements: NOT EOF")
		ExecuteSQLRetStatements = RecordsetToStatements(oRS, vStatement, vEscapeQuotes)
	end if
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing

End Function
'EOF ExecuteSQLRetStatements

'----------------------------------
' Function ExecuteSQLRetRS
'	Given a sql statement, create a connection, command, and recordset 
'	of the data'
'	
' 	Parameters
'		rConn - Connection
'		rCmd - Command
'		vSQL - SQL statement to execute'
'	Returns
'		Recordset based upon vSQL
'----------------------------------
Function ExecuteSQLRetRS(ByRef rConn, ByRef rCmd, byval vSQL)
	dim oErr
	dim oConn, oCmd, oRS
	dim bCreatedConn
	if vSQL <> "" then
		On Error resume next

			WriteDebug "In ExecuteSQLRetRS: SQL = " & vSQL

		if rConn is nothing then 
			set oConn = GetConnection()
			if oConn is Nothing then
				set ExecuteSQLRetRS = nothing
				Exit Function
			end if
			bCreatedConn = true
			set oCmd = createobject("adodb.Command")
		else
			set oConn = rConn
			set oCmd = rCmd
			bCreatedConn = false
		end if
		
		If Err.Number <> 0 then
			WriteError "Unable to create ADO objects:" & Err.Description
			if bCreatedConn then oConn.close
			set oConn = nothing
			set ExecuteSQLRetRS = nothing
			Exit Function
		End If
		
		set oCmd.ActiveConnection = oConn
		oCmd.CommandType = 1
		oCmd.CommandTimeout = cint(GetConfigValueDefault("CMD_TIMEOUT",CMD_TIMEOUT))
		oCmd.CommandText = vSQL

		oConn.Errors.Clear
		set oRS = oCmd.Execute
		
		if Err.Number <> 0 then
			WriteError "Unable to execute sql: "
			WriteCode vSQL
			WriteError Err.Description
			set oCmd = nothing
			if bCreatedConn then oConn.close
			set oConn = nothing
			set ExecuteSQLRetRS = nothing
			Err.Clear
			Exit Function
		elseif oConn.Errors.Count > 0 then
			for each oErr in oConn.Errors
				WriteWarning "SQL Msg: " & oErr.Number & ":" & oErr.Description
			next
		end if
		if bCreatedConn then 
			set rConn = oConn
			set rCmd = oCmd
		end if
		set ExecuteSQLRetRS = oRS
	else
		WriteError "No SQL Found"
		set ExecuteSQLRetRS = nothing
	end if
End Function
'EOF ExecuteSQLRetRS'

'----------------------------------
' Function RecordsetToString
'	Traverses the given recordset and displays the records as text
'	
' 	Parameters
'		vRS - Recordset to traverse
'	Returns
'		a string of records in a table format
'----------------------------------
Function RecordsetToString(byval vRS)
	Const adDBTimeStamp = 135
	
	dim st, cr, i, iUBound, bOracle2
	
	redim aTmp(2)
	redim aFieldNames(vRS.Fields.count-1)
	dim sText, sTmp, iSpc, iCnt
	
	set st = new FastString
	iUBound = vRS.Fields.Count-1
	iCnt = -1
	bOracle2 = isOracleNew()
	for i = 0 to iUBound
		if not SkipType(vRS.Fields(i).type) then 
			iCnt = iCnt + 1
			redim aTmp(2)
			aTmp(0) = vRS.Fields(i).name '& ":" & vRS.Fields(i).Type
			if not bOracle2 then
				aTmp(1) = vRS.Fields(i).DefinedSize + 2
				if vRS.Fields(i).type = adDBTimeStamp then
					aTmp(1) = 25
				end if
			else
				aTmp(1) = vRS.Fields(i).OraMaxDSize + 2
			End If
			if aTmp(1) > MaxFieldLength then
				aTmp(1) = MaxFieldLength
			end if
			if aTmp(1) < len(aTmp(0)) + 2 then
				aTmp(1) = len(aTmp(0)) + 2
			end if
			aTmp(2) = i
			iSpc = aTmp(1) - len(aTmp(0))
			st() = aTmp(0) 
			st() = space(iSpc)
			aFieldNames(iCnt) = aTmp
		end if
	next
	st() = vbcrlf
	iUBound = iCnt
	redim preserve aFieldNames(iUBound)
	iCnt = 0
	dim bDone
	bDone = (vRS.EOF)
	while iCnt < MaxRecords and not bDone 
		iCnt = iCnt + 1
		for i = 0 to iUBound
			sTmp = left(("" & vRS.Fields(aFieldNames(i)(2)).value),MaxFieldLength)
			
			iSpc = aFieldNames(i)(1) - len(sTmp)
			st() = sTmp
			if iSpc > 0 then
				st() =  space( iSpc )
			end if
		next
		st() = vbcrlf
		On Error RESUME NEXT
		vRS.MoveNext
		IF Err.Number <> 0 then
			st() = "Error: " & Err.Description & vbcrlf
			Err.Clear
			bDone = true
		ELSE
			bDone = (vRS.EOF)
		END IF	
			
	wend
	st() = iCnt & " Record(s) Returned "
	RecordsetToString = st()
End Function
'EOF RecordsetToString'

'----------------------------------
' Function RecordsetToHTMLTable
'	Takes a given recordset and returns an html table
'	
' 	Parameters
'		vRS - Recordset to traverse
'	Returns
'		string
'----------------------------------
Function RecordsetToHTMLTable(byval vRS)
	Const adDBTimeStamp = 135
	
	dim st, stHead, cr, i, iUBound
	
	redim aTmp(2)
	redim aFieldNames(vRS.Fields.count-1)
	dim sText, sTmp, iSpc, iCnt
	
	set st = new FastString
	set stHead = new FastString
	iUBound = vRS.Fields.Count-1
	iCnt = -1
	stHead() = "<table class='clTable'>"
	
	for i = 0 to iUBound
		if not SkipType(vRS.Fields(i).type) then 
			iCnt = iCnt + 1
			redim aTmp(2)
			aTmp(0) = vRS.Fields(i).name '& ":" & vRS.Fields(i).Type

			if vRS.Fields(i).type = adDBTimeStamp then
				aTmp(1) = 25
			else
				aTmp(1) = len(aTmp(0)) + 2
			end if
			
			aTmp(2) = i

			aFieldNames(iCnt) = aTmp
		end if
	next
	
	st() = vbcrlf
	iUBound = iCnt
	redim preserve aFieldNames(iUBound)
	iCnt = 0
	dim bDone, sClRow
	bDone = (vRS.EOF)

	if bDone then
		st() = "<tr><td colspan='" & iUBound & "' >No records returned</td></tr>"
	end if

	while iCnt < MaxRecords and not bDone 
		iCnt = iCnt + 1
		st() = "<tr class='"
		if sClRow = "oddRow" then
			sClRow = "evenRow"
		else 
			sClRow = "oddRow"
		end if
		st() = sClRow & "'>"
		for i = 0 to iUBound
			st() = "<td>"
			sTmp = left(("" & vRS.Fields(aFieldNames(i)(2)).value),MaxFieldLength)
			
			iSpc = len(sTmp)
			st() = replace(replace(sTmp,"<","&lt;"),">","&gt;")
			if iSpc = 0 then
				st() =  "&nbsp;"
			end if
			
			if iSpc > aFieldNames(i)(1) and iSpc <= HTML_MAX_CHAR_WIDTH then 
				aFieldNames(i)(1) = iSpc
			elseif iSpc > aFieldNames(i)(1) then
				aFieldNames(i)(1) = HTML_MAX_CHAR_WIDTH
			end if

			st() = "</td>"
		next
		st() = "</tr>"
		
		
'		On Error RESUME NEXT
		vRS.MoveNext
		IF Err.Number <> 0 then
			st() = "Error: " & Err.Description & vbcrlf
			Err.Clear
			bDone = true
		ELSE
			bDone = (vRS.EOF)
		END IF	
'		On Error GOTO 0
	wend
	st() = "</table><br />"
	st() = iCnt & " Records Returned" 
	


	stHead() = "<tr class='headingRow'>"
	for i = 0 to iUBound
		stHead() = "<td>"
		stHead() = aFieldNames(i)(0)
		iSpc = aFieldNames(i)(1) - len(aFieldNames(i)(0))
		while iSpc > 0
			stHead() = "&nbsp;"
			iSpc = iSpc - 1
		wend 
		stHead() = "</td>"
	next
	stHead() = "</tr>"

	stHead() = st()
	
	RecordsetToHTMLTable = stHead() 

End Function
'EOF RecordsetToBrowser

	Sub IE_onQuit()
	 Wscript.Quit
	End Sub

'----------------------------------
' Function RecordsetToInsertSQL
'	Creates insert statements from the given recordset
'	
' 	Parameters
'		vRS - Recordset to traverse
'		vTableName - Name of table to use for the insert statements
'	Returns
'		the insert statements as a string
'----------------------------------
Function RecordsetToInsertSQL(byval vRS, byval vTableName, byval vKeyFields, byval vUpdateIfExists, byval vIdentityField)
	dim st, i, i2, iUBound, iIndex
	redim aFields(vRS.Fields.count-1)
	dim iCnt
	dim sInsertFields
	dim sFieldValue
	redim aTmp(1)
	dim sCmdTerminator
	dim iIdentityField, sIdentityField
	
	if isOracle() then
		sCmdTerminator = ";"
	else
		sCmdTerminator = ""
	end if
	
	set st = new FastString
	iUBound = vRS.Fields.Count-1
	iIndex = 0
	iIdentityField = -1
	sIdentityField = UCase(vIdentityField)
	for i = 0 to iUBound
		if not SkipType(vRS.Fields(i).type) then
			redim aTmp(1)
			aTmp(0) =  vRS.Fields(i).name
			aTmp(1) = i
			aFields(iIndex) = aTmp 
			if StrComp(aTmp(0),sIdentityField,1) = 0 then
				iIdentityField = iIndex
			end if
			iIndex = iIndex + 1
		end if
	next
	if iIndex = 0 then
		RecordsetToInsertSQL = "No supported fields are available."
		Exit Function
	end if
	iUBound = iIndex - 1
	redim preserve aFields(iUBound)
	sInsertFields = "INSERT INTO " & vTableName & " ("
	for i = 0 to iUBound - 1
		if i <> iIdentityField then
			sInsertFields = sInsertFields & aFields(i)(0) & ","
		end if
	next
	sInsertFields = sInsertFields & aFields(iUBound)(0) & ")"
	st() = vbcrlf
	iCnt = 0
	
	if vKeyFields > "" then
		dim sInsertCheck 
		dim asKeyFields, iKeyFields
		vKeyFields = replace(vkeyfields," ","")
		asKeyFields = split(vKeyFields,",")
		iKeyFields = ubound(asKeyFields)
		
		sInsertCheck = "IF NOT EXISTS(SELECT ["
		sInsertCheck = sInsertCheck & asKeyFields(0) & "] FROM [" & vTableName & "] WHERE ["
			
		while iCnt < MaxRecords and (not vRS.EOF)
			iCnt = iCnt + 1
			st() = sInsertCheck 

			for i = 0 to iKeyFields-1
				st() = asKeyFields(i)  
				st() = "] = "
				sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(i)),vRS.Fields(asKeyFields(i)).type)
				st() = sFieldValue & " AND [" 
			next

			st() = asKeyFields(iKeyFields)  
			st() = "] = "
			sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(iKeyFields)),vRS.Fields(asKeyFields(iKeyFields)).type)
			st() = sFieldValue & ") " 
			
			st() = vbcrlf & vbtab & sInsertFields & vbcrlf  
			st() = vbtab & "VALUES( "
			
			for i = 0 to iUBound -1
				if i <> iIdentityField then
					sFieldValue = FieldValueToString(vRS.Fields(aFields(i)(1)),vRS.Fields(aFields(i)(1)).type)
					st() = sFieldValue & ","
				end if
			next

			sFieldValue = FieldValueToString(vRS.Fields(aFields(iUBound)(1)),vRS.Fields(aFields(iUBound)(1)).type)
			st() = sFieldValue & ")" 
			st() = sCmdTerminator
			st() = vbcrlf 
			
			if vUpdateIfExists then
				dim bKeyField
				st() = "ELSE" & vbcrlf & vbtab
				st() = "UPDATE [" & vTableName & "] SET "

				for i = 0 to iUBound -1
					bKeyField = false
					for i2 = 0 to iKeyFields
						if ucase(aFields(i)(0)) = ucase(asKeyFields(i2)) then
							bKeyField = true
							exit for
						end if
					next
					if not bKeyField then
						st() = "[" & aFields(i)(0) & "] = "

						sFieldValue = FieldValueToString(vRS.Fields(aFields(i)(1)),vRS.Fields(aFields(i)(1)).type)
						st() = sFieldValue & ","
					end if
				next
				st() = "[" & aFields(iUBound)(0) & "] = "
				sFieldValue = FieldValueToString(vRS.Fields(aFields(iUBound)(1)),vRS.Fields(aFields(iUBound)(1)).type)
				st() = sFieldValue 

				st() = vbcrlf & vbtab & "WHERE [" 
				
				for i = 0 to iKeyFields-1
					st() = asKeyFields(i)  
					st() = "] = "
					sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(i)),vRS.Fields(asKeyFields(i)).type)
					st() = sFieldValue & " AND [" 
				next
	
				st() = asKeyFields(iKeyFields)  
				st() = "] = "
				sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(iKeyFields)),vRS.Fields(asKeyFields(iKeyFields)).type)
				st() = sFieldValue
			end if
			st() = vbcrlf & vbcrlf
	
			vRS.MoveNext
		wend
	
	else
		while iCnt < MaxRecords and (not vRS.EOF)
			iCnt = iCnt + 1
			st() = sInsertFields 
			st() = "VALUES( "
			
			for i = 0 to iUBound -1
				sFieldValue = FieldValueToString(vRS.Fields(aFields(i)(1)),vRS.Fields(aFields(i)(1)).type)
				st() = sFieldValue & ","
			next

			sFieldValue = FieldValueToString(vRS.Fields(aFields(iUBound)(1)),vRS.Fields(aFields(iUBound)(1)).type)
			st() = sFieldValue & ")" 
			st() = sCmdTerminator
			st() = vbcrlf
			vRS.MoveNext
		wend
	end if
	WriteInfo iCnt & " Record(s) Returned "
	RecordsetToInsertSQL = st()
End Function
'EOF RecordsetToInsertSQL'

'----------------------------------
' Function RecordsetToInsertSQLOracle
'	Creates insert statements from the given recordset
'	
' 	Parameters
'		vRS - Recordset to traverse
'		vTableName - Name of table to use for the insert statements
'	Returns
'		the insert statements as a string
'----------------------------------
Function RecordsetToInsertSQLOracle(byval vRS, byval vTableName, byval vKeyFields, byval vUpdateIfExists)
	dim st, i, i2, iUBound, iIndex
	redim aFields(vRS.Fields.count-1)
	dim iCnt
	dim sInsertFields
	dim sFieldValue
	redim aTmp(1)
	
	set st = new FastString
	iUBound = vRS.Fields.Count-1
	iIndex = 0
	
	for i = 0 to iUBound
		if not SkipType(vRS.Fields(i).type) then
			redim aTmp(1)
			aTmp(0) =  vRS.Fields(i).name
			aTmp(1) = i
			aFields(iIndex) = aTmp 
			iIndex = iIndex + 1
		end if
	next
	if iIndex = 0 then
		RecordsetToInsertSQLOracle = "No supported fields are available."
		Exit Function
	end if
	iUBound = iIndex - 1
	redim preserve aFields(iUBound)
	sInsertFields = "INSERT INTO " & vTableName & " ("
	for i = 0 to iUBound - 1
		sInsertFields = sInsertFields &  aFields(i)(0) & ", "
	next
	sInsertFields = sInsertFields & aFields(iUBound)(0) & ") "

	st() = vbcrlf
	iCnt = 0
	
	if vKeyFields > "" then
		dim sInsertCheck 
		dim asKeyFields, iKeyFields
		vKeyFields = replace(vkeyfields," ","")
		asKeyFields = split(vKeyFields,",")
		iKeyFields = ubound(asKeyFields)
		
		sInsertCheck = "DECLARE" & vbcrlf & vbtab & "l_exists INTEGER;" & vbcrlf
		sInsertCheck = sInsertCheck & "BEGIN" & vbcrlf & vbtab 
		sInsertCheck = sInsertCheck & "SELECT COUNT(*) INTO l_exists" & vbcrlf & vbtab & "FROM " & vTableName 
		sInsertCheck = sInsertCheck & vbcrlf & vbtab & "WHERE "
			
		while iCnt < MaxRecords and (not vRS.EOF)
			iCnt = iCnt + 1
			st() = sInsertCheck 

			for i = 0 to iKeyFields-1
				st() = asKeyFields(i)  
				st() = " = "
				sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(i)),vRS.Fields(asKeyFields(i)).type)
				st() = sFieldValue & " AND " 
			next

			st() = asKeyFields(iKeyFields)  
			st() = " = "
			sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(iKeyFields)),vRS.Fields(asKeyFields(iKeyFields)).type)
			st() = sFieldValue 
			st() = vbcrlf & vbtab
			st() = "AND ROWNUM = 1;"
			st() = vbcrlf & vbtab
			st() = "IF l_exists <> 1 THEN"
			st() = vbcrlf & vbtab & vbtab
			st() = sInsertFields   
			st() = vbcrlf & vbtab & vbtab 
			st() = "VALUES( "
			
			for i = 0 to iUBound -1
				sFieldValue = FieldValueToString(vRS.Fields(aFields(i)(1)),vRS.Fields(aFields(i)(1)).type)
				st() = sFieldValue & ","
			next

			sFieldValue = FieldValueToString(vRS.Fields(aFields(iUBound)(1)),vRS.Fields(aFields(iUBound)(1)).type)
			st() = sFieldValue & ");" 
			st() = vbcrlf & vbtab
			
			if vUpdateIfExists then
				dim bKeyField
				st() = "ELSE" 
				st() = vbcrlf & vbtab & vbtab
				st() = "UPDATE " & vTableName & " SET "

				for i = 0 to iUBound -1
					bKeyField = false
					for i2 = 0 to iKeyFields
						if ucase(aFields(i)(0)) = ucase(asKeyFields(i2)) then
							bKeyField = true
							exit for
						end if
					next
					if not bKeyField then
						st() = aFields(i)(0) & " = "

						sFieldValue = FieldValueToString(vRS.Fields(aFields(i)(1)),vRS.Fields(aFields(i)(1)).type)
						st() = sFieldValue & ","
					end if
				next
				st() = aFields(iUBound)(0) & " = "
				sFieldValue = FieldValueToString(vRS.Fields(aFields(iUBound)(1)),vRS.Fields(aFields(iUBound)(1)).type)
				st() = sFieldValue 

				st() = vbcrlf & vbtab & vbtab & "WHERE " 
				
				for i = 0 to iKeyFields-1
					st() = asKeyFields(i)  
					st() = " = "
					sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(i)),vRS.Fields(asKeyFields(i)).type)
					st() = sFieldValue & " AND " 
				next
	
				st() = asKeyFields(iKeyFields)  
				st() = " = "
				sFieldValue = FieldValueToString(vRS.Fields(asKeyFields(iKeyFields)),vRS.Fields(asKeyFields(iKeyFields)).type)
				st() = sFieldValue
				st() = ";" & vbcrlf & vbtab
			end if
			st() = "END IF ;"
			st() = vbcrlf 
			st() = "END ;"
			st() = vbcrlf
			st() = "/" & vbcrlf
	
			vRS.MoveNext
		wend
	
	else
		while iCnt < MaxRecords and (not vRS.EOF)
			iCnt = iCnt + 1
			st() = sInsertFields 
			st() = "VALUES( "
			
			for i = 0 to iUBound -1
				sFieldValue = FieldValueToString(vRS.Fields(aFields(i)(1)),vRS.Fields(aFields(i)(1)).type)
				st() = sFieldValue & ","
			next

			sFieldValue = FieldValueToString(vRS.Fields(aFields(iUBound)(1)),vRS.Fields(aFields(iUBound)(1)).type)
			st() = sFieldValue & ")" & vbcrlf 

			vRS.MoveNext
		wend
	end if
	WriteInfo iCnt & " Record(s) Returned "
	RecordsetToInsertSQLOracle = st()
End Function
'EOF RecordsetToInsertSQLOracle'




Function FieldValueToString(byval vValue, byval vType)
	Const adBit = 11
	Const adBigInt = 20
	Const adCurrency = 6
	Const adDecimal = 14
	Const adDouble = 5
	Const adInteger = 3
	Const adNumeric = 131
	Const adSingle = 4
	Const adSmallInt = 2
	Const adTinyInt = 16
	

	if isnull(vValue) then
		FieldValueToString = "null"
		Exit Function
	end if

	if vType = adBit then
		if vValue then
			FieldValueToString = "1" 
		else
			FieldValueToString = "0"		
		end if
		Exit Function
	elseif vType = adBigInt or vType = adCurrency or vType = adDecimal _
			or vType = adDouble or vType = adInteger or vType = adNumeric _
			or vType = adSingle or vType = adSmallInt or vType = adTinyInt then

		FieldValueToString = "" & vValue
		Exit Function
	end if
	vValue = "" & vValue
	
	FieldValueToString = "'" & replace(vValue,"'","''") & "'"
End Function

'------------------------------------------------------------------------------!!!!!!!

'----------------------------------
' Function ExecTableToClass
'	Generates a C# Class based upon the table definition
'	
' 	Parameters

'		vTableName - name of the table
'	Returns
'		the C# class returned as a string
'
'----------------------------------
Function ExecTableToClass(byval vConn, byval vCmd, byval vTableName)
	dim oConn, oCmd, oRS, sSQL
	set oConn = nothing
	set oCmd = nothing
	
	sSQL = "SELECT table_name, column_name,is_nullable,data_type,character_maximum_length,numeric_precision, numeric_scale " _
		& "FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name = '" & vTableName & "' order by ordinal_position "

	set oRS = ExecuteSQLRetRS(oConn, oCmd, sSQL)
	if oRS is nothing then
		WriteDebug "oRS is nothing: Table " & vTableName + " not found?"
		Exit Function
	end if
	if not oRS.EOF then
		WriteDebug "calling RecordsetToClass"
		ExecTableToClass = RecordsetToClass(oRS, vTableName) ', "", false)
	else
		WriteDebug "oRS is EOF?"
	end if
	oRS.close
	set oRS = nothing
	set oCmd = nothing
	oConn.close
	set oConn = nothing
End Function
'EOF ExecTableToClass



'----------------------------------
' Function RecordsetToClass
'	Creates a C# Class from the given recordset of a table definition
'	
' 	Parameters
'		vRS - Recordset to traverse
'		vTableName - Name of table to use for the class
'	Returns
'		the insert statements as a string
'----------------------------------
Function RecordsetToClass(byval vRS, byval vTableName)
on error goto 0
	dim stVars, stProps, sFieldType, sNetFieldType
	dim sPrivateName, sPublicName
	dim sResult, iFldCount
	
	set stVars = new FastString
	set stProps = new FastString

	stVars() = vbcrlf
	stProps() = vbcrlf
	iFldCount = 0
	sResult = Replace(CSHARP_CLASS_START_TEMPLATE, "<classname>", getClassNameFromTable(vTableName))
	WriteDebug "Creating class:" & vTableName
	WHILE not vRS.EOF
		sNetFieldType = getDotNetFieldType(vRS("data_type"))
		WriteDebug "1: " + sNetFieldType
		iFldCount = iFldCount + 1

		sPublicName = getPropertyNameFromField(vRS("column_name"))
		WriteDebug "2: " + sPublicName
		
		stProps() = Replace(Replace(CSHARP_PROPERTY_TEMPLATE, "<fieldtype>",sNetFieldType),"<fieldname>",sPublicName)
		
		vRS.MoveNext
	
	WEND
	stProps() = CSHARP_CLASS_END_TEMPLATE
	RecordsetToClass = sResult &  stProps()
End Function
	
Function getPropertyNameFromField(byval fld)
	getPropertyNameFromField = replace(fld, "_","")

End Function	


Function getClassNameFromTable(byval vTable)
		getClassNameFromTable = replace(replace(vTable, "_", ""), "tbl","")
	
End Function


Function getDotNetFieldType(byval vType)

	WriteDebug "Converting type " & vType

	If vType = "varchar" or vtype = "char" or vtype = "text" or _
			vType = "nvarchar" or vType = "nchar" or vType = "ntext" then
		getDotNetFieldType = "string"
		
	elseif vType = "binary" or vType = "varbinary" then
		getDotNetFieldType = "Byte[]"
		
	elseif vType = "int" then
		getDotNetFieldType = "int"

	elseif vType = "smallint" then
		getDotNetFieldType = "int"

	elseif vType = "tinyint" then
		getDotNetFieldType = "Byte"
		
	elseif vType = "bigint" then
		getDotNetFieldType = "long"

	elseif vType = "money" or vType = "smallmoney" or vType = "numeric" or vType = "decimal" then
		getDotNetFieldType = "Decimal"
				
	elseif vType = "datetime" or vType = "smalldatetime" or vType = "datetime2" then
		getDotNetFieldType = "DateTime"
		
	elseif vType = "float" then
		getDotNetFieldType = "Double"

	elseif vType = "real" then
		getDotNetFieldType = "Single"
		
	elseif vType = "bit" then
		getDotNetFieldType = "Bool"
	else
	
		getDotNetFieldType = "UNSUPPORTED TYPE " & vType
	End If
	
End Function


'-----------------------------------------------------------------------------!!!!!!!


'----------------------------------
' Function RecordsetToStatements
'	Creates vStatements from the given recordset
'	
' 	Parameters
'		vRS - Recordset to traverse
'		vStatement - Statement to repeat
'		vEscapeQuote - if True, then double the single quotes in the string
'	Returns
'		one vStatement per record, with substitutions performed as a string
'----------------------------------
Function RecordsetToStatements(byval vRS, byval vStatement, byval vEscapeQuote)
	dim st, i, iCnt
	redim aFields(100)
	
	' Get Each Substitution
	dim sPre, iLen, sChar, stmpFld
	dim iFldCnt, sStatementPre, sValue
	
	iLen = len(vStatement)
	i = 1
	iFldCnt = 0
	set sPre = new FastString
	while i <= iLen
		sChar = mid(vStatement,i,1)
		if sChar = "{" then
			sPre() = "{" & iFldCnt & "}"
			stmpFld = ""
			i = i + 1
			sChar = mid(vStatement,i,1)
			while sChar <> "}" and i <= iLen
				
				stmpFld = stmpFld & sChar
				i = i + 1
				sChar = mid(vStatement,i,1)
			wend
			aFields(iFldCnt) = stmpFld
			iFldCnt = iFldCnt + 1
		else
			sPre() = sChar
		end if
		i = i + 1
	Wend

	redim preserve aFields(iFldCnt)
	
	sStatementPre = sPre()
	set st = new FastString

	st() = vbcrlf
	iCnt = 0

	dim sStatementPost

	On Error Resume Next
	DO While iCnt < MaxRecords and (not vRS.EOF)
	
		sStatementPost = sStatementPre	
		
		for i = 0 to iFldCnt -1
			sValue = "" & vRS.Fields(aFields(i))
			if sValue = "" then
				sValue = "null"
			elseif vEscapeQuote then
				sValue = "'" & replace(sValue, "'","''") & "'"
			end if
			WriteDebug "field:" & aFields(i) & " value:" & vRS.Fields(aFields(i))

'			Issue with odbc - if reading excel file, first time field ok, second time null
'			if isnull(vRS.Fields(aFields(i))) then
'				sValue = "null"
'			elseif vEscapeQuote then
'				sValue = "'" & replace(vRS.Fields(aFields(i)),"'","''") & "'"
'			else
'				sValue = vRS.Fields(aFields(i))
'			end if

			IF Err.Number <> 0 then
				WriteError("Error: " & Err.Number & " " & Err.Description)
				WriteDebug("Failed.  So far:" & sStatementPost)
				Err.Clear
				Exit Do
			End IF
			sStatementPost = replace(sStatementPost,"{" & i & "}",sValue)
		next
		
		st() = sStatementPost
		st() = vbcrlf
		iCnt = iCnt + 1
		vRS.MoveNext
	Loop
	st() = iCnt & " Record(s) Returned "
	RecordsetToStatements = st()
End Function
'EOF RecordsetToStatements

'----------------------------------
' Function SkipType
'	Checks the given field type against a list of non-supported field types
'	
' 	Parameters
'		vType - field type
'	Returns
'		True if vType is field type that should be skipped
'----------------------------------
Function SkipType(byval vType)
	Const adIDispatch = 9
	Const adIUnknown = 13
	Const adVariant = 12
	Const adVarBinary = 204
	Const adChapter = 136
	Const adImage = 205
	Const adBinary = 128

	select case vType
	case adIDispatch, adIUnknown,adVariant, adVarBinary, adChapter, adImage, adBinary
		SkipType = true
	case else
		SkipType = false
	end select
End Function
'EOF SkipType

'----------------------------------
' Function NewDocument
'	Sends the text to a new document.  
'	If the document already exists, the text will be replaced with this text
'	The document is then made active
'----------------------------------
Function NewDocument(byval vFileName, byval vText)
	dim oDoc, oNewDoc, bDocFound
	
	bDocFound = false
	set oDoc  = newEditor()
	On Error resume next
	if vFileName > "" then 
		call oDoc.assignEditorByName(vFileName) 
		if err.number <> 0 then
			err.clear
			call oDoc.newFile(vFileName)
			if err.number <> 0 then
				msgbox "Unable to create new editor document: " & err.number & vbcrlf & err.description
			end if
		end if
		oDoc.Text(vText)
		oDoc.SaveFile()
	else
		oDoc.newFile()
		oDoc.Text(vText)
	end if
	NewDocument = true
End Function
'EOF NewDocument

'----------------------------------
' Function OpenDocument
'	Opens the filename specified.  If the file is already open, reload it.
'----------------------------------
Function OpenDocument(byval vFileName)
	dim oDoc, oNewDoc, bDocFound
	
	set oDoc  = newEditor()
	On Error resume next
	call oDoc.assignEditorByName(vFileName) 
	if err.number <> 0 then
		err.clear
		bDocFound = oDoc.openFile(vFileName)
		if err.number <> 0 or not bDocFound then
			msgbox "Document not found"
			Exit Function
		end if
	else
		oDoc.ReloadFile()
	end if
	OpenDocument = true
End Function
'EOF OpenDocument


'-----------------------
' These functions define different ways to write to the Log window
' Currently PSPad does not support colors through script, so all functions do the same thing
'-----------------------
Function WriteError(byval vText)
	WriteToLog(vText)		
End Function

Function WriteInfo(byval vText)
	WriteToLog(vText)		
End Function

Function WriteWarning(byval vText)
	WriteToLog(vText)		
End Function

Function WriteCode(byval vText)
	WriteToLog(vText)		
End Function

Function WriteDebug(byval vText)
	if bDEBUG then
		WriteToLog(vText)		
	end if
End Function

Sub WriteToLog(byval vText)
	if gbNewLogFunctions then
		dim asLines
		asLines = split(vText,vbcrlf)
		dim i
		for i = 0 to ubound(asLines) 
			logAddLine asLines(i)
		next
	else
		SQLLog.appendText(vblf & vText)
	end if
End Sub

Sub ClearLog
	if gbNewLogFunctions then
		call logClear
	else
		SQLLog.text("")
	End If
End Sub

'--------------------------------
' Function ParseText
'	Purpose: 
'		Parses the passed text into an array of SQL batches to execute.
'		The batches are determined by the presence of the GO keyword.
'		It does not suffer from the same bug as SQL Server Query Analyzer;
'		if the GO statement is part of a commented block, it will not parse it.
'	Parameters
'		vText - SQL to parse
'	Returns
'		An array of parsed SQL batches
'--------------------------------			
Function ParseText(byval vText)
	dim asTmp, i
	dim bCommentMode
	redim asSQL(0)
	dim sCurrent
	dim iUBound, iCount
	dim st
	dim bSemicolonMode
	
	set st = new FastString
	
	asTmp = split(vText, vbcrlf)
	
	sCurrent = ""
	iUBound = 0
	iCount = 0
	bCommentMode = false
	bSemicolonMode = false
	
	for i = 0 to ubound(asTmp)
		if left(trim(asTmp(i)),2) = "/*" then	' comment
			bCommentMode = true
		end if
		if inStr(1,asTmp(i),"*/") > 0 then
			bCommentMode = false
		end if
		
		if ucase(left(trim(asTmp(i)),2)) = "GO" and not bCommentMode then ' ended a block
			if st() > "" then
				asSQL(iCount) = st()
				iCount = iCount + 1
				if iCount > iUBound then
					iUBound = iUBound + 10
					redim preserve asSQL(iUBound)
				end if
			end if
			st.Reset
		elseif st() > "" or asTmp(i) > "" then
			st() = asTmp(i) & vbcrlf
		end if
	next
	if st() > "" then
		asSQL(iCount) = st()
		iCount = iCount + 1
	end if
	redim preserve asSQL(iCount-1)
	ParseText =  asSQL
End Function
'EOF ParseText


'----------------------------------
' FastString Class by Mark Kahn
'	Purpose
'		Very fast alternative to string concatination.	
' Downloaded from 
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=9123&lngWId=4
'----------------------------------
Class FastString
	private aFStrings()
	private iFSPos
	private iFSLen
	private iFSIncr

	Private Sub Class_Initialize()
		iFSPos  = 0
		iFSIncr = 1000

		ReAllocString(iFSIncr)
	End Sub

	Private Sub Class_Terminate()
		
	End Sub

	Private Sub ReAllocString(byval length)
		if length>iFSLen then
			iFSLen = length+iFSIncr
		else
			exit sub
		end if

		redim preserve aFStrings(iFSLen)
	End Sub

	Public Property Let sContents(byval sData)
		ReAllocString(iFSPos)
		aFStrings(iFSPos) = sData
		iFSPos = iFSPos + 1
	End Property

	Public Default Property Get sContents()
		sContents = join(aFStrings, "")
	End Property

	Public Property Get JoinString(byval vSeparator)
		dim atmp
		atmp = aFStrings
		redim preserve atmp(iFSPos-1)
		JoinString = join(atmp, vSeparator)
	End Property

	Public Function Reset()
		iFSLen = iFSIncr
		redim aFStrings(iFSLen)
	End Function
End Class
