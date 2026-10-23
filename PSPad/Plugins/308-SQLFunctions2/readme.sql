-- %CN=Northwind% %DB=mydatabase% %MAX_RECORDS=300% %MAX_FIELD_LENGH=300% %CMD_TIMEOUT=30% %RESULTS_TO_FILE=false% %ALLOW_FILE_EXECUTION=false% %REUSE_BROWSER=true%

-- SQLFunctions.vbs - Set of scripts I developed for doing stuff with SQL Server.  
-- Most of the statements will work with any database with an odbc or oledb connection string.
-- License: Use at your own risk and do with it as you please.  I accept no responsibility for any lost data
-- due to anything you do with this.
--
-- Before you can use these scripts, you need to set up your databases 
-- The first line of the tab is used to configure the SQL statements executed.
-- The format is %<variable>=<value>%.
-- Possible values:
-- CN: The connection string to use.  This is usually set in each file.  Do not include the CONN_ in the value.

-- DB: Name of database.  Used to override the database in the CN connection string.  Convenient if referering multiple 
-- databases on save server

-- MAX_RECORDS: Maximum number of records returned in a SQL query

-- MAX_FIELD_LENGTH: Maximum characters displayed for a string field.  Useful for formatting the string output.

-- CMD_TIMEOUT: ADO Command timeout value.

-- RESULTS_TO_FILE: Query results default to the Log window.  If this is set to true, the results go to a new tab.

-- ALLOW_FILE_EXECUTE: When SQL is not selected, for recordset queries, the script attempts to determine the SQL statement 
-- based upon the cursor position.  For action queries, the SQL needs to be selected.  If this value is set to true, then 
-- the entire file is executed.  NOT recommended. 

-- The settings above are not required, and are only needed to override the defaults in the script file.
-- In most cases, only the %CN line is listed.  The same variables can also be set on the User tab of the User Variables form,
-- although the first line always takes priority.

-- Functions

-- Execute SQL
--	Execute the selected SQL, or the entire file if ALLOW_FILE_EXECUTION is true (not recommended).  

-- Get Data (Alt+G)/ Get Data In Browser (Alt+B)
-- 	Executes the selected SQL and returns the result.  If text is not selected and ALLOW_FILE_EXECUTION is false, 
--	the script will take the current line and check if it begins with 'select'.  If not, then the script will continue 
--	through the previous lines until it is found.
-- 	For Example:

SELECT * FROM Employees
WHERE Title = 'Sales Representative'
AND Country = 'United States'

--	If the cursor is placed on the WHERE line and query is executed (Alt+X or Alt+B), the Country clause will not be included.
--	If the cursor is placed on the SELECT line, then no WHERE clause will be included.
--	Alt+G will output the results to the Log window if RESULTS_TO_FILE is false otherwise the results are sent to a new tab.
--	Alt+B will output the results to an Internet Explorer window.  

-- Select * from <selected> / Select * from <selected> in Browser
--	This is a shortcut to retrieve all data from the table under the cursor.

-- Table List
--	This command dumps the names of all the Tables in the current database into a file called tables.sql and opens in a 
--	new tab.  If the file is already open, it is overwritten.  An error will occur if the account does not have sufficient permissions.

-- Stored Procedure List
--	This command dumps the names of all the Stored Procedures in the current database into a file called sp.sql and opens
--	in a new tab.  if the file is already open, it is overwritten.  An error will occur if the account does not have sufficient permissions.

-- View List
--	This command dumps the names of all the Views in the current database into a file called view.sql and opens
--	in a new tab.  if the file is already open, it is overwritten.  An error will occur if the account does not have sufficient permissions.

-- UDF List
--	This command dumps the names of all the User Defined Functions in the current database into a file called udf.sql and opens
--	in a new tab.  if the file is already open, it is overwritten.  An error will occur if the account does not have sufficient permissions.

-- Database List
--	This command dumps the names of all the User Defined Functions in the current database into a file called databases.sql and opens
--	in a new tab.  if the file is already open, it is overwritten.  An error will occur if the account does not have sufficient permissions.

-- Get SQL Code
--	If the account has the required permissions, the definition of the selected Stored Procedures/Views/UDF will be returned
--	as a Drop/Create.  When combined with the list functions, this is a convenient way to retrieve a database definition.

-- Get Table Defintion
--	Retrieves column information for the selected table name.  Sadly this does not generate a Create Table statement.

-- Create Inserts
--	Generate Insert statements for the provided SELECT statement.  Script will prompt for the table name to use.

-- Create Inserts with Checks
--	Generate Insert statements for the provided SELECT statement.  Script provides the option to check if the record exists,
--	and optionally perform an update if exists.  Also prompts to exclude Identity fields in the generated insert.

-- Create Statements For Each Record - {field}
--	For the given SELECT statement, repeat the selected string after the :: delimiter for each record.
--	Fields are included in the repeat string by enclosing the field names in brackets {field}.  Example:

SELECT First_Name, Last_Name 
FROM Employees :: Hello {First_Name} {Last_Name}, nice to meet you!

-- Create Statements For Each Text Record - {%1}
--	This is great for generating insert statemets from just about any type of text data.
--	The script prompts for the field seperator and record separator.  After the data put :: and enter the 
--	string to repeat for each text record.  To insert the column data in the repeated string use {%1}, {%2}, etc. for
--	the columns.  For example, here is some tab separated data.  Go ahead and try it.  Select the text below and pick
--	the function from the menu.  Use \t as the column separator and \n as the row separator
Larry	Smith	123 Fourth St	Beverly Hills	CA	90210	18
Sally	Straub	333 Third Ave	Beverly Hills	CA	90210	25
Ronald	McDonald	987 Burger Lane	New Orleanes	LA	99999	40
Jim	O'leary	777 Eighth Street	Witchita	KS	88888	99 
:: INSERT INTO Employee(First_Name, Last_Name, Address, City, State, Zip, Age) VALUES('{%1}', '{%2}', '{%3}', '{%4}', '{%5}', '{%6}',{%7})

-- Generate C# Classes for Tables
--	Create a SELECT statement and just try it.  

