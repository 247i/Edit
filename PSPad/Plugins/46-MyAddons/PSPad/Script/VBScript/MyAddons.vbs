' ------------------------------------------------------------------------------
' file      MyAddons.vbs
' date      07/28/2005 - 09/15/2006
' author    Scott Greenberg <me@gogogadgetscott.com>
' link      http://gogogadgetscott.info/computers/cc/snippet.php?sid=34
' copyright Copyright (c) 2005-2006. SEG Technology. All rights reserved.
' requires  GoGo v1.01 or greater <http://gogogadgetscott.info/computers/cc/snippet.php?sid=54>
' ------------------------------------------------------------------------------
' Available functions:
'
' Find
' Replace - Uses Regular Expressions and correctly handles newline (\n) searches.
' HTML list items - Adds li tags at beginning and end of each selected line.
' Find To Dos - Automates search in files to display todo items.
' DEC2HEX
' Toggle Value - Toggles selected value while preserving case.
'  true/false, on/off, begin/end, 0/1, +/-, public/private
'  north/south, east/west
'
' Paste to log
' List open files - Lists open files in output window.
' Join open files - Joins open files in tab order and sends result to output window.
'
' Toggle Real Tabs Settings
'
' Maximize All
' Minimize All
' Restore All
' Close All
' Save All & Exit
'
' ------------------------------------------------------------------------------
Const module_name  = "MyAddons"
Const module_ver   = "1.11"

' ------------------------------------------------------------------------------
'                                 Text editing.
' ------------------------------------------------------------------------------

Sub FindPos
    Pattern = InputBox("Find: ", module_name, "Pspad")
    aPos = GoGo.getStrPos(Pattern, FALSE, FALSE, TRUE)
    If IsArray(aPos) Then
        '// Found a match, now move caret to match.
        Set editor = newEditor()
        editor.assignActiveEditor
        editor.setCaretPos aPos(0), aPos(1)
        '// Select match
        For i = 1 To Len(Pattern)
            editor.command("ecSelRight")
        Next
    Else
        MsgBox "No match found.", 0, module_name & "::FindPos"
    End If
End Sub

Sub replaceRegExp
    Pattern = InputBox("Enter RegExp.", module_name, "([^\n\r]+)[\n\r]{1,2}")
    msg = "Enter replacement text for RegExp /" & Pattern & "/."
    newtext = InputBox(msg, module_name, "$1\+\+\+")
    text = GoGo.handleSelText("")
    text = runRegExpReplace(text, Pattern, newtext)
    GoGo.handleSelText text
End Sub

Sub listItems
    addTags "<li>", "</li>"
End Sub

Sub ToggleValue
    text = GoGo.handleSelText("")
    If text = "" Then
        Exit Sub
    End If
    Dim toggles(25, 1)

    i = 0
    toggles(i,0) = "true"
    toggles(i,1) = "false"
    i = i + 1
    toggles(i,0) = "on"
    toggles(i,1) = "off"
    i = i + 1
    toggles(i,0) = "begin"
    toggles(i,1) = "end"
    i = i + 1
    toggles(i,0) = "0"
    toggles(i,1) = "1"
    i = i + 1
    toggles(i,0) = "+"
    toggles(i,1) = "-"
    i = i + 1
    toggles(i,0) = "public"
    toggles(i,1) = "private"
    i = i + 1
    toggles(i,0) = "north"
    toggles(i,1) = "south"
    i = i + 1
    toggles(i,0) = "east"
    toggles(i,1) = "west"

    For i = 0 To UBound(toggles, 1) - 1

        tokenA = toggles(i, 0)
        tokenB = toggles(i, 1)
        If tokenA = "" Or tokenB = "" Then
            Exit Sub
        End If

        pos = InStr(1, text, tokenA, 1)
        If Not pos > 0 Then
            tokenA = toggles(i, 1)
            tokenB = toggles(i, 0)
            pos = InStr(1, text, tokenA, 1)
        End If

        If pos > 0 Then
            On Error Resume Next
            '// Check if first letter for token is Uppercase.
            temp = Mid(text, pos, 1)
            If Not IsNumeric(temp) And Asc(temp) < 97 Then

                '// Check if second letter for token is Uppercase.
                temp = Mid(text, pos + 1, 1)
                If temp = "" Or Asc(temp) < 97 Then
                    tokenB = UCase(tokenB)
                Else
                    tokenB = UCase(Left(tokenB, 1)) & Right(tokenB, Len(tokenB) - 1)
                End If
            End If

            text = Replace(text, tokenA, tokenB, pos, 1, 1)
            Exit For
        End If

    Next
    GoGo.handleSelText text
End Sub

Private Sub addTags(openTag, closeTag)
    text = GoGo.handleSelText("")
    text = runRegExpMatch(text, "[\S| |\t]*", openTag, closeTag & vbNewLine)
    GoGo.handleSelText(text)
End Sub


' ------------------------------------------------------------------------------
'                                 PSPad Control.
' ------------------------------------------------------------------------------

Sub MaxAll
    WinAll "X"
End Sub

Sub MinAll
    WinAll "N"
End Sub

Sub RestoreAll
    WinAll "R"
End Sub

Sub CloseAll
    WinAll "C%Y"
End Sub

Sub SaveAllExit
    Set SH = Wscript.CreateObject("WScript.Shell")
    SH.SendKeys "%FL"
    SH.SendKeys "%{F4}"
End Sub

Sub WinAll(Value)
    On Error Resume Next
    Set SH = Wscript.CreateObject("WScript.Shell")
    Set tmpeditor = newEditor()
    For i = 0 To editorsCount - 1
        tmpeditor.assignEditorByIndex(i)
        SH.SendKeys "%-"
        SH.SendKeys Value
        Sleep(100)
    Next
End Sub

Sub FindToDos
    Set SH = Wscript.CreateObject("WScript.Shell")
    SH.AppActivate "PSPad"
    Sleep(100)
    SH.SendKeys "%S"
    Sleep(100)
    SH.SendKeys "S"
    Sleep(100)
    SH.SendKeys "todo"
    Sleep(100)
    SH.SendKeys "{TAB 6}"
    Sleep(100)
    SH.SendKeys " "
    '     Sleep(100)
    '     SH.SendKeys "%O"
    Set SH = Nothing
End Sub

Sub GetOpenFiles()
    On Error Resume Next
    Set tmpeditor = newEditor()
    Set outputlog = NewEditor()
    outputlog.AssignLog
    outputlog.selText(vbNewLine)
    i = 0
    Do While tmpeditor.assignEditorByIndex(i)
        FileName = tmpeditor.fileName
        '// assignEditorByIndex always return 1, hack to determine when all
        '// open editors have been handled.
        If FileName = LastFileName Then
            Exit Do
        End If
        LastFileName = FileName
        outputlog.selText(FileName & vbNewLine)
        i = i + 1
    Loop
    outputlog.selText("Number of open file: " & i & vbNewLine)
End Sub

Sub JoinOpenFiles()
    'On Error Resume Next
    Set oEditor = newEditor()
    sData = ""
    For i = 0 To editorsCount - 1
        oEditor.assignEditorByIndex(i)
        sData = sData & oEditor.text & vbNewLine
    Next

    GoGo.clearLog
    GoGo.write2Log sData
End Sub

'// Paste clipboard content to log
'// Requires PSPad 4.3.3 (2065)
Sub paste2Log
    Set outputlog = NewEditor
    With outputlog
        .AssignLog
        .selText(vbNewLine)
        .command "ecPaste"
    End With
End Sub

Sub ConvertDEC2HEX
    GoGo.handleSelText(dec2hex(GoGo.handleSelText("")))
End Sub

Sub ToggleTabs
    Set SH = Wscript.CreateObject("WScript.Shell")
    SH.AppActivate "PSPad"
    Sleep(100)
    SH.SendKeys "%N"
    Sleep(100)
    SH.SendKeys "P"
    Sleep(100)
    SH.SendKeys "{UP 17}"
    SH.SendKeys "{DOWN 4}"
    SH.SendKeys "{TAB 3}"
    SH.SendKeys " "
    Sleep(100)
    SH.SendKeys "%O"
    Set SH = Nothing
End Sub


' ------------------------------------------------------------------------------
'                          File system subroutines.
' ------------------------------------------------------------------------------

Sub openFile
    Set obj1 = newEditor()
    obj1.openFile(moduleFileName(module_name))
End Sub

' ------------------------------------------------------------------------------
'                               Private functions.
' ------------------------------------------------------------------------------


Private Function runRegExpMatch(haystack, Pattern, beforMatch, afterMatch)
    Set regEx = New RegExp
    If Pattern = "" Then
        runRegExpMatch = haystack
        Exit Function
    End If
    regEx.Pattern    = Pattern
    regEx.Global     = TRUE
    regEx.IgnoreCase = TRUE
    Set Matches = regEx.Execute(haystack)
    haystack = ""
    For Each Match in Matches
        If Match.Value <> "" Then
            haystack = haystack & beforMatch & Match.Value & afterMatch
        End If
    Next
    runRegExpMatch = haystack
End Function

Private Function runRegExpReplace(haystack, Pattern, newtext)
    Set regEx = New RegExp
    If Pattern = "" Then
        runRegExpReplace = haystack
        Exit Function
    End If
    regEx.Pattern    = Pattern
    regEx.IgnoreCase = TRUE
    regEx.Global     = TRUE
    runRegExpReplace = regEx.Replace(haystack, newtext)
End Function

Private Function dec2hex(lngDec)
    On Error Resume Next
    ' Convert number to string representing hexadecimal value
    strHex = CStr(Hex(lngDec))
    ' Add leading zeros if necessary
    intLen = Len(strHex)
    If intLen < 8 Then
        Do Until intLen = 8
            strHex = "0" & strHex
            intLen = Len(strHex)
        Loop
    End If
    dec2hex = strHex
End Function

' ------------------------------------------------------------------------------
'                           Default module subroutines.
' ------------------------------------------------------------------------------

Sub modVersion
    GoGo.displayHelp "Module (" & module_name & ") version: " & moduleVersion(module_name)
End Sub

Sub testing
    msg = ""
    For i = 100 To 130
        msg = msg & "sc" & i & ":: MsgBox %A_ThisHotkey%" & vbNewLine
    Next
    GoGo.handleSelText(msg)
End Sub

' "Init" is required
' It is called automatically during initialization to create menu items.
Sub Init
    menuName = "&" & module_name
    addMenuItem "Find"                  , menuName, "FindPos"
    addMenuItem "&Replace"              , menuName, "replaceRegExp"
    addMenuItem "HTML &list items"      , menuName, "listItems"
    addMenuItem "Find To Dos"           , menuName, "FindToDos"
    addMenuItem "DEC2HEX"               , menuName, "ConvertDEC2HEX"
    addMenuItem "Toggle Value"          , menuName, "ToggleValue", "Ctrl+t"

    addMenuItem "-", menuName, ""
    addMenuItem "Paste to log"          , menuName, "paste2Log"
    addMenuItem "List open files"       , menuName, "GetOpenFiles"
    addMenuItem "Join open files"       , menuName, "JoinOpenFiles"

    addMenuItem "-", menuName, ""
    addMenuItem "Toggle Real &Tabs Settings" , menuName, "ToggleTabs"

    addMenuItem "-", menuName, ""
    addMenuItem "Maximize All"          , menuName, "MaxAll"
    addMenuItem "Minimize All"          , menuName, "MinAll"
    addMenuItem "Restore All"           , menuName, "RestoreAll"
    addMenuItem "Close All"             , menuName, "CloseAll"
    addMenuItem "Save All & Exit"       , menuName, "SaveAllExit", "Alt+F3"

    addMenuItem "-", menuName, ""
    addMenuItem "Testing"               , menuName, "testing"
    addMenuItem "Script module &version", menuName, "modVersion"
    addMenuItem "&Open script file"     , menuName, "openFile"

End Sub
