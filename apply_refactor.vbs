' apply_refactor.vbs
' Usage: cscript apply_refactor.vbs <full path to xlsm file>

Option Explicit

Dim xlApp, wb, vbProj, xlsmPath, scriptDir, cm, n

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript apply_refactor.vbs <xlsm-file-path>"
    WScript.Quit 1
End If

xlsmPath = WScript.Arguments(0)
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Open(xlsmPath)
Set vbProj = wb.VBProject

' Remove old modules
On Error Resume Next
vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
Err.Clear
vbProj.VBComponents.Remove vbProj.VBComponents("UserForm1")
Err.Clear
vbProj.VBComponents.Remove vbProj.VBComponents("UserForm2")
Err.Clear
On Error GoTo 0

' Import FormNodeSelect (frm + frx must be in test_260525_03\ subfolder)
vbProj.VBComponents.Import scriptDir & "test_260525_03\FormNodeSelect.frm"

' Remove Option Private Module from ParseSupportReaction
Set cm = vbProj.VBComponents("ParseSupportReaction").CodeModule
For n = 1 To cm.CountOfLines
    If Trim(cm.Lines(n, 1)) = "Option Private Module" Then
        cm.DeleteLines n, 1
        Exit For
    End If
Next

' Add wrapper subs to Module01
Set cm = vbProj.VBComponents("Module01").CodeModule
Dim all_code
all_code = cm.Lines(1, cm.CountOfLines)
If InStr(all_code, "RunParseAndSelectNodes") = 0 Then
    n = cm.CountOfLines
    cm.InsertLines n + 1, ""
    cm.InsertLines n + 2, "Public Sub RunParseAndSelectNodes()"
    cm.InsertLines n + 3, "    ParseSupportReaction.ParseAndSelectNodes"
    cm.InsertLines n + 4, "End Sub"
    cm.InsertLines n + 5, ""
    cm.InsertLines n + 6, "Public Sub RunFilterByNode()"
    cm.InsertLines n + 7, "    ParseSupportReaction.FilterByNode"
    cm.InsertLines n + 8, "End Sub"
End If

wb.Save
wb.Close
xlApp.Quit

WScript.Echo "Done."
