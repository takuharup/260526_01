' apply_refactor.vbs
' Usage: cscript apply_refactor.vbs <full path to xlsm file>

Option Explicit

Dim xlApp, wb, vbProj, xlsmPath, scriptDir, srcDir

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript apply_refactor.vbs <xlsm-file-path>"
    WScript.Quit 1
End If

xlsmPath = WScript.Arguments(0)
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
srcDir = scriptDir & "test_260525_03\"

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Open(xlsmPath)
Set vbProj = wb.VBProject

' Remove old/modified modules
On Error Resume Next
vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
Err.Clear
vbProj.VBComponents.Remove vbProj.VBComponents("UserForm1")
Err.Clear
vbProj.VBComponents.Remove vbProj.VBComponents("UserForm2")
Err.Clear
vbProj.VBComponents.Remove vbProj.VBComponents("ParseSupportReaction")
Err.Clear
vbProj.VBComponents.Remove vbProj.VBComponents("Module01")
Err.Clear
On Error GoTo 0

' Import all modules from source files
vbProj.VBComponents.Import srcDir & "FormNodeSelect.frm"
vbProj.VBComponents.Import srcDir & "ParseSupportReaction.bas"
vbProj.VBComponents.Import srcDir & "Module01.bas"

wb.Save
wb.Close
xlApp.Quit

WScript.Echo "Done."
