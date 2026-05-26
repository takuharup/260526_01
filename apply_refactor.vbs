'+------------------------------------------------------------------
' apply_refactor.vbs
'
' Usage:
'   cscript apply_refactor.vbs
'
' Place this script in the same folder as test_260525_03.xlsm
' and the test_260525_03\ subfolder.
'
' Requirement:
'   Excel > Options > Trust Center > Trust Center Settings
'   > Macro Settings > check "Trust access to the VBA project object model"
'+------------------------------------------------------------------
Option Explicit

Dim fso, script_dir, xlsm_path, src_dir
Set fso = CreateObject("Scripting.FileSystemObject")
script_dir = fso.GetParentFolderName(WScript.ScriptFullName)
xlsm_path  = fso.BuildPath(script_dir, "test_260525_03.xlsm")
src_dir    = fso.BuildPath(script_dir, "test_260525_03")

If Not fso.FileExists(xlsm_path) Then
    WScript.Echo "ERROR: xlsm not found: " & xlsm_path
    WScript.Quit 1
End If

WScript.Echo "Opening Excel..."
Dim xl
Set xl = CreateObject("Excel.Application")
xl.Visible       = False
xl.DisplayAlerts = False

WScript.Echo "Opening: " & xlsm_path
Dim wb
Set wb = xl.Workbooks.Open(xlsm_path)

Dim vbp
On Error Resume Next
Set vbp = wb.VBProject
On Error GoTo 0

If vbp Is Nothing Then
    WScript.Echo "ERROR: Cannot access VBA project."
    WScript.Echo "  Enable [Trust access to the VBA project object model] in Excel Trust Center."
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

'--- Step 1: Remove old modules ---
Dim del_list(2)
del_list(0) = "Module1"
del_list(1) = "UserForm1"
del_list(2) = "UserForm2"

Dim i, comp
For i = 0 To 2
    Set comp = Nothing
    On Error Resume Next
    Set comp = vbp.VBComponents(del_list(i))
    On Error GoTo 0
    If Not comp Is Nothing Then
        vbp.VBComponents.Remove comp
        WScript.Echo "  Removed: " & del_list(i)
    End If
Next

'--- Step 2: Import FormNodeSelect (frm + frx must be in same folder) ---
Dim frm_path
frm_path = fso.BuildPath(src_dir, "FormNodeSelect.frm")
If fso.FileExists(frm_path) Then
    vbp.VBComponents.Import frm_path
    WScript.Echo "  Imported: FormNodeSelect"
Else
    WScript.Echo "  WARNING: FormNodeSelect.frm not found: " & frm_path
End If

'--- Step 3: Remove Option Private Module from ParseSupportReaction ---
Set comp = Nothing
On Error Resume Next
Set comp = vbp.VBComponents("ParseSupportReaction")
On Error GoTo 0

If Not comp Is Nothing Then
    Dim cm, n, line_text
    Set cm = comp.CodeModule
    For n = 1 To cm.CountOfLines
        line_text = cm.Lines(n, 1)
        If Trim(line_text) = "Option Private Module" Then
            cm.DeleteLines n, 1
            WScript.Echo "  Removed: Option Private Module from ParseSupportReaction"
            Exit For
        End If
    Next
End If

'--- Step 4: Add wrapper subs to Module01 ---
Set comp = Nothing
On Error Resume Next
Set comp = vbp.VBComponents("Module01")
On Error GoTo 0

If Not comp Is Nothing Then
    Set cm = comp.CodeModule
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
        WScript.Echo "  Added: wrapper subs to Module01"
    Else
        WScript.Echo "  Skipped: wrapper subs already exist in Module01"
    End If
End If

'--- Save and close ---
wb.Save
wb.Close False
xl.Quit
Set xl = Nothing

WScript.Echo ""
WScript.Echo "Done: " & xlsm_path
