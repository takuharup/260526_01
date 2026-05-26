Attribute VB_Name = "Module01"
Option Explicit

Public Function FormatValue(v As Double) As String
    FormatValue = Format(v, "0.##########")
End Function

Public Function GenerateCombText() As String
    Dim ws As Worksheet
    Dim ur As Range
    Dim lastRow As Long, lastCol As Long
    Dim firstDataRow As Long, firstDataCol As Long
    Dim i As Long, j As Long
    Dim combNum As Long, loadNum As Long
    Dim cellVal As Variant
    Dim loadCell As Variant
    Dim line As String
    Dim result As String
    Dim hasData As Boolean

    'Set ws = ThisWorkbook.Sheets("Sheet1")
    Set ws = ThisWorkbook.ActiveSheet
    Set ur = ws.UsedRange

    firstDataRow = ur.Row
    firstDataCol = ur.Column
    lastRow = firstDataRow + ur.Rows.Count - 1
    lastCol = firstDataCol + ur.Columns.Count - 1

    If ur.Rows.Count < 2 Or ur.Columns.Count < 2 Then
        GenerateCombText = ""
        Exit Function
    End If

    result = ""

    For j = firstDataCol + 1 To lastCol
        cellVal = ws.Cells(firstDataRow, j).Value
        If Not IsEmpty(cellVal) And cellVal <> "" And IsNumeric(cellVal) Then
            combNum = CLng(cellVal)
            line = "Load comb " & combNum & " comb"
            hasData = False

            For i = firstDataRow + 1 To lastRow
                loadCell = ws.Cells(i, firstDataCol).Value
                If Not IsEmpty(loadCell) And loadCell <> "" And IsNumeric(loadCell) Then
                    loadNum = CLng(loadCell)
                    cellVal = ws.Cells(i, j).Value
                    If Not IsEmpty(cellVal) And IsNumeric(cellVal) Then
                        If CDbl(cellVal) <> 0 Then
                            line = line & " " & loadNum & " " & FormatValue(CDbl(cellVal))
                            hasData = True
                        End If
                    End If
                End If
            Next i

            If hasData Then
                If result <> "" Then result = result & vbCrLf
                result = result & line
            End If
        End If
    Next j

    GenerateCombText = result
End Function

Public Sub SaveDatFile(content As String)
    Dim defaultName As String
    Dim defaultPath As String
    Dim savePath As Variant
    Dim fileNum As Integer

    defaultName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_comb.dat"
    defaultPath = ThisWorkbook.Path & "\" & defaultName

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultPath, _
        FileFilter:="DATファイル (*.dat), *.dat", _
        Title:="保存先を選択してください")

    If savePath = False Then Exit Sub

    fileNum = FreeFile
    Open CStr(savePath) For Output As #fileNum
    Print #fileNum, content
    Close #fileNum

    MsgBox "保存しました: " & CStr(savePath), vbInformation
End Sub

Public Sub ShowForm()
    UserForm01.Show
End Sub

Public Sub RunParseAndSelectNodes()
    ParseSupportReaction.ParseAndSelectNodes
End Sub

Public Sub RunFilterByNode()
    ParseSupportReaction.FilterByNode
End Sub

Public Sub CreateFormNodeSelect()
    Dim vbp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim cm As Object
    Dim n As Long

    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("FormNodeSelect")
    On Error GoTo 0

    Set frm = vbp.VBComponents.Add(3)
    frm.Name = "FormNodeSelect"
    frm.Properties("Caption") = "節点番号選択"
    frm.Properties("Width") = 270
    frm.Properties("Height") = 360

    With frm.Designer
        Set ctrl = .Controls.Add("Forms.ListBox.1", "lstNodes")
        ctrl.Left = 6
        ctrl.Top = 6
        ctrl.Width = 258
        ctrl.Height = 300
        ctrl.MultiSelect = 1

        Set ctrl = .Controls.Add("Forms.CommandButton.1", "btnOK")
        ctrl.Caption = "OK"
        ctrl.Left = 48
        ctrl.Top = 318
        ctrl.Width = 72
        ctrl.Height = 24

        Set ctrl = .Controls.Add("Forms.CommandButton.1", "btnCancel")
        ctrl.Caption = "Cancel"
        ctrl.Left = 150
        ctrl.Top = 318
        ctrl.Width = 72
        ctrl.Height = 24
    End With

    Set cm = frm.CodeModule
    n = cm.CountOfLines
    cm.InsertLines n + 1, "Private Sub btnOK_Click()"
    cm.InsertLines n + 2, "    Me.Tag = " & Chr(34) & "OK" & Chr(34) & ": Me.Hide"
    cm.InsertLines n + 3, "End Sub"
    cm.InsertLines n + 4, "Private Sub btnCancel_Click()"
    cm.InsertLines n + 5, "    Me.Tag = " & Chr(34) & "Cancel" & Chr(34) & ": Me.Hide"
    cm.InsertLines n + 6, "End Sub"
    cm.InsertLines n + 7, "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    cm.InsertLines n + 8, "    If CloseMode = vbFormControlMenu Then"
    cm.InsertLines n + 9, "        Me.Tag = " & Chr(34) & "Cancel" & Chr(34) & ": Me.Hide: Cancel = True"
    cm.InsertLines n + 10, "    End If"
    cm.InsertLines n + 11, "End Sub"

    MsgBox "FormNodeSelect created.", vbInformation
End Sub

Public Sub CreateUserForm01()
    Dim vbp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim cm As Object
    Dim n As Long
    Dim q As String
    q = Chr(34)

    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("UserForm01")
    On Error GoTo 0

    Set frm = vbp.VBComponents.Add(3)
    frm.Name = "UserForm01"
    frm.Properties("Caption") = "UserForm1"
    frm.Properties("Width") = 450
    frm.Properties("Height") = 370

    With frm.Designer
        Set ctrl = .Controls.Add("Forms.ComboBox.1", "cmbTemplate")
        ctrl.Left = 12
        ctrl.Top = 12
        ctrl.Width = 180
        ctrl.Height = 18
        ctrl.Style = 2

        Set ctrl = .Controls.Add("Forms.TextBox.1", "txtPreview")
        ctrl.Left = 12
        ctrl.Top = 42
        ctrl.Width = 420
        ctrl.Height = 240
        ctrl.MultiLine = True
        ctrl.ScrollBars = 2

        Set ctrl = .Controls.Add("Forms.CommandButton.1", "btnPreview")
        ctrl.Caption = "Preview"
        ctrl.Left = 60
        ctrl.Top = 300
        ctrl.Width = 90
        ctrl.Height = 24

        Set ctrl = .Controls.Add("Forms.CommandButton.1", "btnCreate")
        ctrl.Caption = "Create"
        ctrl.Left = 180
        ctrl.Top = 300
        ctrl.Width = 90
        ctrl.Height = 24

        Set ctrl = .Controls.Add("Forms.CommandButton.1", "btnCancel")
        ctrl.Caption = "Cancel"
        ctrl.Left = 300
        ctrl.Top = 300
        ctrl.Width = 90
        ctrl.Height = 24
    End With

    Set cm = frm.CodeModule
    n = cm.CountOfLines
    cm.InsertLines n + 1,  "Option Explicit"
    cm.InsertLines n + 2,  ""
    cm.InsertLines n + 3,  "Private Sub UserForm_Initialize()"
    cm.InsertLines n + 4,  "    Me.Width = 450"
    cm.InsertLines n + 5,  "    Me.Height = 370"
    cm.InsertLines n + 6,  "    cmbTemplate.AddItem " & q & "STRUDL dat" & q
    cm.InsertLines n + 7,  "    cmbTemplate.ListIndex = 0"
    cm.InsertLines n + 8,  "    Call RefreshPreview"
    cm.InsertLines n + 9,  "End Sub"
    cm.InsertLines n + 10, ""
    cm.InsertLines n + 11, "Private Sub RefreshPreview()"
    cm.InsertLines n + 12, "    Dim content As String"
    cm.InsertLines n + 13, "    content = GenerateCombText()"
    cm.InsertLines n + 14, "    If content = " & q & q & " Then"
    cm.InsertLines n + 15, "        txtPreview.Value = " & q & "(データが見つかりません)" & q
    cm.InsertLines n + 16, "    Else"
    cm.InsertLines n + 17, "        txtPreview.Value = content"
    cm.InsertLines n + 18, "    End If"
    cm.InsertLines n + 19, "End Sub"
    cm.InsertLines n + 20, ""
    cm.InsertLines n + 21, "Private Sub btnPreview_Click()"
    cm.InsertLines n + 22, "    Call RefreshPreview"
    cm.InsertLines n + 23, "End Sub"
    cm.InsertLines n + 24, ""
    cm.InsertLines n + 25, "Private Sub btnCreate_Click()"
    cm.InsertLines n + 26, "    Dim content As String"
    cm.InsertLines n + 27, "    content = GenerateCombText()"
    cm.InsertLines n + 28, "    If content = " & q & q & " Then"
    cm.InsertLines n + 29, "        MsgBox " & q & "データが見つかりません。Sheet1のデータを確認してください。" & q & ", vbExclamation"
    cm.InsertLines n + 30, "        Exit Sub"
    cm.InsertLines n + 31, "    End If"
    cm.InsertLines n + 32, "    Call SaveDatFile(content)"
    cm.InsertLines n + 33, "End Sub"
    cm.InsertLines n + 34, ""
    cm.InsertLines n + 35, "Private Sub btnCancel_Click()"
    cm.InsertLines n + 36, "    Unload Me"
    cm.InsertLines n + 37, "End Sub"

    MsgBox "UserForm01 created.", vbInformation
End Sub
