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
    lastRow = firstDataRow + ur.rows.Count - 1
    lastCol = firstDataCol + ur.Columns.Count - 1

    If ur.rows.Count < 2 Or ur.Columns.Count < 2 Then
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
    frm.Properties("Caption") = ChrW(31680) & ChrW(28857) & ChrW(30058) & ChrW(21495) & ChrW(36984) & ChrW(25246)
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
    cm.InsertLines n + 2, "    Me.Tag = ""OK"": Me.Hide"
    cm.InsertLines n + 3, "End Sub"
    cm.InsertLines n + 4, "Private Sub btnCancel_Click()"
    cm.InsertLines n + 5, "    Me.Tag = ""Cancel"": Me.Hide"
    cm.InsertLines n + 6, "End Sub"
    cm.InsertLines n + 7, "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    cm.InsertLines n + 8, "    If CloseMode = vbFormControlMenu Then"
    cm.InsertLines n + 9, "        Me.Tag = ""Cancel"": Me.Hide: Cancel = True"
    cm.InsertLines n + 10, "    End If"
    cm.InsertLines n + 11, "End Sub"

    MsgBox "FormNodeSelect created.", vbInformation
End Sub
