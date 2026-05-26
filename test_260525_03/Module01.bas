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
        FileFilter:="DAT�t�@�C�� (*.dat), *.dat", _
        Title:="�ۑ����I�����Ă�������")

    If savePath = False Then Exit Sub

    fileNum = FreeFile
    Open CStr(savePath) For Output As #fileNum
    Print #fileNum, content
    Close #fileNum

    MsgBox "�ۑ����܂���: " & CStr(savePath), vbInformation
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
