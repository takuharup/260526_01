Attribute VB_Name = "ParseSupportReaction"
Option Explicit

' 1�p�[�X�Ŏ��W����s�f�[�^
Private Type ReactionRow
    LoadNo   As Long
    loadName As String
    NodeNo   As String
    rx       As Double
    ry       As Double
    rz       As Double
    rmx      As Double
    rmy      As Double
    rmz      As Double
End Type

' -------------------------------------------------------
' �A���X�y�[�X�𐳋K������Split
' -------------------------------------------------------
Private Function SplitNormalized(ByVal s As String) As String()
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    SplitNormalized = Split(Trim(s), " ")
End Function

' -------------------------------------------------------
' �׏d�ԍ��𒊏o: "�׏d�ԍ� =  100" �̐��l����
' -------------------------------------------------------
Private Function ExtractLoadNumber(ByVal line As String) As Long
    Dim pos As Long
    Dim token As String
    pos = InStr(line, "�׏d�ԍ�")
    If pos = 0 Then
        ExtractLoadNumber = 0
        Exit Function
    End If
    ' "�׏d�ԍ� = " �ȍ~����׏d���� ��O�܂Ŏ��o��
    ' ���s�� "�׏d����" �����邽�߁A���̎�O�Ő؂���
    Dim sub1 As String
    Dim posName As Long
    posName = InStr(line, "�׏d����")
    If posName > 0 Then
        sub1 = Mid(line, pos, posName - pos)
    Else
        sub1 = Mid(line, pos)
    End If
    ' "=" �ȍ~��Trim
    Dim eqPos As Long
    eqPos = InStr(sub1, "=")
    If eqPos = 0 Then
        ExtractLoadNumber = 0
        Exit Function
    End If
    token = Trim(Mid(sub1, eqPos + 1))
    ' �擪�̐��l�g�[�N�������o��
    token = SplitNormalized(token)(0)
    If IsNumeric(token) Then
        ExtractLoadNumber = CLng(token)
    Else
        ExtractLoadNumber = 0
    End If
End Function

' -------------------------------------------------------
' �׏d���̂𒊏o: "�׏d���� =  �i������j" �̕����񕔕�
' -------------------------------------------------------
Private Function ExtractLoadName(ByVal line As String) As String
    Dim pos As Long
    pos = InStr(line, "�׏d����")
    If pos = 0 Then
        ExtractLoadName = ""
        Exit Function
    End If
    Dim sub1 As String
    sub1 = Mid(line, pos)
    Dim eqPos As Long
    eqPos = InStr(sub1, "=")
    If eqPos = 0 Then
        ExtractLoadName = ""
        Exit Function
    End If
    ExtractLoadName = Trim(Mid(sub1, eqPos + 1))
End Function

' -------------------------------------------------------
' �w��͈͂�parts�v�f���S��Numeric������
' -------------------------------------------------------
Private Function ValidateNumericParts(ByRef parts() As String, _
                                       ByVal fromIdx As Integer, _
                                       ByVal toIdx As Integer, _
                                       ByVal lineNum As Long) As Boolean
    Dim i As Integer
    For i = fromIdx To toIdx
        If Not IsNumeric(parts(i)) Then
            Debug.Print "Warning: ���l�ϊ����s parts(" & i & ")='" & parts(i) & "' (line " & lineNum & ")"
            ValidateNumericParts = False
            Exit Function
        End If
    Next i
    ValidateNumericParts = True
End Function

' -------------------------------------------------------
' �ߓ_�ԍ��t�B���^���o
' -------------------------------------------------------
Public Sub FilterByNode()

    Dim srcSheetName As String
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim dstSheetName As String

    Dim nodeInput As String
    Dim nodeTokens() As String
    Dim i As Integer

    Dim lastRow As Long
    Dim dstRow As Long
    Dim cellVal As String

    ' --- �X�e�b�v�@ ���e�[�u���V�[�g�I�� ---
    srcSheetName = InputBox("���o���̃V�[�g������͂��Ă�������" & vbCrLf & _
                            "�i��F�x�_����_20260525_143022�j", "FilterByNode")
    If StrPtr(srcSheetName) = 0 Then Exit Sub       ' �L�����Z��
    If Trim(srcSheetName) = "" Then Exit Sub        ' ����͂��T�C�����g�I��

    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(srcSheetName)
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "�V�[�g�u" & srcSheetName & "�v��������܂���B", vbCritical, "FilterByNode"
        Exit Sub
    End If

    ' --- �X�e�b�v�A �ߓ_�ԍ����X�g���� ---
    nodeInput = InputBox("���o����ߓ_�ԍ����J���}��؂�œ��͂��Ă�������" & vbCrLf & _
                         "�i��F 1,3,���v�@�܂��́@���v�j", "FilterByNode")
    If StrPtr(nodeInput) = 0 Then Exit Sub          ' �L�����Z��
    If Trim(nodeInput) = "" Then
        MsgBox "�ߓ_�ԍ������͂���Ă��܂���B", vbExclamation, "FilterByNode"
        Exit Sub
    End If

    ' �J���}Split���Ċe�v�f��Trim
    nodeTokens = Split(nodeInput, ",")
    For i = 0 To UBound(nodeTokens)
        nodeTokens(i) = Trim(nodeTokens(i))
    Next i

    ' --- �V�K�V�[�g�쐬 ---
    dstSheetName = "���o_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dstWs.Name = dstSheetName

    ' --- �w�b�_�[�R�s�[ ---
    dstWs.rows(1).Value = srcWs.rows(1).Value
    dstRow = 2

    ' --- �f�[�^�s���� ---
    lastRow = srcWs.Cells(srcWs.rows.Count, 1).End(xlUp).Row

    Dim extractCount As Long
    extractCount = 0

    For i = 2 To lastRow
        cellVal = Trim(CStr(srcWs.Cells(i, 3).Value))
        If IsInList(cellVal, nodeTokens) Then
            dstWs.rows(dstRow).Value = srcWs.rows(i).Value
            dstRow = dstRow + 1
            extractCount = extractCount + 1
        End If
    Next i

    ' --- �������b�Z�[�W ---
    If extractCount = 0 Then
        MsgBox "�Y������ߓ_�ԍ��̍s������܂���ł����B" & vbCrLf & _
               "�o�̓V�[�g��: " & dstSheetName, vbExclamation, "FilterByNode"
    Else
        MsgBox "�������܂����B" & vbCrLf & _
               "�o�̓V�[�g��: " & dstSheetName & vbCrLf & _
               "���o�s��: " & extractCount & " �s", vbInformation, "FilterByNode"
    End If

End Sub

' -------------------------------------------------------
' �ߓ_�ԍ����X�g��v����
' -------------------------------------------------------
Private Function IsInList(ByVal target As String, ByRef list() As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(list)
        If Trim(list(i)) = Trim(target) Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function

' -------------------------------------------------------
' txt��1�p�[�X���ăt�H�[���Őߓ_�I�� �� �V�K�V�[�g�֓]�L
' -------------------------------------------------------
Public Sub ParseAndSelectNodes()

    ' --- �t�@�C���I�� ---
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "�x�_���̓e�L�X�g�t�@�C����I�����Ă�������"
    fd.Filters.Clear
    fd.Filters.Add "�e�L�X�g�t�@�C��", "*.txt"
    fd.AllowMultiSelect = False
    If fd.Show <> True Then Exit Sub
    Dim filePath As String
    filePath = fd.SelectedItems(1)

    ' --- ��1�p�[�X�F�S�s���W & �ߓ_�ԍ����j�[�N���X�g ---
    Dim rows()    As ReactionRow
    Dim rowCount  As Long
    Dim nodeList() As String
    Dim nodeCount  As Long
    rowCount = 0
    nodeCount = 0

    Dim fileNum As Integer
    fileNum = FreeFile
    On Error GoTo FileOpenError2
    Open filePath For Input As #fileNum
    On Error GoTo 0

    Dim line As String
    Dim trimmedLine As String
    Dim parts() As String
    Dim loadNumber As Long
    Dim loadName As String
    Dim lineNum As Long
    Dim nodeStr As String
    Dim rx As Double, ry As Double, rz As Double
    Dim rmx As Double, rmy As Double, rmz As Double
    lineNum = 0
    loadNumber = 0
    loadName = ""

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lineNum = lineNum + 1
        trimmedLine = Trim(line)

        If Len(trimmedLine) = 0 Then GoTo Skip1
        If Left(trimmedLine, 5) = "=====" Then GoTo Skip1
        If InStr(line, "�׏d�ԍ�") > 0 Then
            loadNumber = ExtractLoadNumber(line)
            loadName = ExtractLoadName(line)
            GoTo Skip1
        End If
        If InStr(trimmedLine, "�ߓ_�ԍ�") > 0 Then GoTo Skip1

        If Left(trimmedLine, 2) = "���v" Then
            parts = SplitNormalized(trimmedLine)
            If UBound(parts) < 6 Then
                Debug.Print "Warning: ���v�s�t�B�[���h�s�� (line " & lineNum & ")"
                GoTo Skip1
            End If
            If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo Skip1
            nodeStr = "���v"
        ElseIf IsNumeric(Left(trimmedLine, InStr(trimmedLine & " ", " ") - 1)) Then
            parts = SplitNormalized(trimmedLine)
            If UBound(parts) < 6 Then
                Debug.Print "Warning: �ߓ_�s�t�B�[���h�s�� (line " & lineNum & ")"
                GoTo Skip1
            End If
            If Not IsNumeric(parts(0)) Then GoTo Skip1
            If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo Skip1
            nodeStr = parts(0)
        Else
            GoTo Skip1
        End If

        rx = CDbl(parts(1)): ry = CDbl(parts(2)): rz = CDbl(parts(3))
        rmx = CDbl(parts(4)): rmy = CDbl(parts(5)): rmz = CDbl(parts(6))

        ReDim Preserve rows(rowCount)
        rows(rowCount).LoadNo = loadNumber
        rows(rowCount).loadName = loadName
        rows(rowCount).NodeNo = nodeStr
        rows(rowCount).rx = rx: rows(rowCount).ry = ry: rows(rowCount).rz = rz
        rows(rowCount).rmx = rmx: rows(rowCount).rmy = rmy: rows(rowCount).rmz = rmz
        rowCount = rowCount + 1

        ' �o�������j�[�N���X�g�֒ǉ�
        Dim alreadyExists As Boolean
        alreadyExists = False
        If nodeCount > 0 Then alreadyExists = IsInList(nodeStr, nodeList)
        If Not alreadyExists Then
            ReDim Preserve nodeList(nodeCount)
            nodeList(nodeCount) = nodeStr
            nodeCount = nodeCount + 1
        End If
Skip1:
    Loop
    Close #fileNum

    If rowCount = 0 Then
        MsgBox "�f�[�^�s��������܂���ł����B", vbExclamation, "ParseAndSelectNodes"
        Exit Sub
    End If

    ' --- FormNodeSelect �Őߓ_�I�� ---
    Dim frm As FormNodeSelect
    Set frm = New FormNodeSelect
    Dim j As Integer
    For j = 0 To nodeCount - 1
        frm.lstNodes.AddItem nodeList(j)
    Next j
    frm.Show

    If frm.Tag <> "OK" Then
        Unload frm
        Exit Sub
    End If

    Dim selectedNodes() As String
    Dim selCount As Integer
    selCount = 0
    For j = 0 To frm.lstNodes.ListCount - 1
        If frm.lstNodes.Selected(j) Then
            ReDim Preserve selectedNodes(selCount)
            selectedNodes(selCount) = frm.lstNodes.list(j)
            selCount = selCount + 1
        End If
    Next j
    Unload frm

    If selCount = 0 Then
        MsgBox "�ߓ_�ԍ����I������Ă��܂���B", vbExclamation, "ParseAndSelectNodes"
        Exit Sub
    End If

    ' --- �V�K�V�[�g�쐬 & �w�b�_�[ ---
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "�x�_����_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

    Dim headers As Variant
    headers = Array("�׏d�ԍ�", "�׏d����", "�ߓ_�ԍ�", "RX", "RY", "RZ", "RMX", "RMY", "RMZ")
    Dim c As Integer
    For c = 0 To 8
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    ' --- �]�L ---
    Dim rowIdx As Long
    rowIdx = 2
    Dim k As Long
    For k = 0 To rowCount - 1
        If IsInList(rows(k).NodeNo, selectedNodes) Then
            ws.Cells(rowIdx, 1).Value = rows(k).LoadNo
            ws.Cells(rowIdx, 2).Value = rows(k).loadName
            If IsNumeric(rows(k).NodeNo) Then
                ws.Cells(rowIdx, 3).Value = CLng(rows(k).NodeNo)
            Else
                ws.Cells(rowIdx, 3).Value = rows(k).NodeNo
            End If
            ws.Cells(rowIdx, 4).Value = rows(k).rx
            ws.Cells(rowIdx, 5).Value = rows(k).ry
            ws.Cells(rowIdx, 6).Value = rows(k).rz
            ws.Cells(rowIdx, 7).Value = rows(k).rmx
            ws.Cells(rowIdx, 8).Value = rows(k).rmy
            ws.Cells(rowIdx, 9).Value = rows(k).rmz
            rowIdx = rowIdx + 1
        End If
    Next k

    MsgBox "�������܂����B" & vbCrLf & _
           "�o�̓V�[�g��: " & sheetName & vbCrLf & _
           "�]�L�s��: " & (rowIdx - 2) & " �s", vbInformation, "ParseAndSelectNodes"
    Exit Sub

FileOpenError2:
    MsgBox "�t�@�C�����J���܂���ł����B" & vbCrLf & Err.Description, vbCritical, "ParseAndSelectNodes"

End Sub

' -------------------------------------------------------
' 1�s�����V�[�g�֏�������
' -------------------------------------------------------
Private Sub WriteRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                     ByVal loadNum As Long, ByVal loadNm As String, _
                     ByVal nodeVal As String, _
                     ByVal rx As Double, ByVal ry As Double, ByVal rz As Double, _
                     ByVal rmx As Double, ByVal rmy As Double, ByVal rmz As Double)
    ws.Cells(rowIdx, 1).Value = loadNum
    ws.Cells(rowIdx, 2).Value = loadNm
    ' �ߓ_�ԍ������l������Ȃ� Long�A"���v"�Ȃ炻�̂܂ܕ�����
    If IsNumeric(nodeVal) Then
        ws.Cells(rowIdx, 3).Value = CLng(nodeVal)
    Else
        ws.Cells(rowIdx, 3).Value = nodeVal
    End If
    ws.Cells(rowIdx, 4).Value = rx
    ws.Cells(rowIdx, 5).Value = ry
    ws.Cells(rowIdx, 6).Value = rz
    ws.Cells(rowIdx, 7).Value = rmx
    ws.Cells(rowIdx, 8).Value = rmy
    ws.Cells(rowIdx, 9).Value = rmz
End Sub


