Attribute VB_Name = "CopyToOpenBook"
Option Explicit

Public Sub CopyRangeToOpenBook()

    ' マクロ起動時点のアクティブシートを貼り付け先として確保
    Dim destSheet As Worksheet
    Set destSheet = ThisWorkbook.ActiveSheet

    ' STEP 1: コピー元ファイル選択
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="コピー元ファイルを選択してください")
    If srcPath = False Then Exit Sub

    Dim srcWB As Workbook
    Set srcWB = Workbooks.Open(CStr(srcPath))

    ' STEP 2: コピー元範囲をマウスで選択（Ctrl+クリックで複数選択可）
    Dim srcRange As Range
    On Error GoTo CancelDrag
    Set srcRange = Application.InputBox( _
        Prompt:="コピー元の範囲をドラッグで選択してください" & vbCrLf & _
                "（Ctrl+クリックで複数範囲を選択できます）", _
        Title:="コピー元範囲の選択", _
        Type:=8)
    On Error GoTo 0

    ' STEP 3: エリアごとに順番にA1から横に貼り付け
    Dim pasteCol As Long
    pasteCol = 1

    Dim area As Range
    For Each area In srcRange.Areas
        area.Copy
        destSheet.Cells(1, pasteCol).PasteSpecial Paste:=xlPasteValues
        destSheet.Cells(1, pasteCol).PasteSpecial Paste:=xlPasteFormats
        pasteCol = pasteCol + area.Columns.Count
    Next area
    Application.CutCopyMode = False

    srcWB.Close False

    MsgBox "コピーが完了しました。", vbInformation
    Exit Sub

CancelDrag:
    On Error GoTo 0
    srcWB.Close False
End Sub
