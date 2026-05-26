'+------------------------------------------------------------------
' apply_refactor.vbs
'
' 使い方:
'   cscript apply_refactor.vbs
'   （スクリプトと test_260525_03.xlsm を同じフォルダに置いて実行）
'
' 事前設定:
'   Excel > ファイル > オプション > セキュリティセンター
'   > セキュリティセンターの設定 > マクロの設定
'   > [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] にチェック
'+------------------------------------------------------------------
Option Explicit

Dim fso, script_dir, xlsm_path, src_dir
Set fso = CreateObject("Scripting.FileSystemObject")
script_dir = fso.GetParentFolderName(WScript.ScriptFullName)
xlsm_path  = fso.BuildPath(script_dir, "test_260525_03.xlsm")
src_dir    = fso.BuildPath(script_dir, "test_260525_03")

If Not fso.FileExists(xlsm_path) Then
    WScript.Echo "ERROR: xlsm が見つかりません: " & xlsm_path
    WScript.Quit 1
End If

WScript.Echo "Excel を起動しています..."
Dim xl
Set xl = CreateObject("Excel.Application")
xl.Visible        = False
xl.DisplayAlerts  = False

WScript.Echo "ファイルを開いています: " & xlsm_path
Dim wb
Set wb = xl.Workbooks.Open(xlsm_path)

Dim vbp
On Error Resume Next
Set vbp = wb.VBProject
On Error GoTo 0

If vbp Is Nothing Then
    WScript.Echo "ERROR: VBA プロジェクトにアクセスできません。"
    WScript.Echo "       Excel のセキュリティセンターで"
    WScript.Echo "       [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する]"
    WScript.Echo "       を有効にしてから再実行してください。"
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

' ----------------------------------------------------------------
' Step 1: 不要モジュール削除
' ----------------------------------------------------------------
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
        WScript.Echo "  削除: " & del_list(i)
    End If
Next

' ----------------------------------------------------------------
' Step 2: FormNodeSelect をインポート（frm + frx をセットで）
'   ※ frx は frm と同じフォルダに同名で置いておく必要あり
' ----------------------------------------------------------------
Dim frm_path
frm_path = fso.BuildPath(src_dir, "FormNodeSelect.frm")
If fso.FileExists(frm_path) Then
    vbp.VBComponents.Import frm_path
    WScript.Echo "  インポート: FormNodeSelect"
Else
    WScript.Echo "  WARNING: FormNodeSelect.frm が見つかりません: " & frm_path
End If

' ----------------------------------------------------------------
' Step 3: ParseSupportReaction から Option Private Module を削除
' ----------------------------------------------------------------
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
            WScript.Echo "  削除: ParseSupportReaction の Option Private Module"
            Exit For
        End If
    Next
End If

' ----------------------------------------------------------------
' Step 4: Module01 にラッパーサブを追加
' ----------------------------------------------------------------
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
        WScript.Echo "  追加: Module01 にラッパーサブ"
    Else
        WScript.Echo "  スキップ: Module01 のラッパーサブは既に存在"
    End If
End If

' ----------------------------------------------------------------
' 保存して終了
' ----------------------------------------------------------------
wb.Save
wb.Close False
xl.Quit
Set xl = Nothing

WScript.Echo ""
WScript.Echo "完了: " & xlsm_path
