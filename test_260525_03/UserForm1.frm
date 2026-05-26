VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Width = 450
    Me.Height = 370
    cmbTemplate.AddItem "STRUDL dat"
    cmbTemplate.ListIndex = 0
    Call RefreshPreview
End Sub

Private Sub RefreshPreview()
    Dim content As String
    content = GenerateCombText()
    If content = "" Then
        txtPreview.Value = "(データが見つかりません)"
    Else
        txtPreview.Value = content
    End If
End Sub

Private Sub btnPreview_Click()
    Call RefreshPreview
End Sub

Private Sub btnCreate_Click()
    Dim content As String
    content = GenerateCombText()
    If content = "" Then
        MsgBox "データが見つかりません。Sheet1のデータを確認してください。", vbExclamation
        Exit Sub
    End If
    Call SaveDatFile(content)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

