VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "節点番号選択"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    Me.Tag = "OK": Me.Hide
End Sub
Private Sub btnCancel_Click()
    Me.Tag = "Cancel": Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Tag = "Cancel": Me.Hide: Cancel = True
    End If
End Sub
