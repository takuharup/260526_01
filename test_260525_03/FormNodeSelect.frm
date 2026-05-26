Attribute VB_Name = "FormNodeSelect"
Attribute VB_Base = "0{21E84600-F260-4A52-9737-DBDA07164C87}{36A1D1B0-7831-4F84-B3D8-71056EE40FB4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
