VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Encaminha eventos dos controles criados em tempo de execucao para o form.
'==============================================================================
Option Explicit

Public WithEvents Txt As MSForms.TextBox
Attribute Txt.VB_VarHelpID = -1
Public WithEvents Cmb As MSForms.ComboBox
Attribute Cmb.VB_VarHelpID = -1
Public WithEvents Chk As MSForms.CheckBox
Attribute Chk.VB_VarHelpID = -1
Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public WithEvents Scr As MSForms.ScrollBar
Attribute Scr.VB_VarHelpID = -1
Public Owner As Object
Public Tag As String

Private Sub Txt_Change()
    Owner.OnCampo Tag
End Sub

Private Sub Cmb_Change()
    Owner.OnCampo Tag
End Sub

Private Sub Chk_Click()
    Owner.OnCampo Tag
End Sub

Private Sub Btn_Click()
    Owner.OnBotao Tag
End Sub

Private Sub Scr_Change()
    Owner.OnScroll Tag
End Sub

Private Sub Scr_Scroll()
    Owner.OnScroll Tag
End Sub
