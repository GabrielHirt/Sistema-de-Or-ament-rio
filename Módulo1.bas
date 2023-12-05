Attribute VB_Name = "Módulo1"
Sub OpenForm()

sPrincipais = True
sTerceiros = False
sDiversos = False
UserForm2.addTerceiros = True
UserForm2.addDiversos = True


'=======================
'IRÁ DESABILITAR OS BOTÕES ABAIXO [TEMPORÁRIO]
UserForm2.s_diversos.Enabled = True
UserForm2.ex_BtnSelectionDiversos.Enabled = False
UserForm2.ex_s_diversos.Enabled = False

'=======================
'user2.Show

CarregarNivelUm

CarregarNivelDois

CarregarNivelTres

CarregarInsumos



UserForm2.txtBox_nv1_1.Value = ""
'UserForm2.TextBox9.Font.Italic = True

UserForm2.txtBox_nv2_2.Value = ""
'UserForm2.txtBox_nv2_2.Font.Italic = True

UserForm2.txtBox_nv3_3.Value = ""
'UserForm2.txtBox_nv3_3.Font.Italic = True

UserForm2.txtBox_nv4_4.Value = ""
'UserForm2.txtBox_nv4_4.Font.Italic = True


UserForm2.optionButton_nv1_2.Value = True
UserForm2.optionButton_nv2_3.Value = True
UserForm2.optionButton_nv3_6.Value = True
UserForm2.optionButton_nv4_8.Value = True



UserForm2.txtBox_un.BackColor = &H80000016
UserForm2.txtBox_un.Enabled = False

UserForm2.txtBox_rendimento.BackColor = &H80000016
UserForm2.txtBox_rendimento.Enabled = False

UserForm2.txtBox_custoInsumo.BackColor = &H80000016
UserForm2.txtBox_custoInsumo.Enabled = False

UserForm2.txtBox_pvs.BackColor = &H80000016
UserForm2.txtBox_pvs.Enabled = False

UserForm2.txtBox_cmo.BackColor = &H80000016
UserForm2.txtBox_cmo.Enabled = False

UserForm2.changeIndex

UserForm2.Show vbModeless



    With UserForm2.Background
        .Left = 0
        .Top = 0
        .Width = UserForm2.InsideWidth
        .Height = UserForm2.InsideHeight
    End With
    UserForm2.Background.BackColor = RGB(32, 55, 100)
End Sub



Sub CloseForm()

UserForm2.Hide
End Sub

Sub OpenForm2()

'user2.Show
UserForm1.Show

'UserForm2.Width = 900
'UserForm2.Height = 900

End Sub



