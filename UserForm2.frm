VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   ClientHeight    =   6360
   ClientLeft      =   -15020
   ClientTop       =   -61710
   ClientWidth     =   15390
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_nv1 As Integer
Public id_nv2 As Integer
Public id_nv3 As Integer
Public Insumo_nv4 As String
Public id_nv4 As Integer
Public rendimentoValue As Boolean
Public pvs As Double
Public cmo As Double
Public newInsumo As Boolean
Public insumoBolean As Boolean
Public PvsUpdate As Double
'Define a lista a ser carregada p/ cada serviço
Public sPrincipais As Boolean
'Define a reinicialização após um envio do forms, qual lista irá carregar de acordo com o tipo de serviço
Public addPrincipal As Boolean
Public sDiversos As Boolean
Public addDiversos As Boolean
Public sTerceiros As Boolean
Public addTerceiros As Boolean
'Define o layout de edição
Public edicao As Boolean
'Define o layout de adição
Public adicao As Boolean
Public remocao As Boolean
Public c_pvs As Boolean
Public c_cmo As Boolean
Public c_rendimento As Boolean

Public ed_btnCmoPvsBoolean As Boolean
Public ed_btnNv4Boolean As Boolean
Public ed_btnRendimentoBoolean As Boolean
Public ed_btnNv1Boolean As Boolean
Public ed_btnNv2Boolean As Boolean
Public ed_btnNv3Boolean As Boolean
Public Visible As Boolean

Public listaNv1 As Boolean
Public txtBoxNv1 As Boolean
Public listaNv2 As Boolean
Public txtBoxNv2 As Boolean
Public listaNv3 As Boolean
Public txtBoxNv3 As Boolean
Public listaNv4 As Boolean
Public txtBoxNv4 As Boolean


Public nv3TopTxt As Long
Public nv3LeftTxt As Long
Public nv3TopCombo As Long
Public nv3LeftCombo As Long

Public nv2TopTxt As Long
Public nv2LeftTxt As Long
Public nv2TopCombo As Long
Public nv2LeftCombo As Long

Public nv2TitleTop As Long
Public nv2TitleLeft As Long

Public nv3TitleTop As Long
Public nv3TitleLeft As Long

Public nv1LeftTxt As Long
Public nv1TopTxt As Long

Public nv1LeftCombo As Long
Public nv1TopCombo As Long

Public nv1TitleTop As Long
Public nv1TitleLeft As Long

Public ed_title_nv4Left As Long
Public ed_txtComboBox_nv4_4Left As Long
Public ed_txtBox_nv4_4Left As Long
Public ed_descripUnLeft As Long
Public ed_txtBox_unLeft As Long
Public ed_descripCustInsumoLeft As Long
Public ed_txtBox_custoInsumoLeft As Long

Public ed_title_nv4Top As Long
Public ed_txtComboBox_nv4_4Top As Long
Public ed_txtBox_nv4_4Top As Long
Public ed_descripUnTop As Long
Public ed_txtBox_unTop As Long
Public ed_descripCustInsumoTop As Long
Public ed_txtBox_custoInsumoTop As Long



Public ed_descripCmoTop As Long
Public ed_descripCmoLeft As Long

Public ed_txtBox_cmoTop As Long
Public ed_txtBox_cmoLeft As Long

Public ed_descripRendTop As Long
Public ed_descripRendLeft As Long

Public ed_txtBox_rendimentoTop As Long
Public ed_txtBox_rendimentoLeft As Long


Public ed_descripPvsTop As Long
Public ed_descripPvsLeft As Long
    
Public ed_txtBox_pvsTop  As Long
Public ed_txtBox_pvsLeft As Long



'==== Variáveis do Menu de Exclusão
Public ex_sPrincipais As Boolean
Public ex_sDiversos As Boolean
Public ex_sTerceiros As Boolean
Public ex_sGeneralInsumo As Boolean


Public ex_btnNv1BooleanP As Boolean
Public ex_btnNv2BooleanP As Boolean
Public ex_btnNv3BooleanP As Boolean
Public ex_btnNv4BooleanP As Boolean
Public ex_btnNv3BooleanD As Boolean
Public ex_btnEstruturaBoolean As Boolean

Public ex_BtnSelectionPrincipaisBoolean As Boolean
Public ex_BtnSelectionDiversosBoolean As Boolean
Public menuEstrutura As Boolean
'Sem função definida, usar quando preciso
Public var_global As Integer


'====== Variáveis Log =======
'ADIÇÃO

Public log_ad_txtBox_cmo As Long
Public log_ad_txtBox_pvs As Long
Public log_ad_txtBox_custoInsumo As Long
Public log_ad_txtBox_rendimento As Long
Public log_ad_txtBox_un As String
Public log_ad_flagBox As Boolean


Public log_ad_txtBox_nv1_1 As String
Public log_ad_txtComboBox_nv1_1 As String

Public log_ad_txtBox_nv2_2 As String
Public log_ad_txtComboBox_nv2_2 As String

Public log_ad_txtBox_nv3_3 As String
Public log_ad_txtComboBox_nv3_3 As String

Public log_ad_txtBox_nv4_4 As String
Public log_ad_txtComboBox_nv4_4 As String

'============================



'Public ItemAtivo As Boolean


Sub atualizarPvsCmo()



GetAdicionais

End Sub

Sub GetPvs_Cmo()


Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
'Dim id_master As Long
Dim id_master As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
Dim insumo As String
Dim selectedRow As Integer
Dim coluna2Valor As String


Me.GetIdNv1
Me.GetIdNv2
Me.GetIdNv3
'Me.GetRendimento1

id_nv1 = txtBoxID_nv1.Value
id_nv2 = txtBoxID_nv2.Value
If txtBoxID_nv3.Value = "" Then
    txtBoxID_nv3.Value = 0
    id_nv3 = txtBoxID_nv3.Value

Else
id_nv3 = txtBoxID_nv3.Value
End If

sql1 = "t_Servicos_Principais"


If sDiversos = True Then
'id_master = 7 & 0 & id_nv3
id_master = 7 & "-" & 0 & "-" & id_nv3
id_nv1 = 7
id_nv2 = 0
Else
'id_master = id_nv1 & id_nv2 & id_nv3
id_master = id_nv1 & "-" & id_nv2 & "-" & id_nv3
End If

    i = 0
    While i <= 1
    ConectarBanco conexao
    If sDiversos = True Then
        If i = 0 Then
            sql1 = "SELECT precoVendaSugerido FROM t_Servicos_Diversos WHERE idNv1 = " & id_nv1 & " AND idNv2 = " & id_nv2 & " AND idNv3 = " & id_nv3
        Else
            sql1 = "SELECT CustoMaoObra FROM t_Servicos_Diversos WHERE idNv1 = " & id_nv1 & " AND idNv2 = " & id_nv2 & " AND idNv3 = " & id_nv3
        End If
    Else
        If i = 0 Then
            sql1 = "SELECT precoVendaSugerido FROM t_Servicos_Principais WHERE idNv1 = " & id_nv1 & " AND idNv2 = " & id_nv2 & " AND idNv3 = " & id_nv3
        Else
            sql1 = "SELECT CustoMaoObra FROM t_Servicos_Principais WHERE idNv1 = " & id_nv1 & " AND idNv2 = " & id_nv2 & " AND idNv3 = " & id_nv3
        End If
    End If
    
    



    rs.Open sql1, conexao

    If sql1 <> "" Then
        If i = 0 Then
        precoVendaSugerido = rs.Fields("precoVendaSugerido").Value
        pvs = precoVendaSugerido
        Me.txtBox_pvs.Value = Format(pvs, "R$ #,##0.00")
        c_pvs = False
        Else
        CustoMaoObra = rs.Fields("CustoMaoObra").Value
        cmo = CustoMaoObra
        c_cmo = False
        Me.txtBox_cmo.Value = Format(cmo, "R$ #,##0.00")
        End If
        On Error GoTo here


   
    
    
    
    
    Else

        If i = 0 Then
here:
            MsgBox "Não há preço de venda surgerido, crie um novo"
            c_pvs = True
        Else
            MsgBox "Não há preço de custo para mão de obra, crie um novo"
            c_cmo = True
        End If

    End If
    
    rs.Close
    conexao.Close

    i = i + 1
Wend

End Sub

Sub GetAdicionais()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
'Dim id_master As Long
Dim id_master As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
Dim insumo As String
Dim selectedRow As Integer
Dim coluna2Valor As String
ConectarBanco conexao

Me.GetIdNv1
Me.GetIdNv2
Me.GetIdNv3


id_nv1 = txtBoxID_nv1.Value
id_nv2 = txtBoxID_nv2.Value
If txtBoxID_nv3.Value = "" Then
    txtBoxID_nv3.Value = 0
    id_nv3 = txtBoxID_nv3.Value

Else
id_nv3 = txtBoxID_nv3.Value
End If

sql1 = "t_Servicos_Principais_Insumos"

If sDiversos = True Then
'id_master = 7 & 0 & id_nv3
id_master = 7 & "-" & 0 & "-" & id_nv3
Else
'id_master = id_nv1 & id_nv2 & id_nv3
id_master = id_nv1 & "-" & id_nv2 & "-" & id_nv3
End If



selectedRow = Me.txtComboBox_nv4_4.ListIndex

If sDiversos = True Then
   sql1 = "SELECT ID FROM t_Servicos_Diversos_Insumos WHERE ID = '" & id_master & "';"
Else
   sql1 = "SELECT ID FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "';"
End If


rs.Open sql1, conexao

On Error GoTo here
id = rs.Fields("ID").Value
If id <> "" Then

GetPvs_Cmo
 

  If sPrincipais = True Then
 
  End If
  
    UserForm2.txtBox_pvs.Value = Format(pvs, "R$ #,##0.00")
    UserForm2.txtBox_pvs.Enabled = True
    Me.txtBox_pvs.BackColor = RGB(255, 255, 255)
    
    If sDiversos = True Then
    UserForm2.txtBox_pvs.Value = Format("", "R$ #,##0.00")
    UserForm2.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    End If
    
    UserForm2.txtBox_cmo.Value = Format(cmo, "R$ #,##0.00")
    UserForm2.txtBox_cmo.Enabled = False
    UserForm2.txtBox_cmo.BackColor = &H80000016
  


Else
'Sairá do código, como não existe valor já cadastrado (ID) para PVS e CMO, será cadastrado de forma manual
'Liberará os campos de PVS e CMO, assim sempre estarão bloqueados, menos quando não houver valor
here:

c_cmo = True
c_pvs = True

UserForm2.txtBox_pvs.Value = Format("", "R$ #,##0.00")
UserForm2.txtBox_pvs.Enabled = True
Me.txtBox_pvs.BackColor = RGB(255, 255, 255)

UserForm2.txtBox_cmo.Value = ""
UserForm2.txtBox_cmo.Enabled = True
Me.txtBox_cmo.BackColor = RGB(255, 255, 255)

    If sDiversos = True Then
    UserForm2.txtBox_pvs.Value = Format("", "R$ #,##0.00")
    UserForm2.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    End If
End If
rs.Close
conexao.Close

End Sub

Sub att_pvs()
If UserForm2.txtBox_pvs.Value <> "" Then
    If UserForm2.txtBox_pvs.Value = pvs Then
        
        'MsgBox "O usuário escolheu manter o valor existente de pvs = " & Pvs
        Exit Sub
    Else
        'PvsUpdate = UserForm2.txtBox_pvs.Value
        If UserForm2.txtBox_pvs.Value = "R$ " Or UserForm2.txtBox_pvs.Value = "R$" Or UserForm2.txtBox_pvs.Value = "," Or UserForm2.txtBox_pvs.Value = ", " Then
    
        
        UserForm2.txtBox_pvs.Value = pvs
        Else
        PvsUpdate = UserForm2.txtBox_pvs.Value
        End If
    End If
    
Else
Exit Sub
End If

End Sub


Private Sub ComboBox10_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Background_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton14_Click()

 If Me.Image1.Visible = True Then
        Me.Image1.Visible = False
    Else
        Me.Image1.Visible = True
    End If
End Sub

Private Sub CommandButton25_Click()

End Sub

Private Sub CommandButton27_Click()

End Sub

Private Sub CommandButton32_Click()

If ex_sPrincipais = True Then
Exclusão.listarPendentes
ElseIf ex_sDiversos = True Or ex_sTerceiros = True Then
Exclusão.ex_deleteEstrutura
ElseIf Me.ex_sGeneralInsumo = True Then
    'Exclusão.listarPendentes
Exclusão.verificarPendenteInsumo
End If

End Sub

Private Sub ed_txtBox_nv2_2_Change()

End Sub

Private Sub ex_BtnSelectionDiversos_Click()

''\!/---BLOQUEIA O ACESSO DO USUÁRIO---\!/
'stopApplication
'If stop_Application Then Exit Sub
''\!/----------------------------------\!/



ex_BtnSelectionDiversosBoolean = True
ex_BtnSelectionPrincipaisBoolean = False

ex_ComboBoxNv1.Enabled = False
Me.ex_ComboBoxNv1.BackColor = &H8000000F
ex_ComboBoxNv2.Enabled = False
Me.ex_ComboBoxNv2.BackColor = &H8000000F
ex_ComboBoxNv3.Enabled = True
ex_ComboBoxNv3.BackColor = RGB(255, 255, 255)
ex_ComboBoxNv4.Enabled = False
ex_ComboBoxNv4.BackColor = &H8000000F


'Limpar ComboBoxes
ex_ComboBoxNv1.Clear
ex_ComboBoxNv2.Clear
ex_ComboBoxNv3.Clear
ex_ComboBoxNv4.Clear
'Carregar função de lista
'Exclusão.ex_CarregarNivelPrincipal
ex_BtnSelectionPrincipais.BackColor = &H8000000F
ex_BtnSelectionPrincipais.Font.Bold = False
ex_BtnSelectionDiversos.BackColor = RGB(255, 230, 153)
ex_BtnSelectionDiversos.Font.Bold = True
carregarEstruturaPrincipais3


End Sub

Private Sub ex_BtnSelectionPrincipais_Click()
ex_BtnSelectionPrincipaisBoolean = True
ex_BtnSelectionDiversosBoolean = False
ex_ComboBoxNv1.Enabled = True
ex_ComboBoxNv1.BackColor = RGB(255, 255, 255)
ex_ComboBoxNv2.Enabled = False
ex_ComboBoxNv2.BackColor = &H8000000F
ex_ComboBoxNv3.Enabled = False
ex_ComboBoxNv3.BackColor = &H8000000F
ex_ComboBoxNv4.Visible = True
ex_ComboBoxNv4.Enabled = False
ex_ComboBoxNv4.BackColor = &H8000000F
ex_ComboBoxNv1.Clear
ex_ComboBoxNv2.Clear
ex_ComboBoxNv3.Clear
ex_ComboBoxNv4.Clear
'ex_CarregarNivelPrincipal
ex_BtnSelectionPrincipais.BackColor = RGB(142, 162, 219)
ex_BtnSelectionPrincipais.Font.Bold = True
ex_BtnSelectionDiversos.BackColor = &H8000000F
ex_BtnSelectionDiversos.Font.Bold = False
carregarEstruturaPrincipais
End Sub

Private Sub ex_ComboBoxNv4_Change()

End Sub

Private Sub ex_estruturaInsumo_Click()
If ex_sPrincipais = True Or Me.ex_sGeneralInsumo = True And Me.ex_BtnSelectionPrincipaisBoolean = True Then
    If ex_ComboBoxNv1 <> "" And ex_ComboBoxNv2 <> "" And ex_ComboBoxNv3 <> "" Then
        Exclusão.ex_deleteEstrutura
        Exit Sub
        Else
        MsgBox "Existem valores a serem preenchidos."
        Exit Sub
    End If
End If

If Me.ex_sGeneralInsumo = True And Me.ex_BtnSelectionDiversosBoolean = True Then
    Exclusão.ex_deleteEstrutura
    Exit Sub
    Else
    MsgBox "Existem valores a serem preenchidos."
    Exit Sub
End If

End Sub

Private Sub CommandButton36_Click()

End Sub

Private Sub CommandButton35_Click()

End Sub

Private Sub CommandButton33_Click()

End Sub

Private Sub CommandButton37_Click()
menuEstrutura = True
Me.MultiPage4(1).Visible = False
Me.MultiPage4(2).Visible = False
Me.MultiPage4(0).Visible = True
Me.MultiPage4.Value = 0

ex_ComboBoxNv1.Clear
ex_ComboBoxNv2.Clear
ex_ComboBoxNv3.Clear
ex_ComboBoxNv4.Clear
End Sub

Private Sub CommandButton38_Click()

Me.MultiPage4(1).Visible = False
Me.MultiPage4(2).Visible = False
Me.MultiPage4(0).Visible = True
Me.MultiPage4.Value = 0
End Sub

Private Sub CommandButton39_Click()
Me.MultiPage2.Value = 0
End Sub

Private Sub ex_insumoGerenal_Click()



Dim resposta As VbMsgBoxResult

If GlobalServiceType = "Servicos Diversos" Or GlobalServiceType = "Servicos de Terceiros" Then

GlobalTable = ""
GlobalComboBoxValue = ""
End If

If GlobalServiceType <> "" And GlobalTable = "" And GlobalComboBoxValue = "" Then

    Me.ex_BtnSelectionPrincipaisBoolean = True
    Me.ex_BtnSelectionDiversosBoolean = False
    menuEstrutura = False
    ex_sPrincipais = False
    ex_sDiversos = False
    ex_sTerceiros = False
    ex_sGeneralInsumo = True
    
    
    Me.ex_btnNv1.Enabled = False
    Me.ex_btnNv1.BackColor = &H8000000F
    Me.ex_btnNv2.Enabled = False
    Me.ex_btnNv2.BackColor = &H8000000F
    Me.ex_btnNv3.Enabled = False
    Me.ex_btnNv3.BackColor = &H8000000F
    Me.ex_btnNv4.Enabled = True
    

    Me.ex_btnEstrutura.Enabled = True
    'Me.ex_btnEstrutura.BackColor = &H8000000F
    
    ex_OptionButtonPrincipais = True
    
    Me.ex_s_principais.BackColor = &H8000000F
    Me.ex_s_principais.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = RGB(100, 120, 150)
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = True
    
    Me.ex_s_diversos.BackColor = &H8000000F
    Me.ex_s_diversos.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = False
    
    Me.ex_s_terceiros.BackColor = &H8000000F
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = False
    
    ex_ListBox.Clear
    
    GlobalServiceType = "Insumos"
    GlobalTable = ""
    GlobalComboBoxValue = ""
'    GlobalComboBoxValue = UserForm2.ex_ComboBoxNiveis

End If




If GlobalServiceType = "Servicos Diversos" And GlobalTable = "TabelaNv2" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos de Terceiros" And GlobalTable = "TabelaNv3" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos Principais" And GlobalTable <> "" And GlobalComboBoxValue <> "" Then
    resposta = MsgBox("Os dados da última lista gerada serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then

    Me.ex_BtnSelectionPrincipaisBoolean = True
    Me.ex_BtnSelectionDiversosBoolean = False
    menuEstrutura = False
    ex_sPrincipais = False
    ex_sDiversos = False
    ex_sTerceiros = False
    ex_sGeneralInsumo = True
    
    
    Me.ex_btnNv1.Enabled = False
    Me.ex_btnNv1.BackColor = &H8000000F
    Me.ex_btnNv2.Enabled = False
    Me.ex_btnNv2.BackColor = &H8000000F
    Me.ex_btnNv3.Enabled = False
    Me.ex_btnNv3.BackColor = &H8000000F
    Me.ex_btnNv4.Enabled = True
    
    'Me.ex_btnNv4.BackColor = &H8000000F
    'ActiveSheet.Shapes("btnNv4").Fill.ForeColor.RGB = -4142
    Me.ex_btnEstrutura.Enabled = True
    'Me.ex_btnEstrutura.BackColor = &H8000000F
    
    ex_OptionButtonPrincipais = True
    
    Me.ex_s_principais.BackColor = &H8000000F
    Me.ex_s_principais.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = RGB(100, 120, 150)
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = True
    
    Me.ex_s_diversos.BackColor = &H8000000F
    Me.ex_s_diversos.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = False
    
    Me.ex_s_terceiros.BackColor = &H8000000F
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = False
    
    ex_ListBox.Clear
    
    GlobalServiceType = "Insumos"
    GlobalTable = ""
    GlobalComboBoxValue = UserForm2.ex_ComboBoxNiveis

    
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub ed_btnCmoPvs_Click()
ed_btnNv1Boolean = False
ed_btnNv2Boolean = False
ed_btnNv3Boolean = False
ed_btnCmoPvsBoolean = True
ed_btnNv4Boolean = False
ed_btnRendimentoBoolean = False

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True

    RestoreObjects

Me.MultiPage3.Value = 1
Edição.ed_CarregarNivelUm
Edição.ed_CarregarNivelDois
Edição.ed_CarregarNivelTres
'Edição.ed_CarregarInsumos

Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = True
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = True
Me.ed_txtBox_nv3_3.Visible = False
Me.ed_txtComboBox_nv3_3.Visible = True
Me.ed_txtBox_nv4_4.Visible = False
Me.ed_txtComboBox_nv4_4.Visible = False
Me.ed_txtBox_un.Visible = False
Me.ed_txtBox_rendimento.Visible = False
Me.ed_txtBox_custoInsumo.Visible = False
Me.ed_txtBox_pvs.Visible = True
Me.ed_txtBox_cmo.Visible = True
Me.ed_descripUn.Visible = False
Me.ed_descripRend.Visible = False
Me.ed_descripPvs.Visible = True
Me.ed_descripCmo.Visible = True
Me.ed_descripCustInsumo.Visible = False
Me.ed_title_nv1.Visible = True
Me.ed_title_nv2.Visible = True
Me.ed_title_nv3.Visible = True
Me.ed_title_nv4.Visible = False

If ed_sDiversos = True Then
Me.ed_title_nv1.Visible = False
Me.ed_title_nv2.Visible = False
Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = False
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = False

Me.ed_descripPvs.Visible = False
Me.ed_txtBox_pvs.Visible = False

End If

Me.ed_txtComboBox_nv2_2.Enabled = False
Me.ed_txtComboBox_nv2_2.BackColor = &H8000000F

Me.ed_txtComboBox_nv3_3.Enabled = False
Me.ed_txtComboBox_nv3_3.BackColor = &H8000000F


If ed_sDiversos = True Then
Me.ed_txtComboBox_nv3_3.Enabled = True
Me.ed_txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
    ed_carregarEstruturaPrincipais3
End If


Me.ed_txtBox_pvs.Enabled = False
Me.ed_txtBox_pvs.BackColor = &H8000000F
Me.ed_txtBox_cmo.Enabled = False
Me.ed_txtBox_cmo.BackColor = &H8000000F

RestoreObjects
moveObjects
'ed_sDiversos = True
End Sub

Private Sub ed_btnRendimento_Click()

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True
Me.MultiPage3.Value = 1

ed_btnNv1Boolean = False
ed_btnNv2Boolean = False
ed_btnNv3Boolean = False
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = False
ed_btnRendimentoBoolean = True



Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = True
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = True
Me.ed_txtBox_nv3_3.Visible = False
Me.ed_txtComboBox_nv3_3.Visible = True
Me.ed_txtBox_nv4_4.Visible = False
Me.ed_txtComboBox_nv4_4.Visible = True
Me.ed_txtBox_un.Visible = False
Me.ed_txtBox_rendimento.Visible = True
Me.ed_txtBox_custoInsumo.Visible = False
Me.ed_txtBox_pvs.Visible = False
Me.ed_txtBox_cmo.Visible = False
Me.ed_descripUn.Visible = False
Me.ed_descripRend.Visible = True
Me.ed_descripPvs.Visible = False
Me.ed_descripCmo.Visible = False
Me.ed_descripCustInsumo.Visible = False
Me.ed_title_nv1.Visible = True
Me.ed_title_nv2.Visible = True
Me.ed_title_nv3.Visible = True
Me.ed_title_nv4.Visible = True

If ed_sPrincipais = True Then
Me.ed_txtComboBox_nv2_2.Enabled = False
Me.ed_txtComboBox_nv2_2.BackColor = &H8000000F

Me.ed_txtComboBox_nv3_3.Enabled = False
Me.ed_txtComboBox_nv3_3.BackColor = &H8000000F

Me.ed_txtComboBox_nv4_4.Enabled = False
Me.ed_txtComboBox_nv4_4.BackColor = &H8000000F

Me.ed_txtBox_rendimento.Enabled = False
Me.ed_txtBox_rendimento.BackColor = &H8000000F

Edição.ed_carregarEstruturaPrincipais
ElseIf ed_sDiversos = True Then

    If Me.ed_btnRendimentoBoolean = True Then
    Me.ed_txtComboBox_nv4_4.Enabled = False
    Me.ed_txtComboBox_nv4_4.BackColor = &H8000000F
    
    Me.ed_txtBox_rendimento.Enabled = False
    Me.ed_txtBox_rendimento.BackColor = &H8000000F
    End If
    
    ed_verify_estructure
    ed_carregarEstruturaPrincipais3
    

End If

If ed_btnRendimentoBoolean = True And ed_sDiversos = True Then
Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = False
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = False
Me.ed_title_nv1.Visible = False
Me.ed_title_nv2.Visible = False
End If

RestoreObjects
moveObjects
End Sub

Private Sub ed_btnNv4_Click()
Edição.ed_CarregarInsumos
'Edição.ed_GetIdInsumo
ed_btnNv1Boolean = False
ed_btnNv2Boolean = False
ed_btnNv3Boolean = False
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = True
ed_btnRendimentoBoolean = False
'Nv4 bloqueio

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True
Me.MultiPage3.Value = 1
Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = False
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = False
Me.ed_txtBox_nv3_3.Visible = False
Me.ed_txtComboBox_nv3_3.Visible = False
Me.ed_txtBox_nv4_4.Visible = True
Me.ed_txtComboBox_nv4_4.Visible = True
Me.ed_txtBox_un.Visible = True
Me.ed_txtBox_rendimento.Visible = False
Me.ed_txtBox_custoInsumo.Visible = True
Me.ed_txtBox_pvs.Visible = False
Me.ed_txtBox_cmo.Visible = False
Me.ed_descripUn.Visible = True
Me.ed_descripRend.Visible = False
Me.ed_descripPvs.Visible = False
Me.ed_descripCmo.Visible = False
Me.ed_descripCustInsumo.Visible = True
Me.ed_title_nv1.Visible = False
Me.ed_title_nv2.Visible = False
Me.ed_title_nv3.Visible = False
Me.ed_title_nv4.Visible = True

Me.ed_txtBox_un.Enabled = False
Me.ed_txtBox_un.BackColor = &H80000016
Me.ed_txtBox_custoInsumo.Enabled = False
Me.ed_txtBox_custoInsumo.BackColor = &H80000016

Me.ed_descripTipo.Visible = True
Me.ed_ComboBoxTipo.Visible = True
Me.ed_ComboBoxTipo.Enabled = False
Me.ed_ComboBoxTipo.BackColor = &H80000016

RestoreObjects
moveObjects


'=============================
'Bloqueia de forma temporária o objeto abaixo em  edição
If ed_sDiversos = True Then
UserForm2.ed_txtBox_nv4_4.Visible = False
End If




End Sub

Sub ed_btnNv4_carregar()
Edição.ed_CarregarInsumos
'Edição.ed_GetIdInsumo
ed_btnNv1Boolean = False
ed_btnNv2Boolean = False
ed_btnNv3Boolean = False
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = True
ed_btnRendimentoBoolean = False
'Nv4 bloqueio

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True
Me.MultiPage3.Value = 1
Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = False
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = False
Me.ed_txtBox_nv3_3.Visible = False
Me.ed_txtComboBox_nv3_3.Visible = False
Me.ed_txtBox_nv4_4.Visible = True
Me.ed_txtComboBox_nv4_4.Visible = True
Me.ed_txtBox_un.Visible = True
Me.ed_txtBox_rendimento.Visible = False
Me.ed_txtBox_custoInsumo.Visible = True
Me.ed_txtBox_pvs.Visible = False
Me.ed_txtBox_cmo.Visible = False
Me.ed_descripUn.Visible = True
Me.ed_descripRend.Visible = False
Me.ed_descripPvs.Visible = False
Me.ed_descripCmo.Visible = False
Me.ed_descripCustInsumo.Visible = True
Me.ed_title_nv1.Visible = False
Me.ed_title_nv2.Visible = False
Me.ed_title_nv3.Visible = False
Me.ed_title_nv4.Visible = True

Me.ed_txtBox_un.Enabled = False
Me.ed_txtBox_un.BackColor = &H80000016
Me.ed_txtBox_custoInsumo.Enabled = False
Me.ed_txtBox_custoInsumo.BackColor = &H80000016

Me.ed_descripTipo.Visible = True
Me.ed_ComboBoxTipo.Visible = True
Me.ed_ComboBoxTipo.Enabled = False
Me.ed_ComboBoxTipo.BackColor = &H80000016
Me.ed_ComboBoxTipo.Clear
'RestoreObjects
'moveObjects


'=============================
'Bloqueia de forma temporária o objeto abaixo em  edição
If ed_sDiversos = True Then
UserForm2.ed_txtBox_nv4_4.Visible = False
End If




End Sub

Private Sub CommandButton3_Click()

End Sub

Sub atualizarDados()

'Retira espaços antes e depois do comboBox e repor formatado antes de adicionar no BD
 Dim texto As String
 

    
    ' Obtém o valor da TextBox
    texto = Trim(txtBox_nv1_1.Value)
    texto = UCase(texto)
    

    Do While InStr(texto, "  ") > 0
    texto = Replace(texto, "  ", " ")
    Loop
    RemoverEspacosDuplos = texto
    txtBox_nv1_1.Value = texto
    
    texto = ""
    texto = Trim(txtBox_nv2_2.Value)
    texto = UCase(texto)
    Do While InStr(texto, "  ") > 0
    texto = Replace(texto, "  ", " ")
    Loop
    RemoverEspacosDuplos = texto
    txtBox_nv2_2.Value = texto
    
    texto = ""
    texto = Trim(txtBox_nv3_3.Value)
    texto = UCase(texto)
    Do While InStr(texto, "  ") > 0
    texto = Replace(texto, "  ", " ")
    Loop
    RemoverEspacosDuplos = texto
    txtBox_nv3_3.Value = texto
    
    texto = ""
    texto = Trim(txtBox_nv4_4.Value)
    texto = UCase(texto)
    Do While InStr(texto, "  ") > 0
    texto = Replace(texto, "  ", " ")
    Loop
    RemoverEspacosDuplos = texto
    txtBox_nv4_4.Value = texto
    


If sDiversos = True Then

GetIdNv3
GetIdInsumo
GetCustoInsumo
GetUnidade
atualizarPvsCmo
GetRendimentoDiversos

ElseIf sPrincipais = True Then
GetIdNv1
GetIdNv2
GetIdNv3
GetIdInsumo
GetCustoInsumo
GetUnidade
atualizarPvsCmo
Me.GetRendimento1

ElseIf sTerceiros = True Then
GetIdNv3
End If
End Sub

Sub naoMostrarCampos()


Visible = False


If sPrincipais = True Then

    If listaNv1 = True Then
    txtComboBox_nv1_1.Enabled = True
    txtComboBox_nv1_1.BackColor = RGB(255, 255, 255)
    Else
    txtBox_nv1_1.Enabled = True
    txtBox_nv1_1.BackColor = RGB(255, 255, 255)
    End If
    
    If listaNv2 = True Then
    txtComboBox_nv2_2.Enabled = True
    txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)
    Else
    txtBox_nv2_2.Enabled = True
    txtBox_nv2_2.BackColor = RGB(255, 255, 255)
    End If
    
    If listaNv3 = True Then
    txtComboBox_nv3_3.Enabled = True
    txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
    Else
    txtBox_nv3_3.Enabled = True
    txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    End If
    
    If listaNv4 = True Then
    txtComboBox_nv4_4.Enabled = True
    txtComboBox_nv4_4.BackColor = RGB(255, 255, 255)
    Else
    txtBox_nv4_4.Enabled = True
    txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    End If
    
    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True

End If



If sDiversos = True Then


    If listaNv3 = True Then
    txtComboBox_nv3_3.Enabled = True
    txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
    Else
    txtBox_nv3_3.Enabled = True
    txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    End If
    
    If listaNv4 = True Then
    txtComboBox_nv4_4.Enabled = True
    txtComboBox_nv4_4.BackColor = RGB(255, 255, 255)
    Else
    txtBox_nv4_4.Enabled = True
    txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    End If
    
    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True

End If




Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.CheckBoxGeneric.Visible = False
Me.CheckBoxGeneric.Value = False

Me.CommandButton15.Visible = True
Me.EditarNiveis.Visible = False

listaNv1 = False
txtBoxNv1 = False
listaNv2 = False
txtBoxNv2 = False
listaNv3 = False
txtBoxNv3 = False
listaNv4 = False
txtBoxNv4 = False
End Sub

Sub mostrarCampos()




txtBoxNv1 = False
listaNv1 = False
txtBoxNv2 = False
listaNv2 = False
txtBoxNv3 = False
listaNv3 = False
txtBoxNv4 = False
listaNv4 = False





Visible = True

If sPrincipais = True Then
    
    If txtBox_nv1_1.Enabled = True Then
    txtBoxNv1 = True
    Else
    listaNv1 = True
    End If

    If txtBox_nv2_2.Enabled = True Then
    txtBoxNv2 = True
    Else
    listaNv2 = True
    End If
    
    If txtBox_nv3_3.Enabled = True Then
    txtBoxNv3 = True
    Else
    listaNv3 = True
    End If
    
    If txtBox_nv4_4.Enabled = True Then
    txtBoxNv4 = True
    Else
    listaNv4 = True
    End If

    
    Me.Label27.Visible = True
    Me.txtBox_un.Visible = True
    Me.Label38.Visible = True
    Me.txtBox_rendimento.Visible = True
    Me.Label36.Visible = True
    Me.txtBox_cmo.Visible = True
    Me.Label44.Visible = True
    Me.txtBox_custoInsumo.Visible = True
    Me.Label35.Visible = True
    Me.txtBox_pvs.Visible = True
    Me.Enviar.Visible = True
    Me.CommandButton15.Visible = False
    Me.EditarNiveis.Visible = True

    
    If txtBox_nv1_1.Enabled = True Then
    txtBox_nv1_1.Enabled = False
    txtBox_nv1_1.BackColor = &H80000016
    Else
    txtComboBox_nv1_1.Enabled = False
    txtComboBox_nv1_1.BackColor = &H80000016
    End If
    
    If txtBox_nv2_2.Enabled = True Then
    txtBox_nv2_2.Enabled = False
    txtBox_nv2_2.BackColor = &H80000016
    Else
    txtComboBox_nv2_2.Enabled = False
    txtComboBox_nv2_2.BackColor = &H80000016
    End If
    
    If txtBox_nv3_3.Enabled = True Then
    txtBox_nv3_3.Enabled = False
    txtBox_nv3_3.BackColor = &H80000016
    Else
    txtComboBox_nv3_3.Enabled = False
    txtComboBox_nv3_3.BackColor = &H80000016
    End If
    
    If txtBox_nv4_4.Enabled = True Then
    txtBox_nv4_4.Enabled = False
    txtBox_nv4_4.BackColor = &H80000016
    Me.CheckBoxGeneric.Visible = True
    Else
    txtComboBox_nv4_4.Enabled = False
    txtComboBox_nv4_4.BackColor = &H80000016
    End If
    
    optionButton_nv1_1.Enabled = False
    optionButton_nv1_2.Enabled = False
    optionButton_nv2_3.Enabled = False
    optionButton_nv2_4.Enabled = False
    optionButton_nv3_5.Enabled = False
    optionButton_nv3_6.Enabled = False
    optionButton_nv4_7.Enabled = False
    optionButton_nv4_8.Enabled = False

End If

If sDiversos = True Then




    If txtBox_nv3_3.Enabled = True Then
    txtBoxNv3 = True
    Else
    listaNv3 = True
    End If
    
    If txtBox_nv4_4.Enabled = True Then
    txtBoxNv4 = True
    Else
    listaNv4 = True
    End If
    '

    
    
    Me.Label27.Visible = True
    Me.txtBox_un.Visible = True
    Me.Label38.Visible = True
    Me.txtBox_rendimento.Visible = True
    Me.Label36.Visible = True
    Me.txtBox_cmo.Visible = True
    Me.Label44.Visible = True
    Me.txtBox_custoInsumo.Visible = True
    Me.Label35.Visible = False
    Me.txtBox_pvs.Visible = False
    Me.Enviar.Visible = True
    Me.CommandButton15.Visible = False
    Me.EditarNiveis.Visible = True

    
    If txtBox_nv3_3.Enabled = True Then
    txtBox_nv3_3.Enabled = False
    txtBox_nv3_3.BackColor = &H80000016
    Else
    txtComboBox_nv3_3.Enabled = False
    txtComboBox_nv3_3.BackColor = &H80000016
    End If
    
    If txtBox_nv4_4.Enabled = True Then
    txtBox_nv4_4.Enabled = False
    txtBox_nv4_4.BackColor = &H80000016
    CheckBoxGeneric.Visible = True
    CheckBoxGeneric.Value = False
    Else
    txtComboBox_nv4_4.Enabled = False
    txtComboBox_nv4_4.BackColor = &H80000016
    End If



    optionButton_nv1_1.Enabled = False
    optionButton_nv1_2.Enabled = False
    optionButton_nv2_3.Enabled = False
    optionButton_nv2_4.Enabled = False
    optionButton_nv3_5.Enabled = False
    optionButton_nv3_6.Enabled = False
    optionButton_nv4_7.Enabled = False
    optionButton_nv4_8.Enabled = False


End If



'EditarNiveis.Visible = True

End Sub



Private Sub CommandButton15_Click()




If sPrincipais Then
    If (txtBox_nv1_1.Value <> "" Or txtComboBox_nv1_1.Value <> "") And (txtBox_nv2_2.Value <> "" Or txtComboBox_nv2_2.Value <> "") And (txtBox_nv3_3.Value <> "" Or txtComboBox_nv3_3.Value <> "") And (txtBox_nv4_4.Value <> "" Or txtComboBox_nv4_4.Value <> "") Then
'Atualizar dados irá gerenciar todos os métodos responsáveis por pegar IDs, checar existência de rendimento, cmo, pvs, etc
    atualizarDados
        If c_rendimento <> True Then
        mostrarCampos
        Else
        MsgBox "Esta estrutura de serviço já existe.", vbExclamation
        
        ClearFields
        End If
    Else
    MsgBox "Existem campos para serem preenchidos.", vbExclamation
    End If
ElseIf sDiversos Then
    If (txtBox_nv3_3.Value <> "" Or txtComboBox_nv3_3.Value <> "") And (txtBox_nv4_4.Value <> "" Or txtComboBox_nv4_4.Value <> "") Then
    atualizarDados
        If c_rendimento <> True Then
        mostrarCampos
        Else
        MsgBox "Esta estrutura de serviço já existe.", vbExclamation
        ClearFields
        End If
    Else
    MsgBox "Existem campos para serem preenchidos.", vbExclamation
    End If
ElseIf sTerceiros Then
    If txtBox_nv3_3.Value <> "" Then
    atualizarDados
        If id_nv3 <> 0 Then
        MsgBox "Esta estrutura de serviço já existe.", vbExclamation
        ClearFields
        Exit Sub
        End If
    CommandButton15.Visible = False
    Enviar.Visible = True
    
    Else
    MsgBox "Existem campos para serem preenchidos.", vbExclamation
    End If

    
    
End If


End Sub

Private Sub ed_btnNv3_Click()
Edição.ed_CarregarNivelTres
ed_btnNv1Boolean = False
ed_btnNv2Boolean = False
ed_btnNv3Boolean = True
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = False
ed_btnRendimentoBoolean = False
Edição.ed_GetIdNv3

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True
Me.MultiPage3.Value = 1
Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = False
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = False
Me.ed_txtBox_nv3_3.Visible = True
Me.ed_txtComboBox_nv3_3.Visible = True
Me.ed_txtBox_nv4_4.Visible = False
Me.ed_txtComboBox_nv4_4.Visible = False
Me.ed_txtBox_un.Visible = False
Me.ed_txtBox_rendimento.Visible = False
Me.ed_txtBox_custoInsumo.Visible = False
Me.ed_txtBox_pvs.Visible = False
Me.ed_txtBox_cmo.Visible = False
Me.ed_descripUn.Visible = False
Me.ed_descripRend.Visible = False
Me.ed_descripPvs.Visible = False
Me.ed_descripCmo.Visible = False
Me.ed_descripCustInsumo.Visible = False
Me.ed_title_nv1.Visible = False
Me.ed_title_nv2.Visible = False
Me.ed_title_nv3.Visible = True
Me.ed_title_nv4.Visible = False

RestoreObjects
moveObjects
End Sub

Private Sub ed_btnNv2_Click()
ed_CarregarNivelDois
ed_btnNv1Boolean = False
ed_btnNv2Boolean = True
ed_btnNv3Boolean = False
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = False
ed_btnRendimentoBoolean = False
ed_GetIdNv2

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True
Me.MultiPage3.Value = 1
Me.ed_txtBox_nv1_1.Visible = False
Me.ed_txtComboBox_nv1_1.Visible = False
Me.ed_txtBox_nv2_2.Visible = True
Me.ed_txtComboBox_nv2_2.Visible = True
Me.ed_txtBox_nv3_3.Visible = False
Me.ed_txtComboBox_nv3_3.Visible = False
Me.ed_txtBox_nv4_4.Visible = False
Me.ed_txtComboBox_nv4_4.Visible = False
Me.ed_txtBox_un.Visible = False
Me.ed_txtBox_rendimento.Visible = False
Me.ed_txtBox_custoInsumo.Visible = False
Me.ed_txtBox_pvs.Visible = False
Me.ed_txtBox_cmo.Visible = False
Me.ed_descripUn.Visible = False
Me.ed_descripRend.Visible = False
Me.ed_descripPvs.Visible = False
Me.ed_descripCmo.Visible = False
Me.ed_descripCustInsumo.Visible = False
Me.ed_title_nv1.Visible = False
Me.ed_title_nv2.Visible = True
Me.ed_title_nv3.Visible = False
Me.ed_title_nv4.Visible = False


RestoreObjects
moveObjects




End Sub

Sub saveObjectsPosition()
nv2LeftCombo = Me.ed_txtComboBox_nv2_2.Left
nv2TopCombo = Me.ed_txtComboBox_nv2_2.Top

nv2LeftTxt = Me.ed_txtBox_nv2_2.Left
nv2TopTxt = Me.ed_txtBox_nv2_2.Top

nv3LeftCombo = Me.ed_txtComboBox_nv3_3.Left
nv3TopCombo = Me.ed_txtComboBox_nv3_3.Top

nv3LeftTxt = Me.ed_txtBox_nv3_3.Left
nv3TopTxt = Me.ed_txtBox_nv3_3.Top

nv2TitleTop = ed_title_nv2.Top
nv2TitleLeft = ed_title_nv2.Left

nv3TitleTop = ed_title_nv3.Top
nv3TitleLeft = ed_title_nv3.Left

nv1LeftTxt = Me.ed_txtBox_nv1_1.Left
nv1TopTxt = Me.ed_txtBox_nv1_1.Top

nv1LeftCombo = Me.ed_txtComboBox_nv1_1.Left
nv1TopCombo = Me.ed_txtComboBox_nv1_1.Top

nv1TitleTop = ed_title_nv1.Top
nv1TitleLeft = ed_title_nv1.Left

ed_title_nv4Left = ed_title_nv4.Left
ed_txtComboBox_nv4_4Left = ed_txtComboBox_nv4_4.Left
ed_txtBox_nv4_4Left = ed_txtBox_nv4_4.Left
ed_descripUnLeft = ed_descripUn.Left
ed_txtBox_unLeft = ed_txtBox_un.Left
ed_descripCustInsumoLeft = ed_descripCustInsumo.Left
ed_txtBox_custoInsumoLeft = ed_txtBox_custoInsumo.Left

ed_title_nv4Top = ed_title_nv4.Top
ed_txtComboBox_nv4_4Top = ed_txtComboBox_nv4_4.Top
ed_txtBox_nv4_4Top = ed_txtBox_nv4_4.Top
ed_descripUnTop = ed_descripUn.Top
ed_txtBox_unTop = ed_txtBox_un.Top
ed_descripCustInsumoTop = ed_descripCustInsumo.Top
ed_txtBox_custoInsumoTop = ed_txtBox_custoInsumo.Top


    ed_descripCmoTop = ed_descripCmo.Top
    ed_descripCmoLeft = ed_descripCmo.Left

    ed_txtBox_cmoTop = ed_txtBox_cmo.Top
    ed_txtBox_cmoLeft = ed_txtBox_cmo.Left
    
    ed_descripRendTop = ed_descripRend.Top
    ed_descripRendLeft = ed_descripRend.Left

    ed_txtBox_rendimentoTop = ed_txtBox_rendimento.Top
    ed_txtBox_rendimentoLeft = ed_txtBox_rendimento.Left
    
    
    ed_descripPvsTop = ed_descripPvs.Top
     ed_descripPvsLeft = ed_descripPvs.Left
    
     ed_txtBox_pvsTop = ed_txtBox_pvs.Top
     ed_txtBox_pvsLeft = ed_txtBox_pvs.Left


End Sub


Private Sub RestoreObjects()

    'Restaura as coordenadas originais dos objetos'
    Me.ed_txtComboBox_nv2_2.Left = nv2LeftCombo
    Me.ed_txtComboBox_nv2_2.Top = nv2TopCombo
    
    Me.ed_txtBox_nv2_2.Top = nv2TopTxt
    Me.ed_txtBox_nv2_2.Left = nv2LeftTxt
    
    Me.ed_title_nv2.Left = nv2TitleLeft
    Me.ed_title_nv2.Top = nv2TitleTop
    
    Me.ed_txtComboBox_nv3_3.Left = nv3LeftCombo
    Me.ed_txtComboBox_nv3_3.Top = nv3TopCombo
    
    Me.ed_txtBox_nv3_3.Top = nv3TopTxt
    Me.ed_txtBox_nv3_3.Left = nv3LeftTxt
    
    Me.ed_title_nv3.Left = nv3TitleLeft
    Me.ed_title_nv3.Top = nv3TitleTop
    
    
    ed_title_nv4.Left = ed_title_nv4Left
    ed_txtComboBox_nv4_4.Left = ed_txtComboBox_nv4_4Left
    ed_txtBox_nv4_4.Left = ed_txtBox_nv4_4Left
    ed_descripUn.Left = ed_descripUnLeft
    ed_txtBox_un.Left = ed_txtBox_unLeft
    ed_descripCustInsumo.Left = ed_descripCustInsumoLeft
    ed_txtBox_custoInsumo.Left = ed_txtBox_custoInsumoLeft

    ed_title_nv4.Top = ed_title_nv4Top
    ed_txtComboBox_nv4_4.Top = ed_txtComboBox_nv4_4Top
    ed_txtBox_nv4_4.Top = ed_txtBox_nv4_4Top
    ed_descripUn.Top = ed_descripUnTop
    ed_txtBox_un.Top = ed_txtBox_unTop
    ed_descripCustInsumo.Top = ed_descripCustInsumoTop
    ed_txtBox_custoInsumo.Top = ed_txtBox_custoInsumoTop
    
    
    ed_descripCmo.Top = ed_descripCmoTop
    ed_descripCmo.Left = ed_descripCmoLeft

    ed_txtBox_cmo.Top = ed_txtBox_cmoTop
    ed_txtBox_cmo.Left = ed_txtBox_cmoLeft
    
    ed_descripRend.Top = ed_descripRendTop
    ed_descripRend.Left = ed_descripRendLeft

    ed_txtBox_rendimento.Top = ed_txtBox_rendimentoTop
    ed_txtBox_rendimento.Left = ed_txtBox_rendimentoLeft
    
    ed_descripPvs.Top = ed_descripPvsTop
    ed_descripPvs.Left = ed_descripPvsLeft
    
    ed_txtBox_pvs.Top = ed_txtBox_pvsTop
    ed_txtBox_pvs.Left = ed_txtBox_pvsLeft



End Sub




Sub moveObjects()

If ed_btnNv2Boolean = True Then
' Posiciona as combo box na mesma posição da combo box do Serviço1
    Me.ed_txtComboBox_nv2_2.Left = Me.ed_txtComboBox_nv1_1.Left
    Me.ed_txtComboBox_nv2_2.Top = Me.ed_txtComboBox_nv1_1.Top
    
    Me.ed_txtBox_nv2_2.Left = Me.ed_txtBox_nv1_1.Left
    Me.ed_txtBox_nv2_2.Top = Me.ed_txtBox_nv1_1.Top
    
    Me.ed_title_nv2.Left = Me.ed_title_nv1.Left
    Me.ed_title_nv2.Top = Me.ed_title_nv1.Top
    
ElseIf ed_btnNv3Boolean = True Then


    Me.ed_txtComboBox_nv3_3.Left = Me.ed_txtComboBox_nv1_1.Left
    Me.ed_txtComboBox_nv3_3.Top = Me.ed_txtComboBox_nv1_1.Top

    Me.ed_txtBox_nv3_3.Left = Me.ed_txtBox_nv1_1.Left
    Me.ed_txtBox_nv3_3.Top = Me.ed_txtBox_nv1_1.Top

    Me.ed_title_nv3.Left = Me.ed_title_nv1.Left
    Me.ed_title_nv3.Top = Me.ed_title_nv1.Top
    
ElseIf ed_btnNv4Boolean = True Then

    ed_title_nv4.Top = nv1TitleTop
    ed_title_nv4.Left = nv1TitleLeft
    ed_txtComboBox_nv4_4.Left = nv1LeftCombo
    ed_txtComboBox_nv4_4.Top = nv1TopCombo
    ed_txtBox_nv4_4.Left = nv1LeftTxt
    ed_txtBox_nv4_4.Top = nv1TopTxt

    ed_descripCustInsumo.Left = 515 '636
    ed_descripCustInsumo.Top = 112 '117.5
    ed_txtBox_custoInsumo.Left = 570 '696
    ed_txtBox_custoInsumo.Top = 110 '108
    ed_txtBox_un.Left = 78
    ed_txtBox_un.Top = 112
    ed_descripUn.Left = 12
    ed_descripUn.Top = 112









ElseIf ed_btnCmoPvsBoolean = True And ed_sDiversos = True Then
    Me.ed_title_nv3.Left = Me.ed_title_nv1.Left
    Me.ed_title_nv3.Top = Me.ed_title_nv1.Top

    Me.ed_txtComboBox_nv3_3.Left = Me.ed_txtComboBox_nv1_1.Left
    Me.ed_txtComboBox_nv3_3.Top = Me.ed_txtComboBox_nv1_1.Top
    

    
    Me.ed_txtBox_cmo.Left = 63
    Me.ed_txtBox_cmo.Top = 108
    
    ed_descripCmo.Left = 6
    ed_descripCmo.Top = 108
    
    


ElseIf ed_btnCmoPvsBoolean = True And ed_sPrincipais = True Then
ed_descripPvs.Top = 300
ed_descripPvs.Left = 510
    


ed_txtBox_pvs.Top = 300
ed_txtBox_pvs.Left = 570

Me.ed_txtBox_cmo.Left = 78
Me.ed_txtBox_cmo.Top = 300

ed_descripCmo.Left = 12
ed_descripCmo.Top = 300

ElseIf Me.ed_btnRendimentoBoolean = True And ed_sDiversos = True Then

    Me.ed_title_nv3.Left = Me.ed_title_nv1.Left
    Me.ed_title_nv3.Top = Me.ed_title_nv1.Top

    Me.ed_txtComboBox_nv3_3.Left = Me.ed_txtComboBox_nv1_1.Left
    Me.ed_txtComboBox_nv3_3.Top = Me.ed_txtComboBox_nv1_1.Top
    
    Me.ed_title_nv4.Left = Me.ed_title_nv2.Left
    Me.ed_title_nv4.Top = Me.ed_title_nv2.Top

    Me.ed_txtComboBox_nv4_4.Left = Me.ed_txtComboBox_nv2_2.Left
    Me.ed_txtComboBox_nv4_4.Top = Me.ed_txtComboBox_nv2_2.Top
    




    ed_txtBox_rendimento.Left = 252
    ed_txtBox_rendimento.Top = 186
    
    ed_descripRend.Left = 192
    ed_descripRend.Top = 186
End If



End Sub

Private Sub ed_btnNv1_Click()
ed_btnNv1Boolean = True
ed_btnNv2Boolean = False
ed_btnNv3Boolean = False
Edição.ed_CarregarNivelUm
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = False
ed_btnRendimentoBoolean = False

ed_GetIdNv1

Me.MultiPage3(0).Visible = False
Me.MultiPage3(1).Visible = True
Me.MultiPage3.Value = 1
Me.ed_txtBox_nv1_1.Visible = True
Me.ed_txtComboBox_nv1_1.Visible = True
Me.ed_txtBox_nv2_2.Visible = False
Me.ed_txtComboBox_nv2_2.Visible = False
Me.ed_txtBox_nv3_3.Visible = False
Me.ed_txtComboBox_nv3_3.Visible = False
Me.ed_txtBox_nv4_4.Visible = False
Me.ed_txtComboBox_nv4_4.Visible = False
Me.ed_txtBox_un.Visible = False
Me.ed_txtBox_rendimento.Visible = False
Me.ed_txtBox_custoInsumo.Visible = False
Me.ed_txtBox_pvs.Visible = False
Me.ed_txtBox_cmo.Visible = False
Me.ed_descripUn.Visible = False
Me.ed_descripRend.Visible = False
Me.ed_descripPvs.Visible = False
Me.ed_descripCmo.Visible = False
Me.ed_descripCustInsumo.Visible = False
Me.ed_title_nv1.Visible = True
Me.ed_title_nv2.Visible = False
Me.ed_title_nv3.Visible = False
Me.ed_title_nv4.Visible = False


RestoreObjects

End Sub

Private Sub CommandButton24_Click()
Me.MultiPage2.Value = 0
'------------TESTE-------
RestoreObjects
'------------------------
ed_btnCmoPvsBoolean = False
ed_btnNv4Boolean = False
End Sub

Private Sub CommandButton26_Click()

If Image2.Visible = True Then
Image2.Visible = False
Else
Image2.Visible = True
End If


End Sub



Private Sub ed_Menu_Click()
Me.MultiPage3(0).Visible = True
Me.MultiPage3(1).Visible = False
Me.MultiPage3.Value = 0

Me.ed_txtBox_nv1_1.Value = ""
Me.ed_txtComboBox_nv1_1.Value = ""
Me.ed_txtBox_nv2_2.Value = ""
Me.ed_txtComboBox_nv2_2.Value = ""
Me.ed_txtBox_nv3_3.Value = ""
Me.ed_txtComboBox_nv3_3.Value = ""
Me.ed_txtBox_nv4_4.Value = ""
Me.ed_txtComboBox_nv4_4.Value = ""
Me.ed_txtBox_un.Value = ""
Me.ed_txtBox_rendimento.Value = ""
Me.ed_txtBox_custoInsumo.Value = ""
Me.ed_txtBox_pvs.Value = ""
Me.ed_txtBox_cmo.Value = ""

Me.ed_txtBoxID_nv1 = ""
Me.ed_txtBoxID_nv2 = ""
Me.ed_txtBoxID_nv3 = ""
Me.ed_txtBoxID_nv4 = ""



Me.ed_txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)
Me.ed_txtComboBox_nv2_2.Enabled = True

Me.ed_txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
Me.ed_txtComboBox_nv3_3.Enabled = True

Me.ed_txtComboBox_nv4_4.BackColor = RGB(255, 255, 255)
Me.ed_txtComboBox_nv4_4.Enabled = True

'=============================
'Bloqueia de forma temporária o objeto abaixo em  edição
If id_sDiversos = True Then
UserForm2.ed_txtBox_nv4_4.Visible = True
End If
'=========================




If ed_sPrincipais = True Then
UserForm2.ed_ComboBoxTipo.Clear
Me.ed_descripTipo.Visible = False
Me.ed_ComboBoxTipo.Visible = False
Me.ed_ComboBoxTipo.Enabled = False
ElseIf ed_sDiversos = True Then
UserForm2.ed_ComboBoxTipo.Clear
Me.ed_descripTipo.Visible = False
Me.ed_ComboBoxTipo.Visible = False
Me.ed_ComboBoxTipo.Enabled = False
End If

End Sub

Private Sub sDiversos_Click()

End Sub


Private Sub ed_s_principais_Click()

Me.ed_btnNv1.Enabled = True
Me.ed_btnNv2.Enabled = True
Me.ed_btnNv3.Enabled = True
Me.ed_btnNv4.Enabled = True

Me.ed_btnCmoPvs.Enabled = True
Me.ed_btnRendimento.Enabled = True




UserForm2.ed_btnCmoPvs.Caption = "Editar Custo Mão de Obra | Preço de Venda Sugerido"

ed_sPrincipais = True
ed_sDiversos = False
ed_sTerceiros = False


Me.ed_txtBoxID_nv1 = ""
Me.ed_txtBoxID_nv2 = ""
Me.ed_txtBoxID_nv3 = ""
Me.ed_txtBoxID_nv4 = ""

Me.ed_s_principais.BackColor = RGB(142, 162, 219)
Me.ed_s_terceiros.ForeColor = &H80000008
Me.ed_s_principais.Font.Bold = True



Me.ed_s_diversos.BackColor = &H8000000F
Me.ed_s_diversos.ForeColor = &H80000008
Me.ed_s_diversos.Font.Bold = False

Me.ed_s_terceiros.BackColor = &H8000000F
Me.ed_s_terceiros.ForeColor = &H80000008
Me.ed_s_terceiros.Font.Bold = False

Me.ed_txtBox_nv1_1.Value = ""
Me.ed_txtComboBox_nv1_1.Value = ""
Me.ed_txtBox_nv2_2.Value = ""
Me.ed_txtComboBox_nv2_2.Value = ""
Me.ed_txtBox_nv3_3.Value = ""
Me.ed_txtComboBox_nv3_3.Value = ""
Me.ed_txtBox_nv4_4.Value = ""
Me.ed_txtComboBox_nv4_4.Value = ""
Me.ed_txtBox_un.Value = ""
Me.ed_txtBox_rendimento.Value = ""
Me.ed_txtBox_custoInsumo.Value = ""
Me.ed_txtBox_pvs.Value = ""
Me.ed_txtBox_cmo.Value = ""




End Sub



Private Sub ed_sDiversos_Click()

End Sub

Private Sub ed_s_diversos_Click()
ed_sPrincipais = False
ed_sDiversos = True
ed_sTerceiros = False

Me.ed_btnNv1.Enabled = False
Me.ed_btnNv1.BackColor = &H8000000F
Me.ed_btnNv2.Enabled = False
Me.ed_btnNv2.BackColor = &H8000000F
Me.ed_btnNv3.Enabled = True
Me.ed_btnNv4.Enabled = True

Me.ed_btnCmoPvs.Enabled = True
Me.ed_btnRendimento.Enabled = True

UserForm2.ed_btnCmoPvs.Caption = "Editar Custo Mão de Obra"


Me.ed_txtBoxID_nv1 = ""
Me.ed_txtBoxID_nv2 = ""
Me.ed_txtBoxID_nv3 = ""
Me.ed_txtBoxID_nv4 = ""

Me.ed_s_diversos.BackColor = RGB(255, 230, 153)
Me.ed_s_terceiros.ForeColor = &H80000008
Me.ed_s_diversos.Font.Bold = True


Me.ed_s_principais.BackColor = &H8000000F
Me.ed_s_principais.ForeColor = &H80000008
Me.ed_s_principais.Font.Bold = False

Me.ed_s_terceiros.BackColor = &H8000000F
Me.ed_s_terceiros.ForeColor = &H80000008
Me.ed_s_terceiros.Font.Bold = False

Me.ed_txtBox_nv1_1.Value = ""
Me.ed_txtComboBox_nv1_1.Value = ""
Me.ed_txtBox_nv2_2.Value = ""
Me.ed_txtComboBox_nv2_2.Value = ""
Me.ed_txtBox_nv3_3.Value = ""
Me.ed_txtComboBox_nv3_3.Value = ""
Me.ed_txtBox_nv4_4.Value = ""
Me.ed_txtComboBox_nv4_4.Value = ""
Me.ed_txtBox_un.Value = ""
Me.ed_txtBox_rendimento.Value = ""
Me.ed_txtBox_custoInsumo.Value = ""
Me.ed_txtBox_pvs.Value = ""
Me.ed_txtBox_cmo.Value = ""

End Sub

Private Sub ed_s_terceiros_Click()
ed_sPrincipais = False
ed_sDiversos = False
ed_sTerceiros = True

Me.ed_btnNv1.Enabled = False
Me.ed_btnNv2.Enabled = False
Me.ed_btnNv3.Enabled = True
Me.ed_btnNv4.Enabled = False

Me.ed_btnCmoPvs.Enabled = False
Me.ed_btnRendimento.Enabled = False
UserForm2.ed_btnCmoPvs.Caption = "Editar Custo Mão de Obra | Preço de Venda Sugerido"

Me.ed_txtBoxID_nv1 = ""
Me.ed_txtBoxID_nv2 = ""
Me.ed_txtBoxID_nv3 = ""
Me.ed_txtBoxID_nv4 = ""

Me.ed_s_terceiros.BackColor = RGB(255, 242, 204)
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = True
'Me.MultiPage2.Value = 2
'Me.MultiPage1.BackColor = RGB(255, 249, 231)

Me.ed_s_principais.BackColor = &H8000000F
Me.ed_s_principais.ForeColor = &H80000008
Me.ed_s_principais.Font.Bold = False

Me.ed_s_diversos.BackColor = &H8000000F
Me.ed_s_diversos.ForeColor = &H80000008
Me.ed_s_diversos.Font.Bold = False


Me.ed_txtBox_nv1_1.Value = ""
Me.ed_txtComboBox_nv1_1.Value = ""
Me.ed_txtBox_nv2_2.Value = ""
Me.ed_txtComboBox_nv2_2.Value = ""
Me.ed_txtBox_nv3_3.Value = ""
Me.ed_txtComboBox_nv3_3.Value = ""
Me.ed_txtBox_nv4_4.Value = ""
Me.ed_txtComboBox_nv4_4.Value = ""
Me.ed_txtBox_un.Value = ""
Me.ed_txtBox_rendimento.Value = ""
Me.ed_txtBox_custoInsumo.Value = ""
Me.ed_txtBox_pvs.Value = ""
Me.ed_txtBox_cmo.Value = ""

End Sub



Private Sub ed_txtBox_nv1_1_Change()

End Sub

Private Sub ed_txtBox_nv4_4_Click()

End Sub

Private Sub ed_txtComboBox_nv1_1_Change()
Edição.ed_GetIdNv1
ed_txtBox_nv1_1 = ed_txtComboBox_nv1_1.Value

If ed_btnCmoPvsBoolean = True Then
checkAll
End If

If ed_btnRendimentoBoolean = True And ed_sPrincipais = True Then
    If ed_txtComboBox_nv1_1 = "" And ed_txtComboBox_nv2_2 = "" And ed_txtComboBox_nv3_3 = "" And ed_txtComboBox_nv4_4 = "" Then
    Edição.ed_GetRendimento
    End If
    ed_verify_estructure
    ed_carregarEstruturaPrincipais2
    
    
    UserForm2.ed_txtComboBox_nv3_3.Enabled = False
    UserForm2.ed_txtComboBox_nv3_3.BackColor = &H8000000F
    UserForm2.ed_txtComboBox_nv3_3.Clear
    UserForm2.ed_txtComboBox_nv4_4.Enabled = False
    UserForm2.ed_txtComboBox_nv4_4.BackColor = &H8000000F
    UserForm2.ed_txtComboBox_nv4_4.Clear
    UserForm2.ed_txtBox_rendimento.Enabled = False
    UserForm2.ed_txtBox_rendimento.BackColor = &H8000000F
    UserForm2.ed_txtBox_rendimento = ""
    
End If

If ed_btnCmoPvsBoolean = True And ed_sPrincipais = True Then
    If ed_txtComboBox_nv1_1 = "" And ed_txtComboBox_nv2_2 = "" And ed_txtComboBox_nv3_3 = "" And ed_txtComboBox_nv4_4 = "" Then
    Edição.ed_GetRendimento
    End If
    ed_verify_estructure
    ed_carregarEstruturaPrincipais2
    
    
    UserForm2.ed_txtComboBox_nv3_3.Enabled = False
    UserForm2.ed_txtComboBox_nv3_3.BackColor = &H8000000F
    UserForm2.ed_txtComboBox_nv3_3.Clear

    UserForm2.ed_txtBox_pvs.Enabled = False
    UserForm2.ed_txtBox_pvs.BackColor = &H8000000F
    UserForm2.ed_txtBox_pvs = ""
    UserForm2.ed_txtBox_cmo.Enabled = False
    UserForm2.ed_txtBox_cmo.BackColor = &H8000000F
    UserForm2.ed_txtBox_cmo = ""
    
End If



End Sub

Sub checkAll()

If ed_txtComboBox_nv1_1 <> "" And ed_txtComboBox_nv2_2 <> "" And ed_txtComboBox_nv3_3 <> "" Then
    ed_GetPvs_Cmo
End If



End Sub

Private Sub ed_txtComboBox_nv2_2_Change()
ed_GetIdNv2
ed_txtBox_nv2_2 = ed_txtComboBox_nv2_2.Value

If ed_btnCmoPvsBoolean = True Then
checkAll
End If

If ed_btnRendimentoBoolean = True And ed_sPrincipais = True Then
    ed_verify_estructure
    ed_carregarEstruturaPrincipais3
End If

If Me.ed_btnCmoPvsBoolean = True And ed_sPrincipais = True Then
    ed_verify_estructure
    ed_carregarEstruturaPrincipais3
End If

End Sub

Private Sub ed_txtComboBox_nv3_3_Change()
ed_GetIdNv3
ed_txtBox_nv3_3 = ed_txtComboBox_nv3_3.Value

If ed_btnCmoPvsBoolean = True Then
checkAll
End If

If ed_btnRendimentoBoolean = True And ed_sPrincipais = True Then
    If ed_txtComboBox_nv1_1 = "" And ed_txtComboBox_nv2_2 = "" And ed_txtComboBox_nv3_3 = "" And ed_txtComboBox_nv4_4 = "" Then
    Edição.ed_GetRendimento
    End If
End If

If ed_sDiversos = True And ed_btnCmoPvsBoolean = True Then
Edição.ed_GetPvs_Cmo
End If

If ed_btnRendimentoBoolean = True And ed_sDiversos = True Then
    If ed_txtComboBox_nv3_3 <> "" And ed_txtComboBox_nv4_4 <> "" Then
    Edição.ed_GetRendimentoDiversos
    End If
    ed_verify_estructure
    ed_carregarEstruturaPrincipais4
End If

If ed_btnCmoPvsBoolean = True And ed_sPrincipais = True Then
    ed_verify_estructure
End If

If ed_btnRendimentoBoolean = True And ed_sPrincipais = True Then
    ed_verify_estructure
    ed_carregarEstruturaPrincipais4
End If


End Sub

Private Sub ed_txtComboBox_nv4_4_Change()
Edição.ed_GetIdInsumo
ed_txtBox_nv4_4.Enabled = True
ed_txtBox_nv4_4 = ed_txtComboBox_nv4_4.Value

ed_txtBox_un.Enabled = True
ed_txtBox_un.BackColor = RGB(255, 255, 255)
ed_txtBox_custoInsumo.Enabled = True
ed_txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
Me.ed_ComboBoxTipo.Enabled = True
Me.ed_ComboBoxTipo.BackColor = RGB(255, 255, 255)

If ed_btnCmoPvsBoolean = True Then
checkAll
End If

If ed_btnNv4Boolean = True Then
Edição.ed_GetUnidade
Edição.ed_GetCustoInsumo
Edição.ed_GetTipoInsumo
End If

If ed_btnRendimentoBoolean = True And ed_sPrincipais = True Then
    If ed_txtComboBox_nv1_1 <> "" And ed_txtComboBox_nv2_2 <> "" And ed_txtComboBox_nv3_3 <> "" And ed_txtComboBox_nv4_4 <> "" Then
    Edição.ed_GetRendimento
    End If
End If

If ed_btnRendimentoBoolean = True And ed_sDiversos = True Then
    If ed_txtComboBox_nv3_3 <> "" And ed_txtComboBox_nv4_4 <> "" Then
    Edição.ed_GetRendimentoDiversos
    End If
End If

If ed_btnRendimentoBoolean = True And ed_sPrincipais = True Then
    ed_verify_estructure
    'ed_carregarRendimento
End If

End Sub

Private Sub Editar_Click()
Edição.ed_Enviar
End Sub

Private Sub EditarNiveis_Click()

   ' Declarando a variável para armazenar o valor de retorno
    Dim resposta As VbMsgBoxResult
    
    ' Exibindo a caixa de diálogo com os botões "OK" e "Cancelar"
    resposta = MsgBox("Os dados de serviço serão perdidos. Deseja continuar?", vbQuestion + vbOKCancel, "Confirmação")
    
    ' Verificando qual botão foi clicado
    If resposta = vbOK Then
    naoMostrarCampos
    Else
        ' Código a ser executado quando o botão "Cancelar" for clicado
    End If






End Sub

Private Sub Label48_Click()

End Sub

Private Sub ex_btnP_Click()
Me.ex_btnNv1.Enabled = True
Me.ex_btnNv2.Enabled = True
Me.ex_btnNv3.Enabled = True
Me.ex_btnNv4.Enabled = True
Me.ex_btnEstrutura.Enabled = True





UserForm2.ed_btnCmoPvs.Caption = "Editar Custo Mão de Obra | Preço de Venda Sugerido"

ed_sPrincipais = True
ed_sDiversos = False
ed_sTerceiros = False


Me.ed_txtBoxID_nv1 = ""
Me.ed_txtBoxID_nv2 = ""
Me.ed_txtBoxID_nv3 = ""
Me.ed_txtBoxID_nv4 = ""

Me.ex_s_principais.BackColor = RGB(142, 162, 219)
Me.ex_s_terceiros.ForeColor = &H80000008
Me.ex_s_principais.Font.Bold = True



Me.ex_s_diversos.BackColor = &H8000000F
Me.ex_s_diversos.ForeColor = &H80000008
Me.ex_s_diversos.Font.Bold = False

Me.ex_s_terceiros.BackColor = &H8000000F
Me.ex_s_terceiros.ForeColor = &H80000008
Me.ex_s_terceiros.Font.Bold = False

Me.ed_txtBox_nv1_1.Value = ""
Me.ed_txtComboBox_nv1_1.Value = ""
Me.ed_txtBox_nv2_2.Value = ""
Me.ed_txtComboBox_nv2_2.Value = ""
Me.ed_txtBox_nv3_3.Value = ""
Me.ed_txtComboBox_nv3_3.Value = ""
Me.ed_txtBox_nv4_4.Value = ""
Me.ed_txtComboBox_nv4_4.Value = ""
Me.ed_txtBox_un.Value = ""
Me.ed_txtBox_rendimento.Value = ""
Me.ed_txtBox_custoInsumo.Value = ""
Me.ed_txtBox_pvs.Value = ""
Me.ed_txtBox_cmo.Value = ""
End Sub

Private Sub ex_btnEstrutura_Click()
Me.MultiPage4.Value = 1
Me.MultiPage4(1).Visible = True
Me.MultiPage4(2).Visible = False
Me.MultiPage4(0).Visible = False

If ex_sPrincipais = True Then
ex_titleEstrutura.Caption = "Serviço Principal"
ex_titleNivel.Caption = "Serviço Principal"
ElseIf ex_sDiversos = True Then
ex_titleEstrutura.Caption = "Serviço Diverso"
ex_titleNivel.Caption = "Serviço Diverso"
ElseIf ex_sTerceiros = True Then
ex_titleEstrutura.Caption = "Serviço de Terceiros"
ex_titleNivel.Caption = "Serviço de Terceiros"
End If

If Me.ex_sPrincipais = True Then
ex_ComboBoxNv4.Visible = False
ex_titleEstrutura = "Exclusão da Estrutura de Serviços Principais"
Else
ex_ComboBoxNv4.Visible = True
End If

If Me.ex_sGeneralInsumo = True Then
ex_titleEstrutura = "Exclusão da Estrutura de Insumo"
ex_BtnSelectionPrincipais_Click
End If

If ex_sPrincipais = True Then
ex_BtnSelectionPrincipais.Visible = False
ex_BtnSelectionDiversos.Visible = False
Else
ex_BtnSelectionPrincipais.Visible = True
ex_BtnSelectionDiversos.Visible = True
ex_ComboBoxNv4.Visible = True
ex_ComboBoxNv4.Enabled = False
ex_ComboBoxNv4.BackColor = &H8000000F
ex_estruturaNv4Title.Visible = True

End If

ex_ComboBoxNv1.Enabled = True
ex_ComboBoxNv1.BackColor = RGB(255, 255, 255)
ex_ComboBoxNv2.Enabled = False
ex_ComboBoxNv2.BackColor = &H8000000F
ex_ComboBoxNv3.Enabled = False
ex_ComboBoxNv3.BackColor = &H8000000F
If ex_sPrincipais = True Then
ex_ComboBoxNv4.Visible = False
ex_ComboBoxNv4.Enabled = False
ex_ComboBoxNv4.BackColor = &H8000000F
ex_estruturaNv4Title.Visible = False
End If

ex_btnNv1BooleanP = False
ex_btnNv2BooleanP = False
ex_btnNv3BooleanP = False
ex_btnNv4BooleanP = False
ex_btnNv3BooleanD = False
ex_btnEstruturaBoolean = True

'ex_CarregarNivelPrincipal
carregarEstruturaPrincipais
'ex_ListBox.Clear




End Sub

Sub verify_estructure()


    
If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
    If ex_ComboBoxNv1 <> "" Then
        ex_ComboBoxNv2.Enabled = True
        ex_ComboBoxNv2.BackColor = RGB(255, 255, 255)
            If ex_ComboBoxNv2.Enabled = True And ex_ComboBoxNv2.Value <> "" Then
            ex_ComboBoxNv3.Enabled = True
            ex_ComboBoxNv3.BackColor = RGB(255, 255, 255)
            
            Else
            ex_ComboBoxNv3.Enabled = False
            ex_ComboBoxNv3.BackColor = &H8000000F
            
            ex_ComboBoxNv4.Enabled = False
            ex_ComboBoxNv4.BackColor = &H8000000F
            End If
        Else
            ex_ComboBoxNv2.Enabled = False
            ex_ComboBoxNv2.BackColor = &H8000000F
        End If
        If ex_ComboBoxNv1 <> "" And ex_ComboBoxNv2 <> "" And ex_ComboBoxNv3 <> "" Then
        ex_ComboBoxNv4.Enabled = True
        ex_ComboBoxNv4.Clear
        ex_ComboBoxNv4.BackColor = RGB(255, 255, 255)
    End If
ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then

    If ex_ComboBoxNv3.Enabled = True And ex_ComboBoxNv3.Value <> "" Then
        ex_ComboBoxNv4.Enabled = True
        ex_ComboBoxNv4.Clear
        ex_ComboBoxNv4.BackColor = RGB(255, 255, 255)
    Else
        ex_ComboBoxNv4.Enabled = False
        ex_ComboBoxNv4.Clear
        ex_ComboBoxNv4.BackColor = &H8000000F
    End If
    
End If
    
End Sub

Private Sub ex_btnNv1_Click()
Dim resposta As VbMsgBoxResult

If GlobalServiceType = "Servicos Principais" And GlobalTable = "TabelaNv1" Or GlobalTable = "" Then

    'UserForm2.ed_btnCmoPvs.Caption = "Editar Custo Mão de Obra"
    Me.MultiPage4(1).Visible = False
    Me.MultiPage4(2).Visible = True
    Me.MultiPage4(0).Visible = False
    Me.MultiPage4.Value = 2
    ex_dinamicTitle.Caption = "NÍVEL 1"
    
    If ex_sPrincipais = True Then
    ex_titleEstrutura.Caption = "Serviços Principais"
    ex_titleNivel.Caption = "Serviços Principais"
    ElseIf ex_sDiversos = True Then
    ex_titleEstrutura.Caption = "Serviços Diversos"
    ex_titleNivel.Caption = "Serviços Diversos"
    ElseIf ex_sTerceiros = True Then
    ex_titleEstrutura.Caption = "Serviços de Terceiros"
    ex_titleNivel.Caption = "Serviços de Terceiros"
    End If
    
    
    ex_btnNv1BooleanP = True
    ex_btnNv2BooleanP = False
    ex_btnNv3BooleanP = False
    ex_btnNv4BooleanP = False
    ex_btnNv3BooleanD = False
    ex_btnEstruturaBoolean = False
    
    ex_CarregarNivelPrincipal
'    ex_ListBox.Clear

End If

If GlobalServiceType = "Servicos Principais" And GlobalTable = "TabelaNv2" Or GlobalTable = "TabelaNv3" Then
    resposta = MsgBox("Os dados da última lista serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then
    GlobalTable = ""
    GlobalComboBoxValue = ""
    
        'UserForm2.ed_btnCmoPvs.Caption = "Editar Custo Mão de Obra"
        Me.MultiPage4(1).Visible = False
        Me.MultiPage4(2).Visible = True
        Me.MultiPage4(0).Visible = False
        Me.MultiPage4.Value = 2
        ex_dinamicTitle.Caption = "NÍVEL 1"
        
        If ex_sPrincipais = True Then
        ex_titleEstrutura.Caption = "Serviços Principais"
        ex_titleNivel.Caption = "Serviços Principais"
        ElseIf ex_sDiversos = True Then
        ex_titleEstrutura.Caption = "Serviços Diversos"
        ex_titleNivel.Caption = "Serviços Diversos"
        ElseIf ex_sTerceiros = True Then
        ex_titleEstrutura.Caption = "Serviços de Terceiros"
        ex_titleNivel.Caption = "Serviços de Terceiros"
        End If
        
        
        ex_btnNv1BooleanP = True
        ex_btnNv2BooleanP = False
        ex_btnNv3BooleanP = False
        ex_btnNv4BooleanP = False
        ex_btnNv3BooleanD = False
        ex_btnEstruturaBoolean = False
        
        ex_CarregarNivelPrincipal
        ex_ListBox.Clear
        
        GlobalTable = ""
        GlobalComboBoxValue = ""
    
    Else
    
        Exit Sub
    
    End If
End If

End Sub

Private Sub ex_btnNv2_Click()



Dim resposta As VbMsgBoxResult

If GlobalServiceType = "Servicos Principais" And GlobalTable = "TabelaNv2" Or GlobalTable = "" Or UserForm2.ex_ListBox = "" Then

'If UserForm2.ex_ListBox = Empty And GlobalTable = "TabelaNv3" Or "TabelaNv1" Then
'GlobalTable = ""
'GlobalComboBoxValue = ""
'End If
Me.MultiPage4(1).Visible = False
Me.MultiPage4(2).Visible = True
Me.MultiPage4(0).Visible = False
Me.MultiPage4.Value = 2
ex_dinamicTitle.Caption = "NÍVEL 2"

If ex_sPrincipais = True Then
ex_titleEstrutura.Caption = "Serviços Principais"
ex_titleNivel.Caption = "Serviços Principais"
ElseIf ex_sDiversos = True Then
ex_titleEstrutura.Caption = "Serviços Diversos"
ex_titleNivel.Caption = "Serviços Diversos"
ElseIf ex_sTerceiros = True Then
ex_titleEstrutura.Caption = "Serviços de Terceiros"
ex_titleNivel.Caption = "Serviços de Terceiros"
End If

ex_btnNv1BooleanP = False
ex_btnNv2BooleanP = True
ex_btnNv3BooleanP = False
ex_btnNv4BooleanP = False
ex_btnNv3BooleanD = False
ex_btnEstruturaBoolean = False

ex_CarregarNivelPrincipal

End If

If GlobalServiceType = "Servicos Principais" And GlobalTable = "TabelaNv3" Or GlobalTable = "TabelaNv1" Then
    resposta = MsgBox("Os dados da última lista serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then

        Me.MultiPage4(1).Visible = False
        Me.MultiPage4(2).Visible = True
        Me.MultiPage4(0).Visible = False
        Me.MultiPage4.Value = 2
        ex_dinamicTitle.Caption = "NÍVEL 2"
        
        If ex_sPrincipais = True Then
        ex_titleEstrutura.Caption = "Serviços Principais"
        ex_titleNivel.Caption = "Serviços Principais"
        ElseIf ex_sDiversos = True Then
        ex_titleEstrutura.Caption = "Serviços Diversos"
        ex_titleNivel.Caption = "Serviços Diversos"
        ElseIf ex_sTerceiros = True Then
        ex_titleEstrutura.Caption = "Serviços de Terceiros"
        ex_titleNivel.Caption = "Serviços de Terceiros"
        End If
        
        ex_btnNv1BooleanP = False
        ex_btnNv2BooleanP = True
        ex_btnNv3BooleanP = False
        ex_btnNv4BooleanP = False
        ex_btnNv3BooleanD = False
        ex_btnEstruturaBoolean = False
        
        ex_CarregarNivelPrincipal
        ex_ListBox.Clear
    
        GlobalTable = ""
        GlobalComboBoxValue = ""
    Else
    
        Exit Sub
    
    End If
End If
End Sub

Private Sub ex_btnNv3_Click()

If ex_sDiversos = True Or ex_sTerceiros = True Then
        Me.MultiPage4(1).Visible = False
        Me.MultiPage4(2).Visible = True
        Me.MultiPage4(0).Visible = False
        Me.MultiPage4.Value = 2
        ex_dinamicTitle.Caption = "SERVIÇO"
        Me.ex_dinamicTitle.FontSize = 14
        If ex_sPrincipais = True Then
        ex_titleEstrutura.Caption = "Serviços Principais"
        ex_titleNivel.Caption = "Serviços Principais"
        ElseIf ex_sDiversos = True Then
        ex_titleEstrutura.Caption = "Serviços Diversos"
        ex_titleNivel.Caption = "Serviços Diversos"
        ElseIf ex_sTerceiros = True Then
        ex_titleEstrutura.Caption = "Serviços de Terceiros"
        ex_titleNivel.Caption = "Exclusão de Serviços de Terceiros"
        End If
        
        
        ex_btnNv1BooleanP = False
        ex_btnNv2BooleanP = False
        ex_btnNv3BooleanP = True
        ex_btnNv3BooleanD = False
        ex_btnNv4BooleanP = False
        If ex_sDiversos = True Or ex_sTerceiros = True Then
        ex_btnNv3BooleanD = True
        ex_btnNv3BooleanP = False
        End If
        ex_btnEstruturaBoolean = False
        
        
        
        If ex_sPrincipais = True Then
        ex_CarregarNivelPrincipal
        ElseIf ex_sDiversos = True Then
        ex_CarregarNivelTres
        ElseIf ex_sTerceiros = True Then
        ex_CarregarNivelTres
        End If
        ex_ListBox.Clear
        Exit Sub
End If

If GlobalServiceType = "Servicos Principais" And GlobalTable = "TabelaNv3" Or GlobalTable = "" Or UserForm2.ex_ListBox = "" Then

        Me.MultiPage4(1).Visible = False
        Me.MultiPage4(2).Visible = True
        Me.MultiPage4(0).Visible = False
        Me.MultiPage4.Value = 2
        ex_dinamicTitle.Caption = "SERVIÇO"
        Me.ex_dinamicTitle.FontSize = 14
        If ex_sPrincipais = True Then
        ex_titleEstrutura.Caption = "Serviços Principais"
        ex_titleNivel.Caption = "Serviços Principais"
        ElseIf ex_sDiversos = True Then
        ex_titleEstrutura.Caption = "Serviços Diversos"
        ex_titleNivel.Caption = "Serviços Diversos"
        ElseIf ex_sTerceiros = True Then
        ex_titleEstrutura.Caption = "Serviços de Terceiros"
        ex_titleNivel.Caption = "Exclusão de Serviços de Terceiros"
        End If
        
        
        ex_btnNv1BooleanP = False
        ex_btnNv2BooleanP = False
        ex_btnNv3BooleanP = True
        ex_btnNv3BooleanD = False
        ex_btnNv4BooleanP = False
        If ex_sDiversos = True Or ex_sTerceiros = True Then
        ex_btnNv3BooleanD = True
        ex_btnNv3BooleanP = False
        End If
        ex_btnEstruturaBoolean = False
        
        
        
        If ex_sPrincipais = True Then
        ex_CarregarNivelPrincipal
        ElseIf ex_sDiversos = True Then
        ex_CarregarNivelTres
        ElseIf ex_sTerceiros = True Then
        ex_CarregarNivelTres
        End If
        'ex_ListBox.Clear


End If

If GlobalServiceType = "Servicos Principais" And GlobalTable = "TabelaNv2" Or GlobalTable = "TabelaNv1" Then
    resposta = MsgBox("Os dados da última lista serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then
        
        Me.MultiPage4(1).Visible = False
        Me.MultiPage4(2).Visible = True
        Me.MultiPage4(0).Visible = False
        Me.MultiPage4.Value = 2
        ex_dinamicTitle.Caption = "SERVIÇO"
        Me.ex_dinamicTitle.FontSize = 14
        If ex_sPrincipais = True Then
        ex_titleEstrutura.Caption = "Serviços Principais"
        ex_titleNivel.Caption = "Serviços Principais"
        ElseIf ex_sDiversos = True Then
        ex_titleEstrutura.Caption = "Serviços Diversos"
        ex_titleNivel.Caption = "Serviços Diversos"
        ElseIf ex_sTerceiros = True Then
        ex_titleEstrutura.Caption = "Serviços de Terceiros"
        ex_titleNivel.Caption = "Exclusão de Serviços de Terceiros"
        End If
        
        
        ex_btnNv1BooleanP = False
        ex_btnNv2BooleanP = False
        ex_btnNv3BooleanP = True
        ex_btnNv3BooleanD = False
        ex_btnNv4BooleanP = False
        If ex_sDiversos = True Or ex_sTerceiros = True Then
        ex_btnNv3BooleanD = True
        ex_btnNv3BooleanP = False
        End If
        ex_btnEstruturaBoolean = False
        
        
        
        If ex_sPrincipais = True Then
        ex_CarregarNivelPrincipal
        ElseIf ex_sDiversos = True Then
        ex_CarregarNivelTres
        ElseIf ex_sTerceiros = True Then
        ex_CarregarNivelTres
        End If
        ex_ListBox.Clear

    
        GlobalTable = ""
        GlobalComboBoxValue = ""
    Else
    
        Exit Sub
    
    End If
End If

End Sub

Private Sub ex_btnNv4_Click()
Me.MultiPage4(1).Visible = False
Me.MultiPage4(2).Visible = True
Me.MultiPage4(0).Visible = False
Me.MultiPage4.Value = 2
ex_dinamicTitle.Caption = "NÍVEL INSUMO"
'    ex_dinamicTitle.MultiLine = True
Me.ex_dinamicTitle.FontSize = 14

If ex_sPrincipais = True Then
ex_titleEstrutura.Caption = "Serviços Principais"
ex_titleNivel.Caption = "Serviços Principais"
ElseIf ex_sDiversos = True Then
ex_titleEstrutura.Caption = "Serviços Diversos"
ex_titleNivel.Caption = "Serviços Diversos"
ElseIf ex_sTerceiros = True Then
ex_titleEstrutura.Caption = "Serviços de Terceiros"
ex_titleNivel.Caption = "Serviços de Terceiros"
ElseIf ex_sGeneralInsumo = True Then
ex_titleNivel.Caption = "Exclusão de Insumo"
End If



ex_btnNv1BooleanP = False
ex_btnNv2BooleanP = False
ex_btnNv3BooleanP = False
ex_btnNv3BooleanD = False
ex_btnNv4BooleanP = True
ex_btnEstruturaBoolean = False



If ex_sPrincipais = True Then
ex_CarregarNivelPrincipal
ElseIf ex_sDiversos = True Then
ex_CarregarNivelTres
ElseIf ex_sTerceiros = True Then
ex_CarregarNivelTres
ElseIf Me.ex_sGeneralInsumo Then
ex_CarregarNivelPrincipal
End If


'ex_ListBox.Clear

End Sub

Private Sub ex_ComboBoxNiveis_Change()



End Sub

Private Sub ex_ComboBoxNv1_Change()
If ex_sPrincipais = True Then
verify_estructure

ex_ComboBoxNv2.Clear
ex_ComboBoxNv3.Clear
ex_ComboBoxNv3.Enabled = False
ex_ComboBoxNv3.BackColor = &H8000000F
'Exclusão.carregarEstruturaPrincipais
Exclusão.carregarEstruturaPrincipais2

ElseIf Me.ex_BtnSelectionPrincipaisBoolean = True Then
verify_estructure

ex_ComboBoxNv2.Clear
ex_ComboBoxNv3.Clear
ex_ComboBoxNv3.Enabled = False
ex_ComboBoxNv3.BackColor = &H8000000F
'Exclusão.carregarEstruturaPrincipais
Exclusão.carregarEstruturaPrincipais2
End If




End Sub

Private Sub ex_ComboBoxNv2_Change()
If ex_sPrincipais = True Then
verify_estructure
Exclusão.carregarEstruturaPrincipais3
    

ElseIf ex_BtnSelectionPrincipaisBoolean = True Then
verify_estructure
Exclusão.carregarEstruturaPrincipais3
End If

If Me.ex_sGeneralInsumo = True Then
verify_estructure
End If
End Sub

Private Sub ex_ComboBoxNv3_Change()
'verify_estructure
If Me.ex_sGeneralInsumo = True Then
verify_estructure

End If

If ex_BtnSelectionPrincipaisBoolean = True And ex_sGeneralInsumo = True Then
verify_estructure
Exclusão.carregarEstruturaPrincipais4
ElseIf ex_BtnSelectionDiversosBoolean = True And ex_sGeneralInsumo = True Then
verify_estructure
Exclusão.carregarEstruturaPrincipais4
End If

End Sub

Private Sub ex_OptionButtonDiversos_Click()
ex_ComboBoxNv1.Enabled = True
ex_ComboBoxNv2.Enabled = False
Me.ex_ComboBoxNv2.BackColor = &H8000000F
ex_ComboBoxNv3.Enabled = True
ex_ComboBoxNv3.BackColor = RGB(255, 255, 255)



ex_ComboBoxNv1.Clear
ex_ComboBoxNv2.Clear
ex_ComboBoxNv3.Clear

Exclusão.ex_CarregarNivelPrincipal



End Sub

Private Sub ex_OptionButtonPrincipais_Click()
ex_ComboBoxNv1.Enabled = True
    
ex_ComboBoxNv2.Enabled = False
ex_ComboBoxNv2.BackColor = &H8000000F
ex_ComboBoxNv3.Enabled = False
ex_ComboBoxNv3.BackColor = &H8000000F
ex_ComboBoxNv4.Visible = True
ex_ComboBoxNv4.Enabled = False
ex_ComboBoxNv4.BackColor = &H8000000F

ex_CarregarNivelPrincipal

End Sub

Private Sub ex_ListBox_Click()

End Sub

Private Sub ex_s_diversos_Click()

Dim resposta As VbMsgBoxResult


''\!/---BLOQUEIA O ACESSO DO USUÁRIO---\!/
'stopApplication
'If stop_Application Then Exit Sub
''\!/----------------------------------\!/


If GlobalServiceType = "Servicos de Terceiros" Or GlobalServiceType = "Servicos Diversos" Then
GlobalTable = ""
GlobalComboBoxValue = ""


End If


If GlobalServiceType <> "" And GlobalTable = "" And GlobalComboBoxValue = "" Then

    ex_sPrincipais = False
    ex_sDiversos = True
    ex_sTerceiros = False
    ex_sGeneralInsumo = False
    
    Me.ex_btnNv1.Enabled = False
    Me.ex_btnNv1.BackColor = &H8000000F
    Me.ex_btnNv2.Enabled = False
    Me.ex_btnNv2.BackColor = &H8000000F
    Me.ex_btnNv3.Enabled = True
    Me.ex_btnNv4.Enabled = False
    Me.ex_btnNv4.BackColor = &H8000000F
    ex_btnEstrutura.BackColor = &H8000000F
    ex_btnEstrutura.Enabled = False

    
    
    ex_ListBox.Clear

    
    Me.ex_s_diversos.BackColor = RGB(255, 230, 153)
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = True

    
    Me.ex_s_principais.BackColor = &H8000000F
    Me.ex_s_principais.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = False
    
    Me.ex_s_terceiros.BackColor = &H8000000F
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = &H8000000F
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = False
    

    
    GlobalComboBoxValue = ""
    GlobalServiceType = "Servicos Diversos"
    GlobalTable = "TabelaNv3"


    ex_ListBox.Clear
'    GlobalTable = ""
'    GlobalComboBoxValue = ""

End If

If GlobalServiceType = "Insumos" And GlobalTable = "TabelaNv4" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos de Terceiros" And GlobalTable = "TabelaNv3" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos Principais" And GlobalTable <> "" And GlobalComboBoxValue <> "" Then
    resposta = MsgBox("Os dados da última lista gerada serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then

    
    ex_sPrincipais = False
    ex_sDiversos = True
    ex_sTerceiros = False
    ex_sGeneralInsumo = False
    
    Me.ex_btnNv1.Enabled = False
    Me.ex_btnNv1.BackColor = &H8000000F
    Me.ex_btnNv2.Enabled = False
    Me.ex_btnNv2.BackColor = &H8000000F
    Me.ex_btnNv3.Enabled = True
    Me.ex_btnNv4.Enabled = False
    Me.ex_btnNv4.BackColor = &H8000000F
    ex_btnEstrutura.BackColor = &H8000000F
    ex_btnEstrutura.Enabled = False
    
    
    ex_ListBox.Clear

    
    Me.ex_s_diversos.BackColor = RGB(255, 230, 153)
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = True

    
    Me.ex_s_principais.BackColor = &H8000000F
    Me.ex_s_principais.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = False
    
    Me.ex_s_terceiros.BackColor = &H8000000F
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = &H8000000F
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = False
    

    
    GlobalComboBoxValue = ""
    GlobalServiceType = "Servicos Diversos"
    GlobalTable = "TabelaNv3"


    ex_ListBox.Clear
'    GlobalTable = ""
'    GlobalComboBoxValue = ""
    
    Else
    
        Exit Sub
    
    End If
End If



End Sub



Private Sub ex_s_principais_Click()


Dim resposta As VbMsgBoxResult

If GlobalServiceType <> "" And GlobalTable = "" Or GlobalComboBoxValue = "" Then

    Me.ex_btnNv1.Enabled = True
    Me.ex_btnNv2.Enabled = True
    Me.ex_btnNv3.Enabled = True
    Me.ex_btnNv4.Enabled = False
    Me.ex_btnNv4.BackColor = &H8000000F
    Me.ex_btnEstrutura.Enabled = True
    
    ex_ListBox.Clear
    

    
    ex_sPrincipais = True
    ex_sDiversos = False
    ex_sTerceiros = False
    ex_sGeneralInsumo = False
    

    Me.ex_s_principais.BackColor = RGB(142, 162, 219)
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = True

    
    
    Me.ex_s_diversos.BackColor = &H8000000F
    Me.ex_s_diversos.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = False
    
    Me.ex_s_terceiros.BackColor = &H8000000F
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = &H8000000F
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = False
    
    Me.ed_txtBox_nv1_1.Value = ""
    Me.ed_txtComboBox_nv1_1.Value = ""
    Me.ed_txtBox_nv2_2.Value = ""
    Me.ed_txtComboBox_nv2_2.Value = ""
    Me.ed_txtBox_nv3_3.Value = ""
    Me.ed_txtComboBox_nv3_3.Value = ""
    Me.ed_txtBox_nv4_4.Value = ""
    Me.ed_txtComboBox_nv4_4.Value = ""
    Me.ed_txtBox_un.Value = ""
    Me.ed_txtBox_rendimento.Value = ""
    Me.ed_txtBox_custoInsumo.Value = ""
    Me.ed_txtBox_pvs.Value = ""
    Me.ed_txtBox_cmo.Value = ""
    
    
    GlobalComboBoxValue = ""
    GlobalServiceType = "Servicos Principais"
    GlobalTable = ""

End If

If GlobalServiceType = "Insumos" And GlobalTable = "TabelaNv4" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos de Terceiros" And GlobalTable = "TabelaNv3" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos Diversos" And GlobalTable = "TabelaNv3" And GlobalComboBoxValue <> "" Then
    resposta = MsgBox("Os dados da última lista gerada serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then

    
    Me.ex_btnNv1.Enabled = True
    Me.ex_btnNv2.Enabled = True
    Me.ex_btnNv3.Enabled = True
    Me.ex_btnNv4.Enabled = False
    Me.ex_btnNv4.BackColor = &H8000000F
    Me.ex_btnEstrutura.Enabled = True
    
    ex_ListBox.Clear
    
    

    
    ex_sPrincipais = True
    ex_sDiversos = False
    ex_sTerceiros = False
    ex_sGeneralInsumo = False
    

    Me.ex_s_principais.BackColor = RGB(142, 162, 219)
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = True

    
    
    Me.ex_s_diversos.BackColor = &H8000000F
    Me.ex_s_diversos.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = False
    
    Me.ex_s_terceiros.BackColor = &H8000000F
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = &H8000000F
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = False
    
    Me.ed_txtBox_nv1_1.Value = ""
    Me.ed_txtComboBox_nv1_1.Value = ""
    Me.ed_txtBox_nv2_2.Value = ""
    Me.ed_txtComboBox_nv2_2.Value = ""
    Me.ed_txtBox_nv3_3.Value = ""
    Me.ed_txtComboBox_nv3_3.Value = ""
    Me.ed_txtBox_nv4_4.Value = ""
    Me.ed_txtComboBox_nv4_4.Value = ""
    Me.ed_txtBox_un.Value = ""
    Me.ed_txtBox_rendimento.Value = ""
    Me.ed_txtBox_custoInsumo.Value = ""
    Me.ed_txtBox_pvs.Value = ""
    Me.ed_txtBox_cmo.Value = ""
    
    
    GlobalComboBoxValue = ""
    GlobalServiceType = "Servicos Principais"
    GlobalTable = ""


    ex_ListBox.Clear
    
    GlobalTable = ""
    GlobalComboBoxValue = ""
    
    Else
    
        Exit Sub
    
    End If
End If


End Sub

Private Sub ex_s_terceiros_Click()



Dim resposta As VbMsgBoxResult

If GlobalServiceType = "Servicos de Terceiros" Or GlobalServiceType = "Servicos Diversos" Then
GlobalTable = ""
GlobalComboBoxValue = ""
End If


If GlobalServiceType <> "" And GlobalTable = "" And GlobalComboBoxValue = "" Then

    ex_sPrincipais = False
    ex_sDiversos = False
    ex_sTerceiros = True
    ex_sGeneralInsumo = False
    
    Me.ex_btnNv1.Enabled = False
    Me.ex_btnNv2.Enabled = False
    Me.ex_btnNv3.Enabled = True
    Me.ex_btnNv4.Enabled = False
    Me.ex_btnNv4.BackColor = &H8000000F
    ex_btnEstrutura.BackColor = &H8000000F
    ex_btnEstrutura.Enabled = False
    
    ex_ListBox.Clear

    
    Me.ex_s_terceiros.BackColor = RGB(255, 242, 204)
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = True

    
    Me.ex_s_principais.BackColor = &H8000000F
    Me.ex_s_principais.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = False
    
    Me.ex_s_diversos.BackColor = &H8000000F
    Me.ex_s_diversos.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = &H8000000F
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = False
    
    GlobalComboBoxValue = ""
    GlobalServiceType = "Servicos de Terceiros"
    GlobalTable = "TabelaNv3"



End If

If GlobalServiceType = "Insumos" And GlobalTable = "TabelaNv4" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos Diversos" And GlobalTable = "TabelaNv3" And GlobalComboBoxValue <> "" Or GlobalServiceType = "Servicos Principais" And GlobalTable <> "" And GlobalComboBoxValue <> "" Then
    resposta = MsgBox("Os dados da última lista gerada serão perdidos. Deseja prosseguir?", vbQuestion + vbYesNo, "Confirmação")
    
    If resposta = vbYes Then

    
    ex_sPrincipais = False
    ex_sDiversos = False
    ex_sTerceiros = True
    ex_sGeneralInsumo = False
    
    Me.ex_btnNv1.Enabled = False
    Me.ex_btnNv2.Enabled = False
    Me.ex_btnNv3.Enabled = True
    Me.ex_btnNv4.Enabled = False
    Me.ex_btnNv4.BackColor = &H8000000F
    ex_btnEstrutura.BackColor = &H8000000F
    ex_btnEstrutura.Enabled = False
    
    ex_ListBox.Clear

    
    Me.ex_s_terceiros.BackColor = RGB(255, 242, 204)
    Me.ex_s_terceiros.ForeColor = &H80000008
    Me.ex_s_terceiros.Font.Bold = True
    'Me.MultiPage2.Value = 2
    'Me.MultiPage1.BackColor = RGB(255, 249, 231)
    
    Me.ex_s_principais.BackColor = &H8000000F
    Me.ex_s_principais.ForeColor = &H80000008
    Me.ex_s_principais.Font.Bold = False
    
    Me.ex_s_diversos.BackColor = &H8000000F
    Me.ex_s_diversos.ForeColor = &H80000008
    Me.ex_s_diversos.Font.Bold = False
    
    Me.ex_insumoGerenal.BackColor = &H8000000F
    Me.ex_insumoGerenal.ForeColor = &H80000008
    Me.ex_insumoGerenal.Font.Bold = False
    
    GlobalComboBoxValue = ""
    GlobalServiceType = "Servicos de Terceiros"
    GlobalTable = "TabelaNv3"

    
    Else
    
        Exit Sub
    
    End If
End If





End Sub

Private Sub ex_titleEstrutura_Click()

End Sub

Private Sub Label58_Click()

End Sub

Private Sub Label56_Click()

End Sub

Private Sub Label55_Click()

End Sub

Private Sub Label65_Click()

End Sub

Private Sub MultiPage3_Change()

End Sub



Private Sub OptionButton2_Click()

End Sub

Private Sub MultiPage4_Change()

End Sub

Private Sub OptionButton4_Click()

End Sub

'INSUMO
Private Sub txtComboBox_nv4_4_Change()

End Sub

Private Sub ComboBox11_Change()

End Sub

Private Sub CommandButton12_Click()
GetAdicionais

End Sub

Private Sub CommandButton13_Click()
GetPvs_Cmo
End Sub

Sub submitTerceiros()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim new_id3 As Integer
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Me.GetIdNv3
If txtBox_nv3_3 = "" Then

    MsgBox "O campo está em branco, insira um valor para prosseguir com o cadastro"
    Exit Sub

ElseIf id_nv3 = 0 Then
't_Nivel3 ADIÇÃO
    ConectarBanco conexao
    sql = "t_Nivel3"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!descricaoNv3 = Me.txtBox_nv3_3.Value
    rs!grupo = "SERVICOS DE TERCEIROS"
    
    rs.Update
    rs.Close
    conexao.Close
    
'Pegar novo ID em t_Nivel3
    ConectarBanco conexao
    sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS DE TERCEIROS'"
    
        
    rs.Open sql1, conexao
    
    new_id3 = rs.Fields("idNv3").Value
    
    rs.Close
    conexao.Close

    
't_Servicos_Terceiros
    ConectarBanco conexao
    sql = "t_Servicos_Terceiros"
    Dim id_master As String

    id_master = 9 & "-" & 0 & "-" & new_id3

    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    

    rs!descricaoNv3 = txtBox_nv3_3.Value
    rs!idNv1 = 9
    rs!idNv2 = 0
    rs!idNv3 = new_id3
    rs!id = id_master

    
    rs.Update
    rs.Close
    conexao.Close
    
'Cria Log
    log_adicao_terceiros
    MsgBox "Cadastro realizado!"
    
    
    
    ClearFields
Else

        MsgBox "Esta estrutura de serviço já existe.", vbExclamation
'    CommandButton15.Visible = True
'    Enviar.Visible = False
    
    CommandButton15.Visible = False
    Enviar.Visible = True
    
End If
    


End Sub


Private Sub Enviar_Click()
        
Dim selectedRow As Integer
Dim i As Integer
Dim serv_atual As String


If sTerceiros = True Then
submitTerceiros

Exit Sub
End If

If sDiversos = True Then
submitDiversos
Exit Sub
End If



i = 0

'Checar N1
    If (Me.txtBox_nv1_1.Value <> "" And Me.txtComboBox_nv1_1.Value = "") Or (Me.txtBox_nv1_1.Value = "" And Me.txtComboBox_nv1_1.Value <> "") Then
    i = i + 1
    'MsgBox "nivel 1 = " & i
    End If
'Checar N2
    If (Me.txtBox_nv2_2.Value <> "" And Me.txtComboBox_nv2_2.Value = "") Or (Me.txtBox_nv2_2.Value = "" And Me.txtComboBox_nv2_2.Value <> "") Then
    i = i + 1
    'MsgBox "nivel 2 = " & i
    End If
'Checar N3
    If (Me.txtBox_nv3_3.Value <> "" And Me.txtComboBox_nv3_3.Value = "") Or (Me.txtBox_nv3_3.Value = "" And Me.txtComboBox_nv3_3.Value <> "") Then
    i = i + 1
    'MsgBox "nivel 3 = " & i
    End If
    
'Checar N4
    If (Me.txtBox_nv4_4.Value <> "" And Me.txtComboBox_nv4_4.Value = "") Or (Me.txtBox_nv4_4.Value = "" And Me.txtComboBox_nv4_4.Value <> "") Then
    i = i + 1
    'MsgBox "nivel 4 = " & i
    End If
    

    
    
'Formata para "R$ #,##0.00" e susbtitui dentro do input antes de verificar se todos os campos após o "avançar" são diferentes de "R$ 0,00" ou vazio
        If IsNumeric(Me.txtBox_custoInsumo.Value) Then
        Me.txtBox_custoInsumo.Value = Format(Me.txtBox_custoInsumo.Value, "R$ #,##0.00")
    End If


    If IsNumeric(Me.txtBox_pvs.Value) Then
        Me.txtBox_pvs.Value = Format(Me.txtBox_pvs.Value, "R$ #,##0.00")
    End If


    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.txtBox_cmo.Value) Then
        Me.txtBox_cmo.Value = Format(Me.txtBox_cmo.Value, "R$ #,##0.00")
    End If



    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.txtBox_rendimento.Value) Then
        Me.txtBox_rendimento.Value = Format(Me.txtBox_rendimento.Value, "#,##0.00")
    End If
    
        ' Formata o conteúdo do TextBox como moeda
If IsNumeric(Me.txtBox_un.Value) Then
    Me.txtBox_un.Value = ""
    MsgBox "Valor presente no campo UNIDADE não é permitido!", vbExclamation
    Exit Sub
ElseIf Not IsNumeric(Me.txtBox_un.Value) Then
    Me.txtBox_un.Value = UCase(Me.txtBox_un.Value)
End If
    

    
    txtBox_un.Value = Trim(Replace(txtBox_un.Value, "  ", " "))
    txtBox_rendimento.Value = Trim(Replace(txtBox_rendimento.Value, "  ", " "))
    txtBox_custoInsumo.Value = Trim(Replace(txtBox_custoInsumo.Value, "  ", " "))
    txtBox_pvs.Value = Trim(Replace(txtBox_pvs.Value, "  ", " "))
    txtBox_cmo.Value = Trim(Replace(txtBox_cmo.Value, "  ", " "))

    
        If Len(Trim(txtBox_un.Value)) <> 0 And Len(Trim(txtBox_rendimento.Value)) <> 0 And Len(Trim(txtBox_custoInsumo.Value)) <> 0 And Len(Trim(txtBox_pvs.Value)) <> 0 And Len(Trim(txtBox_cmo.Value)) <> 0 Then
             If txtBox_un.Value <> "R$ 0,00" And txtBox_rendimento.Value <> "R$ 0,00" And txtBox_custoInsumo.Value <> "R$ 0,00" And txtBox_pvs.Value <> "R$ 0,00" And txtBox_cmo.Value <> "R$ 0,00" Then
                If txtBox_un.Value <> "R$ 0,00" And txtBox_un.Value <> "R$" And txtBox_un.Value <> "R" And txtBox_un.Value <> "$" And txtBox_un.Value <> "," Then
                    If txtBox_rendimento.Value <> "0,00" And txtBox_rendimento.Value <> "R$" And txtBox_rendimento.Value <> "R" And txtBox_rendimento.Value <> "$" And txtBox_rendimento.Value <> "," Then
                        If txtBox_custoInsumo.Value <> "R$ 0,00" And txtBox_custoInsumo.Value <> "R$" And txtBox_custoInsumo.Value <> "R" And txtBox_custoInsumo.Value <> "$" And txtBox_custoInsumo.Value <> "," Then
                            If txtBox_pvs.Value <> "R$ 0,00" And txtBox_pvs.Value <> "R$" And txtBox_pvs.Value <> "R" And txtBox_pvs.Value <> "$" And txtBox_pvs.Value <> "," Then
                                If txtBox_cmo.Value <> "R$ 0,00" And txtBox_cmo.Value <> "R$" And txtBox_cmo.Value <> "R" And txtBox_cmo.Value <> "$" And txtBox_cmo.Value <> "," Then
                                i = i + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
                                

  
    
    
    
    
    
    

'O rendimtneto já roda durante toda mudanças em niveis 1,2,3 e insumo
'GetRendimento1
    If i = 5 And rendimentoValue = True Then
        MsgBox "Cadastro existente!"
        Exit Sub
    ElseIf i <> 5 Then
        'MsgBox "Existem valores a serem preenchidos!" & vbCrLf & " " & vbCrLf & "Apenas valores maiores que zero serão considerados." & vbCrLf & "Campos em branco não serão considerados.", vbExclamation
        MsgBox "Existem campos não preenchidos ou com valor informado igual a zero!", vbExclamation
    Else
    
'O GetAdicionais já roda durante toda mudanças em niveis 1,2,3 e insumo
'GetAdicionais
    
 '==========add service========
    
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset


    If Not IsNull(Me.txtBox_nv1_1.Value) And Len(Me.txtBox_nv1_1.Value) > 0 And id_nv1 = 0 Then
        ConectarBanco conexao
        sql = "t_Nivel1"
        
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs.AddNew
    
        rs!descricaoNv1 = Me.txtBox_nv1_1.Value
        rs!grupo = "SERVICOS PRINCIPAIS"
        
        rs.Update
        rs.Close
        conexao.Close
'== log data ==
Me.log_ad_txtBox_nv1_1 = Me.txtBox_nv1_1.Value
Me.log_ad_txtComboBox_nv1_1 = Me.txtComboBox_nv1_1.Value
'==============
    End If
    
'Pegar new_id1 de txtbox, senão pegar do comboBox
    If Me.txtBox_nv1_1.Value <> "" Then
        ConectarBanco conexao
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtBox_nv1_1.Value & "';"
        rs.Open sql1, conexao
        
        new_id1 = rs.Fields("idNv1").Value
        
        rs.Close
        conexao.Close

    Else


        'Obter o índice da linha selecionada no listbox
        selectedRow = Me.txtComboBox_nv1_1.ListIndex
        
        ConectarBanco conexao
        'sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtBox_nv1_1.Value & "';"
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtComboBox_nv1_1.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
        
        new_id1 = rs.Fields("idNv1").Value
        
        rs.Close
        conexao.Close
    End If

    

    

    If Not IsNull(Me.txtBox_nv2_2.Value) And Len(Me.txtBox_nv2_2.Value) > 0 And id_nv2 = 0 Then
        
    ConectarBanco conexao
    sql = "t_Nivel2"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew

    rs!descricaoNv2 = Me.txtBox_nv2_2.Value
    rs!grupo = "SERVICOS PRINCIPAIS"
    
    rs.Update
    rs.Close
    conexao.Close
'== log data ==
Me.log_ad_txtBox_nv2_2 = Me.txtBox_nv2_2.Value
Me.log_ad_txtComboBox_nv2_2 = Me.txtComboBox_nv2_2.Value
'==============
    End If
    
'Pegar new_id2
If Me.txtBox_nv2_2.Value <> "" Then
ConectarBanco conexao
sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & Me.txtBox_nv2_2.Value & "';"
rs.Open sql1, conexao

new_id2 = rs.Fields("idNv2").Value

rs.Close
conexao.Close
Else

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv2_2.ListIndex

ConectarBanco conexao

sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & Me.txtComboBox_nv2_2.Column(0, selectedRow) & "'"
rs.Open sql1, conexao

new_id2 = rs.Fields("idNv2").Value

rs.Close
conexao.Close

End If
    

    

    If Not IsNull(Me.txtBox_nv3_3.Value) And Len(Me.txtBox_nv3_3.Value) > 0 And id_nv3 = 0 Then
    ConectarBanco conexao
    sql = "t_Nivel3"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!descricaoNv3 = Me.txtBox_nv3_3.Value
    rs!grupo = "SERVICOS PRINCIPAIS"
    
    rs.Update
    rs.Close
    conexao.Close
'== log data ==
Me.log_ad_txtBox_nv3_3 = Me.txtBox_nv3_3.Value
Me.log_ad_txtComboBox_nv3_3 = Me.txtComboBox_nv3_3.Value
'==============
    End If
    
'Pegar new_id3

If Me.txtBox_nv3_3.Value <> "" Then


ConectarBanco conexao
'EXCLUIR  condições abaixo, o trecho atual sempre será apenas para o serviço principal
If sPrincipais = True Then
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS PRINCIPAIS'"
ElseIf sDiversos = True Then
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS DIVERSOS'"
ElseIf sTerceiros = True Then
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS DE TERCEIROS'"
End If
'sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "' AND grupo = '" & serv_atual & "';"

    rs.Open sql1, conexao
    
    new_id3 = rs.Fields("idNv3").Value
    
    rs.Close
    conexao.Close
Else

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv3_3.ListIndex

ConectarBanco conexao
If sPrincipais = True Then
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS PRINCIPAIS'"
ElseIf sDiversos = True Then
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DIVERSOS'"
ElseIf sTerceiros = True Then
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DE TERCEIROS'"
End If

rs.Open sql1, conexao

new_id3 = rs.Fields("idNv3").Value

rs.Close
conexao.Close

End If
    

    If Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 And id_nv4 = 0 Then
    ConectarBanco conexao
    sql = "t_Insumos"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!insumo = Me.txtBox_nv4_4.Value
    rs!CUSTO = Me.txtBox_custoInsumo.Value
    rs!unidade = Me.txtBox_un.Value
    
    rs.Update
    rs.Close
    conexao.Close
'== log data ==
Me.log_ad_txtBox_nv4_4 = Me.txtBox_nv4_4.Value
Me.log_ad_txtBox_nv4_4 = Me.txtComboBox_nv4_4.Value
Me.log_ad_txtBox_custoInsumo = Me.txtBox_custoInsumo.Value
Me.log_ad_txtBox_un = Me.txtBox_un.Value
'==============

newInsumo = True
   End If
   
'Pegar new_id4
If Me.txtBox_nv4_4.Value <> "" Then
ConectarBanco conexao
sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & Me.txtBox_nv4_4.Value & "';"
rs.Open sql1, conexao

new_id4 = rs.Fields("idInsumo").Value

rs.Close
conexao.Close

Else

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex

ConectarBanco conexao
'sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtBox_nv1_1.Value & "';"
sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
rs.Open sql1, conexao

new_id4 = rs.Fields("idInsumo").Value

rs.Close
conexao.Close

End If

'===========================================================================


'ADD Rendimento (t_Servicos_Principais_Rendimento)
    ConectarBanco conexao
    sql = "t_Servicos_Principais_Rendimento"
    


    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!idNv1 = new_id1
    rs!idNv2 = new_id2
    rs!idNv3 = new_id3
    rs!idInsumo = new_id4
    rs!rendimento = UserForm2.txtBox_rendimento.Value


    
    rs.Update
    rs.Close
    conexao.Close
'== log data ==
log_ad_txtBox_rendimento = txtBox_rendimento.Value

    
       If c_cmo = True And c_pvs = True Then
       'MsgBox "Uma nova linha sera formada, pois não há PSV, nem Cmo para a combinação de serviços listada"
        ConectarBanco conexao
        sql = "t_Servicos_Principais"
        
        
        
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs.AddNew
        
        rs!idNv1 = new_id1
        rs!idNv2 = new_id2
        rs!idNv3 = new_id3
        rs!precoVendaSugerido = txtBox_pvs.Value
        rs!CustoMaoObra = txtBox_cmo.Value
        
        rs.Update
        rs.Close
        conexao.Close
    
    ElseIf PvsUpdate <> pvs And PvsUpdate <> 0 Then
    
    
        ConectarBanco conexao
    
        'sql = "SELECT * FROM t_Servicos_Principais_Rendimento WHERE idNv = id_nv1"
        sql = "SELECT precoVendaSugerido FROM t_Servicos_Principais WHERE idNv1 = " & id_nv1 & " AND idNv2 = " & id_nv2 & " AND idNv3 = " & id_nv3
        
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        
        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para usbstituição."
        rs!precoVendaSugerido = PvsUpdate
        rs.Update
        rs.Close
        conexao.Close

    End If

    
        ConectarBanco conexao
        'sql = "SELECT * FROM t_Servicos_Principais_Rendimento WHERE idNv = id_nv1"
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic

        If CheckBoxGeneric.Value = True And newInsumo = True Then
        rs!tipo = "GENERICO"
        ElseIf CheckBoxGeneric.Value = False And newInsumo = False Then
        'rs!tipo = "SERVICOS PRINCIPAIS"
        GoTo jumpit
        ElseIf CheckBoxGeneric.Value = False And newInsumo = True Then
        rs!tipo = "SERVICOS PRINCIPAIS"
        End If
jumpit:
        rs.Update
        rs.Close
        conexao.Close
        
        
    
    
    
    
    
 '============================
'============= Log Data ===========
log_ad_txtBox_cmo = txtBox_cmo.Value
log_ad_txtBox_pvs = txtBox_pvs.Value
'==================================
  'Recarregará o forms para retirar os dados do último envio
  
'Criar Log
    kLogs.log_adicao_principal
    
    MsgBox "Cadastro realizado!", vbInformation

Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.EditarNiveis.Visible = False

CommandButton15.Visible = True
newInsumo = False
    ClearFields




    End If



    
End Sub

Sub submitDiversos()


i = 0


    If (Me.txtBox_nv3_3.Value <> "" And Me.txtComboBox_nv3_3.Value = "") Or (Me.txtBox_nv3_3.Value = "" And Me.txtComboBox_nv3_3.Value <> "") Then
    i = i + 1
    'MsgBox "nivel 3 = " & i
    End If
    
'Checar N4
    If (Me.txtBox_nv4_4.Value <> "" And Me.txtComboBox_nv4_4.Value = "") Or (Me.txtBox_nv4_4.Value = "" And Me.txtComboBox_nv4_4.Value <> "") Then
    i = i + 1
    'MsgBox "nivel 4 = " & i
    End If
    
    
    If IsNumeric(Me.txtBox_un.Value) Then
    Me.txtBox_un.Value = ""
    MsgBox "Valor presente no campo UNIDADE não é permitido!", vbExclamation
    Exit Sub
    ElseIf Not IsNumeric(Me.txtBox_un.Value) Then
        Me.txtBox_un.Value = UCase(Me.txtBox_un.Value)
    End If


    txtBox_un.Value = Trim(Replace(txtBox_un.Value, "  ", " "))
    txtBox_rendimento.Value = Trim(Replace(txtBox_rendimento.Value, "  ", " "))
    txtBox_custoInsumo.Value = Trim(Replace(txtBox_custoInsumo.Value, "  ", " "))
    txtBox_pvs.Value = Trim(Replace(txtBox_pvs.Value, "  ", " "))
    txtBox_cmo.Value = Trim(Replace(txtBox_cmo.Value, "  ", " "))

    
        If Len(Trim(txtBox_un.Value)) <> 0 And Len(Trim(txtBox_rendimento.Value)) <> 0 And Len(Trim(txtBox_custoInsumo.Value)) <> 0 And Len(Trim(txtBox_cmo.Value)) <> 0 Then
             If txtBox_un.Value <> "R$ 0,00" And txtBox_rendimento.Value <> "R$ 0,00" And txtBox_custoInsumo.Value <> "R$ 0,00" And txtBox_cmo.Value <> "R$ 0,00" Then
                If txtBox_un.Value <> "R$ 0,00" And txtBox_un.Value <> "R$" And txtBox_un.Value <> "R" And txtBox_un.Value <> "$" And txtBox_un.Value <> "," Then
                    If txtBox_rendimento.Value <> "0,00" And txtBox_rendimento.Value <> "R$" And txtBox_rendimento.Value <> "R" And txtBox_rendimento.Value <> "$" And txtBox_rendimento.Value <> "," Then
                        If txtBox_custoInsumo.Value <> "R$ 0,00" And txtBox_custoInsumo.Value <> "R$" And txtBox_custoInsumo.Value <> "R" And txtBox_custoInsumo.Value <> "$" And txtBox_custoInsumo.Value <> "," Then
                            If txtBox_cmo.Value <> "R$ 0,00" And txtBox_cmo.Value <> "R$" And txtBox_cmo.Value <> "R" And txtBox_cmo.Value <> "$" And txtBox_cmo.Value <> "," Then
                                i = i + 1
                            End If

                        End If
                    End If
                End If
            End If
        End If










    If i = 3 And rendimentoValue = True Then
        MsgBox "Cadastro existente!"
        Exit Sub
    ElseIf i <> 3 Then
        MsgBox "Existem campos não preenchidos ou com valor informado igual a zero!", vbExclamation
    Else
    

    
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset


 

    

    If Not IsNull(Me.txtBox_nv3_3.Value) And Len(Me.txtBox_nv3_3.Value) > 0 And id_nv3 = 0 Then
    ConectarBanco conexao
    sql = "t_Nivel3"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!descricaoNv3 = Me.txtBox_nv3_3.Value
    rs!grupo = "SERVICOS DIVERSOS"
    
    rs.Update
    rs.Close
    conexao.Close
    End If
    
'Pegar new_id3
If Me.txtBox_nv3_3.Value <> "" Then
ConectarBanco conexao
'sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "';"
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS DIVERSOS'"

    rs.Open sql1, conexao
    
    new_id3 = rs.Fields("idNv3").Value
    
    rs.Close
    conexao.Close
Else

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv3_3.ListIndex

ConectarBanco conexao

'sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "'"
sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DIVERSOS'"
rs.Open sql1, conexao

new_id3 = rs.Fields("idNv3").Value

rs.Close
conexao.Close

End If
    

    If Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 And id_nv4 = 0 Then
    ConectarBanco conexao
    sql = "t_Insumos"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!insumo = Me.txtBox_nv4_4.Value
    rs!CUSTO = Me.txtBox_custoInsumo.Value
    rs!unidade = Me.txtBox_un.Value
    
    rs.Update
    rs.Close
    conexao.Close
    
    newInsumo = True
   End If
   
'Pegar new_id4
If Me.txtBox_nv4_4.Value <> "" Then
ConectarBanco conexao
sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & Me.txtBox_nv4_4.Value & "';"
rs.Open sql1, conexao

new_id4 = rs.Fields("idInsumo").Value

rs.Close
conexao.Close

Else

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex

ConectarBanco conexao
'sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtBox_nv1_1.Value & "';"
sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
rs.Open sql1, conexao

new_id4 = rs.Fields("idInsumo").Value

rs.Close
conexao.Close

End If


    ConectarBanco conexao
    sql = "t_Servicos_Diversos_Rendimento"
    
    rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
    rs.AddNew
    
    rs!idNv1 = 7
    rs!idNv2 = 0
    rs!idNv3 = new_id3
    rs!idInsumo = new_id4
    rs!rendimento = txtBox_rendimento.Value

    rs.Update
    rs.Close
    conexao.Close
    
    

    
       If c_cmo = True And c_pvs = True Then
       'MsgBox "Uma nova linha sera formada, pois não há PSV, nem Cmo para a combinação de serviços listada"
        ConectarBanco conexao
        sql = "t_Servicos_Diversos"
        
        
        
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs.AddNew
        
        rs!idNv1 = 7
        rs!idNv2 = 0
        rs!idNv3 = new_id3
        rs!precoVendaSugerido = 0
        rs!CustoMaoObra = txtBox_cmo.Value
        
        rs.Update
        rs.Close
        conexao.Close
    
    End If



        ConectarBanco conexao
        'sql = "SELECT * FROM t_Servicos_Principais_Rendimento WHERE idNv = id_nv1"
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic

        If CheckBoxGeneric.Value = True And newInsumo = True Then
        rs!tipo = "GENERICO"
        ElseIf CheckBoxGeneric.Value = False And newInsumo = False Then
'        rs!tipo = "SERVICOS DIVERSOS"
        GoTo jumpit
        ElseIf CheckBoxGeneric.Value = False And newInsumo = True Then
        rs!tipo = "SERVICOS DIVERSOS"
        End If
jumpit:
        rs.Update
        rs.Close
        conexao.Close









 '============================
    kLogs.log_adicao_diversos
  
    MsgBox "Cadastro realizado!", vbInformation
    Me.Label27.Visible = False
    Me.txtBox_un.Visible = False
    Me.Label38.Visible = False
    Me.txtBox_rendimento.Visible = False
    Me.Label36.Visible = False
    Me.txtBox_cmo.Visible = False
    Me.Label44.Visible = False
    Me.txtBox_custoInsumo.Visible = False
    Me.Label35.Visible = False
    Me.txtBox_pvs.Visible = False
    Me.Enviar.Visible = False
    Me.EditarNiveis.Visible = False
    '=======================
    CommandButton15.Visible = True
    ClearFields




    End If




End Sub

Sub ClearFields()


    If sPrincipais = True Then


'NIVEL1
'Me.s_diversos.BackColor = RGB(35, 55, 100)
Me.s_principais.BackColor = RGB(142, 162, 219)
Me.s_terceiros.ForeColor = &H80000008
Me.s_principais.Font.Bold = True
'Me.MultiPage2.Value = 2
'Me.MultiPage2.BackColor = RGB(142, 169, 219)
'Me.MultiPage1.BackColor = RGB(142, 169, 219)


Me.s_diversos.BackColor = &H8000000F
Me.s_diversos.ForeColor = &H80000008
Me.s_diversos.Font.Bold = False

Me.s_terceiros.BackColor = &H8000000F
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = False










Me.txtBox_nv1_1.Value = ""
Me.txtComboBox_nv1_1.Value = ""
Me.txtBox_nv2_2.Value = ""
Me.txtComboBox_nv2_2.Value = ""
Me.txtBox_nv3_3.Value = ""
Me.txtComboBox_nv3_3.Value = ""
Me.txtBox_nv4_4.Value = ""
Me.txtComboBox_nv4_4.Value = ""
Me.txtBox_un.Value = ""
Me.txtBox_rendimento.Value = ""
Me.txtBox_custoInsumo.Value = ""
Me.txtBox_pvs.Value = ""
Me.txtBox_cmo.Value = ""
Me.CheckBoxGeneric.Value = False

sPrincipais = True
sTerceiros = False
sDiversos = False
Me.CarregarNivelUm
Adicionar.CarregarNivelDois
Me.CarregarNivelTres

Me.txtBox_nv1_1.Enabled = True
Me.txtBox_nv1_1.BackColor = RGB(255, 255, 255)
'Me.txtBox_nv1_1.Font.Bold = True

Me.txtBoxID_nv1.Enabled = True
Me.txtBoxID_nv1.BackColor = RGB(255, 255, 255)

Me.txtComboBox_nv1_1.Enabled = True
Me.txtComboBox_nv1_1.BackColor = RGB(255, 255, 255)

cbBox_nv1.optionButton_nv1_1.Value = False
cbBox_nv1.optionButton_nv1_2.Value = False

Me.cbBox_nv1.Enabled = True

'NIVEL 2
Me.txtBox_nv2_2.Enabled = True
Me.txtBox_nv2_2.BackColor = RGB(255, 255, 255)
'Me.txtBox_nv2_2.Font.Bold = True

Me.cbBox_nv2_2.Enabled = True
'Me.cbBox_nv2_2.BackColor = RGB(255, 255, 255)

Me.txtComboBox_nv2_2.Enabled = True
Me.txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)

cbBox_nv2_2.optionButton_nv2_4.Value = False
cbBox_nv2_2.optionButton_nv2_3.Value = False


'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = True
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016

'INSUMO
    Me.txtBox_nv4_4.Enabled = True
    Me.txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = True
    'Me.cbBox_nv4_4.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    
    Me.txtBox_rendimento.Enabled = False
    Me.txtBox_rendimento.BackColor = &H80000016
    
    Me.txtBox_custoInsumo.Enabled = False
    Me.txtBox_custoInsumo.BackColor = &H80000016
    
    Me.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    
    Me.txtBox_cmo.Enabled = False
    Me.txtBox_cmo.BackColor = &H80000016
    
    
    
Me.optionButton_nv1_2.Value = True
Me.optionButton_nv2_3.Value = True
Me.optionButton_nv3_6.Value = True
Me.optionButton_nv4_8.Value = True



'========== TESTE =======
Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.CheckBoxGeneric.Visible = False


'=======================



    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True



    End If

    If sTerceiros = True Then


Me.s_terceiros.BackColor = RGB(255, 242, 204)
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = True
Me.MultiPage2.Value = 2
'Me.MultiPage1.BackColor = RGB(255, 249, 231)

Me.s_principais.BackColor = &H8000000F
Me.s_principais.ForeColor = &H80000008
Me.s_principais.Font.Bold = False

Me.s_diversos.BackColor = &H8000000F
Me.s_diversos.ForeColor = &H80000008
Me.s_diversos.Font.Bold = False



Me.txtBox_nv1_1.Value = ""
Me.txtComboBox_nv1_1.Value = ""
Me.txtBox_nv2_2.Value = ""
Me.txtComboBox_nv2_2.Value = ""
Me.txtBox_nv3_3.Value = ""
Me.txtComboBox_nv3_3.Value = ""
Me.txtBox_nv4_4.Value = ""
Me.txtComboBox_nv4_4.Value = ""
Me.txtBox_un.Value = ""
Me.txtBox_rendimento.Value = ""
Me.txtBox_custoInsumo.Value = ""
Me.txtBox_pvs.Value = ""
Me.txtBox_cmo.Value = ""
Me.CheckBoxGeneric.Value = False
Me.CheckBoxGeneric.Visible = False

sTerceiros = True
sPrincipais = False
sDiversos = False

id_nv1 = 9
id_nv2 = 0
'Me.CarregarNivelUm
'NIVEL 1
    Me.txtBox_nv1_1.Enabled = False
    Me.txtBox_nv1_1.BackColor = &H80000016
    'Me.txtBox_nv1_1.Font.Bold = True
    
    Me.txtBoxID_nv1.Enabled = False
    Me.txtBoxID_nv1.BackColor = &H80000016
    
    Me.txtComboBox_nv1_1.Enabled = False
    Me.txtComboBox_nv1_1.BackColor = &H80000016
    
    Me.cbBox_nv1.Enabled = False

'NIVEL 2
    Me.txtBox_nv2_2.Enabled = False
    Me.txtBox_nv2_2.BackColor = &H80000016
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv2_2.Enabled = False
    'Me.cbBox_nv2_2.BackColor = &H80000016
    
    Me.txtComboBox_nv2_2.Enabled = False
    Me.txtComboBox_nv2_2.BackColor = &H80000016

'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = False
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016

'INSUMO
    Me.txtBox_nv4_4.Enabled = False
    Me.txtBox_nv4_4.BackColor = &H80000016
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = False
    'Me.cbBox_nv4_4.BackColor = &H80000016
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    
    Me.txtBox_rendimento.Enabled = False
    Me.txtBox_rendimento.BackColor = &H80000016
    
    Me.txtBox_custoInsumo.Enabled = False
    Me.txtBox_custoInsumo.BackColor = &H80000016
    
    Me.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    
    Me.txtBox_cmo.Enabled = False
    Me.txtBox_cmo.BackColor = &H80000016



'    CommandButton15.Visible = True
'    Enviar.Visible = False
    
    CommandButton15.Visible = False
    Enviar.Visible = True

    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True
    
    End If

    If sDiversos = True Then

Me.s_diversos.BackColor = RGB(255, 230, 153)
Me.s_terceiros.ForeColor = &H80000008
Me.s_diversos.Font.Bold = True
Me.MultiPage2.Value = 2
'Me.MultiPage1.BackColor = RGB(142, 162, 219)

Me.s_principais.BackColor = &H8000000F
Me.s_principais.ForeColor = &H80000008
Me.s_principais.Font.Bold = False

Me.s_terceiros.BackColor = &H8000000F
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = False

sDiversos = True
sTerceiros = False
sPrincipais = False
Me.CarregarNivelTres

id_nv1 = 7
id_nv2 = 0

Me.txtBox_nv1_1.Enabled = False
Me.txtBox_nv1_1.BackColor = &H80000016 '&H80000016
Me.txtBox_nv1_1.Value = ""
'Me.txtBox_nv1_1.Font.Bold = True

Me.txtBoxID_nv1.Enabled = False
Me.txtBoxID_nv1.BackColor = &H80000016

Me.txtComboBox_nv1_1.Enabled = False
Me.txtComboBox_nv1_1.BackColor = &H80000016
Me.txtComboBox_nv1_1.Value = ""
'NIVEL 2
Me.txtBox_nv2_2.Enabled = False
Me.txtBox_nv2_2.BackColor = &H80000016
Me.txtBox_nv2_2.Value = ""
'Me.txtBox_nv2_2.Font.Bold = True

Me.cbBox_nv2_2.Enabled = False
'Me.cbBox_nv2_2.BackColor = &H80000016

Me.txtComboBox_nv2_2.Enabled = False
Me.txtComboBox_nv2_2.BackColor = &H80000016
Me.txtComboBox_nv2_2.Value = ""

Me.cbBox_nv1.Enabled = False

'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    Me.txtBox_nv3_3.Value = ""
    Me.optionButton_nv3_5.Value = False
    Me.optionButton_nv3_6.Value = True
    
    'Me.txtBox_nv3_3.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = True
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016
    Me.txtComboBox_nv3_3.Value = ""
    

'INSUMO
    Me.txtBox_nv4_4.Enabled = True
    Me.txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    Me.txtBox_nv4_4.Value = ""
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = True
    optionButton_nv4_7.Value = False
    optionButton_nv4_8.Value = True
    'Me.cbBox_nv4_4.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    Me.txtComboBox_nv4_4.Value = ""
    
    Me.txtBox_un.Enabled = True
    Me.txtBox_un.BackColor = RGB(255, 255, 255)
    Me.txtBox_un.Value = ""
    
    Me.txtBox_rendimento.Enabled = True
    Me.txtBox_rendimento.BackColor = RGB(255, 255, 255)
    Me.txtBox_rendimento.Value = ""
    
    Me.txtBox_custoInsumo.Enabled = True
    Me.txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
    Me.txtBox_custoInsumo.Value = ""
    
    Me.txtBox_pvs.Enabled = True
    Me.txtBox_pvs.BackColor = RGB(255, 255, 255)
    Me.txtBox_pvs.Value = ""
    
    Me.txtBox_cmo.Enabled = True
    Me.txtBox_cmo.BackColor = RGB(255, 255, 255)
    Me.txtBox_cmo.Value = ""
    
    Me.CheckBoxGeneric.Visible = False
    Me.CheckBoxGeneric.Value = False



    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True


    End If





End Sub

Private Sub CloseForm()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label45_Click()

End Sub

'|!|!|!|!| NIVEL 1 |!|!|!|!|
Private Sub optionButton_nv1_1_Click()
    Me.txtBox_nv1_1.Value = ""
    Me.txtComboBox_nv1_1.Enabled = True
    Me.txtBox_nv1_1.Enabled = False
    Me.txtBox_nv1_1.BackColor = &H80000016
    Me.txtComboBox_nv1_1.BackColor = RGB(255, 255, 255)
End Sub
'|!|!|!|!| NIVEL 1 |!|!|!|!|
Private Sub optionButton_nv1_2_Click()
    Me.txtComboBox_nv1_1.Value = ""
    Me.txtComboBox_nv1_1.BackColor = &H80000016
    Me.txtComboBox_nv1_1.Enabled = False
'    Me.txtComboBox_nv1_1.Value = "Selecione um serviço existente"
'    UserForm2.txtBox_nv1_1.Font.Italic = True
    Me.txtBox_nv1_1.Enabled = True
    Me.txtBox_nv1_1.BackColor = RGB(255, 255, 255)

End Sub

Private Sub TextBox22_Change()

End Sub

'Private Sub optionButton_nv1_2_Click()
'
'End Sub

'|!|!|!|!| NIVEL 1 |!|!|!|!|
Private Sub txtComboBox_nv1_1_Click()

    Me.optionButton_nv1_1.Value = True
    

    Me.txtBox_nv1_1.Value = ""
    Me.txtComboBox_nv1_1.Enabled = True
    Me.txtBox_nv1_1.Enabled = False
End Sub
'|!|!|!|!| NIVEL 1 |!|!|!|!|
Private Sub txtBox_nv1_1_Click()
    Me.optionButton_nv1_2.Value = True
    
    
    Me.txtComboBox_nv1_1.Value = ""
    Me.txtComboBox_nv1_1.Enabled = False
    Me.txtBox_nv1_1.Enabled = True
End Sub


'FUNCTION NÃO É UTILIZADA, APENAS FOI RETIRADA A LOGICA DO "SE NULL..."
Private Sub VerificarValor()
    Dim valor As String
    If Not IsNull(Me.txtComboBox_nv1_1.Value) And Len(Me.txtComboBox_nv1_1.Value) > 0 Then
        valor = Me.txtComboBox_nv1_1.Value
    ElseIf Not IsNull(Me.txtBox_nv1_1.Value) And Len(Me.txtBox_nv1_1.Value) > 0 Then
        valor = Me.txtBox_nv1_1.Value
    Else
        MsgBox "Selecione uma ComboBox com valor para verificar."
        Exit Sub
    End If
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM SuaTabela WHERE SuaColuna = '" & valor & "'")
    
    If Not (rs.EOF And rs.BOF) Then
        MsgBox "O valor " & valor & " existe na tabela."
    Else
        MsgBox "O valor " & valor & " não existe na tabela."
    End If
    
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


Private Sub txtComboBox_nv1_1_Change()



End Sub
'|!|!|!|!| NIVEL 1 |!|!|!|!|
'[ O CÓDIGO ABAIXO IRÁ TRAZER COM BASE NO CAMPO DE "lista de serviço pré criada" ou "novo serviço (campo livre)" O VALOR REFERENTE AQUELA LINHA
'PARA O CAMPO ID NA TABELA ACCESS, COLUNA "idNv1"

Sub GetIdNv1()


Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

id_nv1 = 0
id_nv2 = 0
id_nv3 = 0

ConectarBanco conexao

sql = "t_Nivel1"
sql1 = "t_Nivel1"


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv1_1.ListIndex
Dim selectedRow1 As String
selectedRow1 = Me.txtBox_nv1_1.Value


   Dim valor As String
   'AREAS TERREAS
    If Not IsNull(Me.txtComboBox_nv1_1.Value) And Len(Me.txtComboBox_nv1_1.Value) > 0 Then
        valor = Me.txtComboBox_nv1_1.Value
        selectedRow = Me.txtComboBox_nv1_1.ListIndex
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtComboBox_nv1_1.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(Me.txtBox_nv1_1.Value) And Len(Me.txtBox_nv1_1.Value) > 0 Then
   
        On Error GoTo here
        selectedRow1 = Me.txtBox_nv1_1.Value
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & selectedRow1 & "';"
        'selectedRow = Me.txtBox_nv1_1.Value
        valor = Me.txtBox_nv1_1.Value
        'selectedRow = Me.txtBox_nv1_1.Value
        'sql1 = "SELECT * FROM t_Nivel1 WHERE idNv1 = '" & selectedRow & "';"
     
   End If


If selectedRow = -1 Then
idNv1 = 0
Me.txtBoxID_nv1.Value = idNv1


If (selectedRow = -1 And Me.txtBox_nv1_1.Value <> "Digite o serviço aqui") And (selectedRow = -1 And Me.txtBox_nv1_1.Value <> "") Then

rs.Open sql1, conexao

idNv1 = rs.Fields("idNv1").Value
Me.txtBoxID_nv1.Value = idNv1

id_nv1 = idNv1
rs.Close
conexao.Close

End If

Else


rs.Open sql1, conexao
'grupo = rs.Fields("grupo").Value
idNv1 = rs.Fields("idNv1").Value
'Mostre o valor obtido na célula ativa
'Me.ComboBox10.Value = grupo
Me.txtBoxID_nv1.Value = idNv1
id_nv1 = idNv1
rs.Close
conexao.Close



End If

    GoTo jumpit
here: Exit Sub
jumpit:

End Sub



Sub GetIdNv2()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel2"
sql1 = "t_Nivel2"


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv2_2.ListIndex
Dim selectedRow1 As String
selectedRow1 = Me.txtBox_nv2_2.Value



   Dim valor As String
    If Not IsNull(Me.txtComboBox_nv2_2.Value) And Len(Me.txtComboBox_nv2_2.Value) > 0 Then
        valor = Me.txtComboBox_nv2_2.Value
        selectedRow = Me.txtComboBox_nv2_2.ListIndex
        sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & Me.txtComboBox_nv2_2.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(Me.txtBox_nv2_2.Value) And Len(Me.txtBox_nv2_2.Value) > 0 Then
'        valor = Me.txtBox_nv2_2.Value
'        selectedRow = Me.txtBox_nv2_2.Value
'        sql1 = "SELECT * FROM t_Nivel2 WHERE idNv2 = '" & selectedRow & "';"
        
        
            '---TESTE
            On Error GoTo here
            selectedRow1 = Me.txtBox_nv2_2.Value
            sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & selectedRow1 & "';"
            '---
     
   End If

'===============================================================






            If selectedRow = -1 Then
            idNv2 = 0
            Me.txtBoxID_nv2.Value = idNv2
            
            
            If (selectedRow = -1 And Me.txtBox_nv2_2.Value <> "Digite o serviço aqui") And (selectedRow = -1 And Me.txtBox_nv2_2.Value <> "") Then
            
            rs.Open sql1, conexao
            
            idNv2 = rs.Fields("idNv2").Value
            Me.txtBoxID_nv2.Value = idNv2
            id_nv2 = idNv2
            rs.Close
            conexao.Close
            
            End If
            
            Else
            

            rs.Open sql1, conexao
            'grupo = rs.Fields("grupo").Value
            idNv2 = rs.Fields("idNv2").Value
            'Mostre o valor obtido na célula ativa
            'Me.ComboBox10.Value = grupo
            Me.txtBoxID_nv2.Value = idNv2
            id_nv2 = idNv2
            rs.Close
            conexao.Close
            
            
            
            End If
            
                GoTo jumpit
here:             Exit Sub
jumpit:
            
            
            '----


End Sub




Sub GetIdNv3()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim valor As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel3"
sql1 = "t_Nivel3"


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv3_3.ListIndex
Dim selectedRow1 As String
selectedRow1 = Me.txtBox_nv3_3.Value


If sPrincipais = True Then

    If Not IsNull(Me.txtComboBox_nv3_3.Value) And Len(Me.txtComboBox_nv3_3.Value) > 0 Then
        valor = Me.txtComboBox_nv3_3.Value
        selectedRow = Me.txtComboBox_nv3_3.ListIndex
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS PRINCIPAIS'"
   ElseIf Not IsNull(Me.txtBox_nv3_3.Value) And Len(Me.txtBox_nv3_3.Value) > 0 Then
'        valor = Me.txtBox_nv3_3.Value
'        selectedRow = Me.txtBox_nv3_3.Value
'        sql1 = "SELECT * FROM t_Nivel3 WHERE idNv3 = '" & selectedRow & "';"
            '---TESTE
            On Error GoTo here
            selectedRow1 = Me.txtBox_nv3_3.Value
'            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "';"
            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "' AND grupo = 'SERVICOS PRINCIPAIS';"
            '---
     
     
   End If
ElseIf sDiversos = True Then

    If Not IsNull(Me.txtComboBox_nv3_3.Value) And Len(Me.txtComboBox_nv3_3.Value) > 0 Then
        valor = Me.txtComboBox_nv3_3.Value
        selectedRow = Me.txtComboBox_nv3_3.ListIndex
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DIVERSOS'"
   ElseIf Not IsNull(Me.txtBox_nv3_3.Value) And Len(Me.txtBox_nv3_3.Value) > 0 Then
'        valor = Me.txtBox_nv3_3.Value
'        selectedRow = Me.txtBox_nv3_3.Value
'        sql1 = "SELECT * FROM t_Nivel3 WHERE idNv3 = '" & selectedRow & "';"
            '---TESTE
            On Error GoTo here
            selectedRow1 = Me.txtBox_nv3_3.Value
            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "' AND grupo = 'SERVICOS DIVERSOS';"
            '---
     
     
   End If
ElseIf sTerceiros = True Then

    If Not IsNull(Me.txtComboBox_nv3_3.Value) And Len(Me.txtComboBox_nv3_3.Value) > 0 Then
        valor = Me.txtComboBox_nv3_3.Value
        selectedRow = Me.txtComboBox_nv3_3.ListIndex
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & Me.txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DE TERCEIROS'"
   ElseIf Not IsNull(Me.txtBox_nv3_3.Value) And Len(Me.txtBox_nv3_3.Value) > 0 Then
'        valor = Me.txtBox_nv3_3.Value
'        selectedRow = Me.txtBox_nv3_3.Value
'        sql1 = "SELECT * FROM t_Nivel3 WHERE idNv3 = '" & selectedRow & "';"
            '---TESTE
            On Error GoTo here
            selectedRow1 = Me.txtBox_nv3_3.Value
'            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "';"
            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "' AND grupo = 'SERVICOS DE TERCEIROS';"
            '---
     
     
   End If
End If







            '----TESTE
            If selectedRow = -1 Then
            idNv3 = 0
            Me.txtBoxID_nv3.Value = idNv3
            
            
            If (selectedRow = -1 And Me.txtBox_nv3_3.Value <> "Digite o serviço aqui") And (selectedRow = -1 And Me.txtBox_nv3_3.Value <> "") Then
            
            rs.Open sql1, conexao
            
            idNv3 = rs.Fields("idNv3").Value
            Me.txtBoxID_nv3.Value = idNv3
            id_nv3 = idNv3
            rs.Close
            conexao.Close
            
            End If
            
            Else
            
            '====
            
            'sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtComboBox_nv1_1.Column(0, selectedRow) & "'"
            'coluna2Valor = DLookup("grupo", "t_Nivel1", "descricaoNv1 = '" & Me.txtComboBox_nv1_1.Column(0, selectedRow) & "'")
            rs.Open sql1, conexao
            'grupo = rs.Fields("grupo").Value
            idNv3 = rs.Fields("idNv3").Value
            'Mostre o valor obtido na célula ativa
            'Me.ComboBox10.Value = grupo
            Me.txtBoxID_nv3.Value = idNv3
            id_nv3 = idNv3
            rs.Close
            conexao.Close
            
            
            
            End If
            
                GoTo jumpit
here:
 id_nv3 = 0
Exit Sub
jumpit:
            
            
            '----


End Sub



Sub GetIdInsumo()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Insumos"
sql1 = "t_Insumos"


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex
Dim selectedRow1 As String
selectedRow1 = Me.txtBox_nv4_4.Value



   Dim valor As String
    If Not IsNull(Me.txtComboBox_nv4_4.Value) And Len(Me.txtComboBox_nv4_4.Value) > 0 Then
        valor = Me.txtComboBox_nv4_4.Value
        selectedRow = Me.txtComboBox_nv4_4.ListIndex
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 Then
'        valor = Me.txtBox_nv4_4.Value
'        selectedRow = Me.txtBox_nv4_4.Value
'        sql1 = "SELECT * FROM t_Insumos WHERE idInsumo = '" & selectedRow & "';"
            '---TESTE
            On Error GoTo here
            selectedRow1 = Me.txtBox_nv4_4.Value
            sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & selectedRow1 & "';"
            '---
   End If






            '----TESTE
            If selectedRow = -1 Then
            idInsumo = 0
            Me.txtBoxID_nv4.Value = idInsumo
            
            
            If selectedRow = -1 And Me.txtBox_nv4_4.Value <> "" Then
            
            rs.Open sql1, conexao
            
            idInsumo = rs.Fields("idInsumo").Value
            Me.txtBoxID_nv4.Value = idInsumo
            id_nv4 = idInsumo
            Insumo_nv4 = selectedRow1
            If Insumo_nv4 <> "" Then
            insumoBolean = True
            End If
            
            rs.Close
            conexao.Close
            
            End If
            
            Else
            
            '====
            
            'sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & Me.txtComboBox_nv1_1.Column(0, selectedRow) & "'"
            'coluna2Valor = DLookup("grupo", "t_Nivel1", "descricaoNv1 = '" & Me.txtComboBox_nv1_1.Column(0, selectedRow) & "'")
            rs.Open sql1, conexao
            'grupo = rs.Fields("grupo").Value
            idInsumo = rs.Fields("idInsumo").Value
            'Mostre o valor obtido na célula ativa
            'Me.ComboBox10.Value = grupo
            Me.txtBoxID_nv4.Value = idInsumo
            idInsumo = idInsumo
            rs.Close
            conexao.Close
            
            
            
            End If
            
                GoTo jumpit
here:
            idInsumo = 0
            Me.txtBoxID_nv4.Value = idInsumo
            id_nv4 = idInsumo
Exit Sub
jumpit:
            
            
            '----

End Sub



Sub GetCustoInsumo()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
Dim i As Integer
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Insumos"
sql1 = "t_Insumos"


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String
i = i + 1
'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex
Dim selectedRow1 As String
selectedRow1 = Me.txtBox_nv4_4.Value



   Dim valor As String
    If Not IsNull(Me.txtComboBox_nv4_4.Value) And Len(Me.txtComboBox_nv4_4.Value) > 0 Then
        valor = Me.txtComboBox_nv4_4.Value
        selectedRow = Me.txtComboBox_nv4_4.ListIndex
        sql1 = "SELECT Custo FROM t_Insumos WHERE Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 Then
'        valor = Me.txtBox_nv4_4.Value
'        selectedRow = Me.txtBox_nv4_4.Value
'        sql1 = "SELECT * FROM t_Insumos WHERE Custo = '" & selectedRow & "';"
            '---TESTE
            'On Error GoTo here
            selectedRow1 = Me.txtBox_nv4_4.Value
            sql1 = "SELECT Custo FROM t_Insumos WHERE Insumo = '" & selectedRow1 & "';"
            '---
     
   End If







        rs.Open sql1, conexao
''grupo = rs.Fields("grupo").Value
        On Error GoTo here
        CUSTO = rs.Fields("Custo").Value

        
        If selectedRow = -1 And selectedRow1 = "" Then
        Me.txtBox_custoInsumo.Enabled = True
        Me.txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
        Me.txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
        '===== teste =====
            If selectedRow = -1 And selectedRow1 = "" And Me.txtBox_nv4_4 = "" Then
            Me.txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
            Me.txtBox_custoInsumo.Enabled = False
            Me.txtBox_custoInsumo.BackColor = &H80000016
            End If
        '==============
        Else
        Me.txtBox_custoInsumo.Value = Format(CUSTO, "R$ #,##0.00")
        Me.txtBox_custoInsumo.Enabled = False
        Me.txtBox_custoInsumo.BackColor = &H80000016
        End If
        rs.Close
        conexao.Close


GoTo jumpit

If (i > 1) Then
here:
Me.txtBox_custoInsumo.Enabled = True
Me.txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
'MsgBox "Insira um novo custo!"
Me.txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
End If
Exit Sub
jumpit:




End Sub

Sub GetUnidade()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
Dim i As Integer
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Insumos"
sql1 = "t_Insumos"
i = 0
i = i + 1
Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex
Dim selectedRow1 As String
selectedRow1 = Me.txtBox_nv4_4.Value



   Dim valor As String
    If Not IsNull(Me.txtComboBox_nv4_4.Value) And Len(Me.txtComboBox_nv4_4.Value) > 0 Then
        valor = Me.txtComboBox_nv4_4.Value
        selectedRow = Me.txtComboBox_nv4_4.ListIndex
        sql1 = "SELECT Unidade FROM t_Insumos WHERE Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 Then

            selectedRow1 = Me.txtBox_nv4_4.Value
            sql1 = "SELECT Unidade FROM t_Insumos WHERE Insumo = '" & selectedRow1 & "';"
     
   End If




If (Me.txtBox_nv4_4.Value <> "") Then
Me.txtBox_un.Enabled = True
Me.txtBox_un.BackColor = RGB(255, 255, 255)
End If



rs.Open sql1, conexao
'grupo = rs.Fields("grupo").Value
On Error GoTo here
unidade = rs.Fields("Unidade").Value





If unidade <> "" Then
    If selectedRow1 = "" And selectedRow = -1 Then
    Me.txtBox_un.Value = ""
    Me.txtBox_un.Enabled = True
    Me.txtBox_un.BackColor = RGB(255, 255, 255)
    
    '=======TESTE===
    If selectedRow1 = "" And selectedRow = -1 And Me.txtBox_nv4_4 = "" Then
    Me.txtBox_un.Value = ""
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    End If
    
    Else
    Me.txtBox_un.Value = unidade
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    End If
    
    Else

    Me.txtBox_un.Value = "-"
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    End If

rs.Close
conexao.Close

GoTo jumpit
here:
    Me.txtBox_un.Value = ""
    Me.txtBox_un.Enabled = True
    Me.txtBox_un.BackColor = RGB(255, 255, 255)
Exit Sub

jumpit:

End Sub

'Verifica se há rendimento p serviços principais
Sub GetRendimento1()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
Dim id_master As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
ConectarBanco conexao


If txtBoxID_nv1.Value = "" Then
id_nv1 = 0
txtBoxID_nv1.Value = 0
Else
id_nv1 = txtBoxID_nv1.Value
End If

If txtBoxID_nv2.Value = "" Then
    id_nv2 = 0
    txtBoxID_nv2.Value = 0
Else
id_nv2 = txtBoxID_nv2.Value
End If

If txtBoxID_nv3.Value = "" Then
    txtBoxID_nv3.Value = 0
    id_nv3 = txtBoxID_nv3.Value
Else
id_nv3 = txtBoxID_nv3.Value
End If

'sql = "t_Servicos_Principais_Insumos"
sql1 = "t_Servicos_Principais_Insumos"

'id_master = id_nv1 & id_nv2 & id_nv3
id_master = id_nv1 & "-" & id_nv2 & "-" & id_nv3


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex


   Dim valor As String
    If Not IsNull(Me.txtComboBox_nv4_4.Value) And Len(Me.txtComboBox_nv4_4.Value) > 0 Then
        valor = Me.txtComboBox_nv4_4.Value
        selectedRow = Me.txtComboBox_nv4_4.ListIndex
        'FUNCIONANDO'sql1 = "SELECT rendimento FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & "ACRILICA SEMI-BRILHO 18LT" & "'"
        sql1 = "SELECT rendimento FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            'sql1 = "SELECT * FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "';"
            
   ElseIf Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 Then

        sql1 = "SELECT rendimento FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & Me.txtBox_nv4_4.Value & "'"
     
   'End If
    Else
    Exit Sub
    End If



Dim insumo As String



rs.Open sql1, conexao

On Error GoTo here
rendimento = rs.Fields("rendimento").Value
If (rendimento <> "") And (selectedRow <> -1 Or Me.txtBox_nv4_4 <> "") Then
    'rendimento = rs.Fields("rendimento").Value
    'Me.txtBox_un.Value = idNV1
    If Me.txtComboBox_nv4_4.Value <> "" Then
        insumo = Me.txtComboBox_nv4_4.Value
        c_rendimento = True
    Else
        insumo = Me.txtBox_nv4_4.Value
        c_rendimento = True
    End If
    
    'MsgBox "A soma dos IDs é existente, os valores existem na tabela, sendo ele = " & id_master & " / De valor de rendimento = " & rendimento & " Com nome de insumo = " & insumo
    txtBox_rendimento.Value = Format(rendimento, "#,##0.00")
    Me.txtBox_rendimento.BackColor = &H80000016
    Me.txtBox_rendimento.Enabled = False
'    Me.txtBox_rendimento.BackColor = RGB(255, 255, 255)
'    Me.txtBox_rendimento.Enabled = True
    rendimentoValue = True
    
Else
    'Me.txtBox_un.Value = "-"
here:
c_rendimento = False

    'MsgBox "A soma dos IDs não existente, será necessário preencher um valor para o rendimento"
    Me.txtBox_rendimento.BackColor = RGB(255, 255, 255)
    Me.txtBox_rendimento.Enabled = True
    rendimentoValue = True
    'Me.txtBox_un.Value = ""
    Me.txtBox_rendimento.Value = Format("", "#,##0.00")
    'Me.txtBox_custoInsumo.Value = ""
End If
rs.Close
conexao.Close


End Sub
'Verifica se há rendimentos para serviços diversos
Sub GetRendimentoDiversos()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
Dim id_master As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
ConectarBanco conexao


If txtBoxID_nv1.Value = "" Then
id_nv1 = 0
txtBoxID_nv1.Value = 0
Else
id_nv1 = txtBoxID_nv1.Value
End If

If txtBoxID_nv2.Value = "" Then
    id_nv2 = 0
    txtBoxID_nv2.Value = 0
Else
id_nv2 = txtBoxID_nv2.Value
End If

If txtBoxID_nv3.Value = "" Then
    txtBoxID_nv3.Value = 0
    id_nv3 = txtBoxID_nv3.Value
Else
id_nv3 = txtBoxID_nv3.Value
End If

'sql = "t_Servicos_Principais_Insumos"
sql1 = "t_Servicos_Principais_Insumos"

'id_master = 7 & 0 & id_nv3
id_master = 7 & "-" & 0 & "-" & id_nv3

Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = Me.txtComboBox_nv4_4.ListIndex



   Dim valor As String
    If Not IsNull(Me.txtComboBox_nv4_4.Value) And Len(Me.txtComboBox_nv4_4.Value) > 0 Then
        valor = Me.txtComboBox_nv4_4.Value
        selectedRow = Me.txtComboBox_nv4_4.ListIndex
        'FUNCIONANDO'sql1 = "SELECT rendimento FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & "ACRILICA SEMI-BRILHO 18LT" & "'"
        sql1 = "SELECT rendimento FROM t_Servicos_Diversos_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & Me.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            'sql1 = "SELECT * FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "';"
            
   ElseIf Not IsNull(Me.txtBox_nv4_4.Value) And Len(Me.txtBox_nv4_4.Value) > 0 Then

        sql1 = "SELECT rendimento FROM t_Servicos_Diversos_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & Me.txtBox_nv4_4.Value & "'"
     
   'End If
    Else
    Exit Sub
    End If
'===============================================================


Dim insumo As String



rs.Open sql1, conexao


'Me.ComboBox10.Value = grupo
On Error GoTo here
rendimento = rs.Fields("rendimento").Value
If (rendimento <> "") And (selectedRow <> -1 Or Me.txtBox_nv4_4 <> "") Then
    'rendimento = rs.Fields("rendimento").Value
    'Me.txtBox_un.Value = idNV1
    If Me.txtComboBox_nv4_4.Value <> "" Then
        insumo = Me.txtComboBox_nv4_4.Value
        c_rendimento = True
    Else
        insumo = Me.txtBox_nv4_4.Value
        c_rendimento = True
    End If
    
    'MsgBox "A soma dos IDs é existente, os valores existem na tabela, sendo ele = " & id_master & " / De valor de rendimento = " & rendimento & " Com nome de insumo = " & insumo
    txtBox_rendimento.Value = Format(rendimento, "#,##0.00")
    Me.txtBox_rendimento.BackColor = &H80000016
    Me.txtBox_rendimento.Enabled = False

    rendimentoValue = True
    
Else
    'Me.txtBox_un.Value = "-"
here:
c_rendimento = False
    'MsgBox "A soma dos IDs não existente, será necessário preencher um valor para o rendimento"
    Me.txtBox_rendimento.BackColor = RGB(255, 255, 255)
    Me.txtBox_rendimento.Enabled = True
    rendimentoValue = True
    'Me.txtBox_un.Value = ""
    Me.txtBox_rendimento.Value = Format("", "#,##0.00")
    'Me.txtBox_custoInsumo.Value = ""
End If
rs.Close
conexao.Close


End Sub


'|!|!|!|!| NIVEL 1 |!|!|!|!|
'IRÁ BUSCAR O VALOR DO LIST BOX, TRARÁ SE ELE EXISTE NA COLUNA descricaoNv1 OU SE NÃO EXISTE

Private Sub btnPesquisar_Click()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String

    Dim strTabela As String
    Dim strColuna As String
    Dim strValorPesquisa As String
    
    
    
      Dim valor As String
    If Not IsNull(Me.txtComboBox_nv1_1.Value) And Len(Me.txtComboBox_nv1_1.Value) > 0 Then
        valor = Me.txtComboBox_nv1_1.Value


   ' Definir o nome da tabela, a coluna e o valor a ser pesquisado
    strTabela = "t_Nivel1"
    strColuna = "descricaoNv1"
    strValorPesquisa = Me.txtComboBox_nv1_1.Value
    sql = "t_Nivel1"

     Set conexao = New ADODB.Connection
    ConectarBanco conexao

      sql = "SELECT * FROM " & strTabela & " WHERE " & strColuna & "='" & strValorPesquisa & "'"
    rs.Open sql, conexao
    
   ElseIf Not IsNull(Me.txtBox_nv1_1.Value) And Len(Me.txtBox_nv1_1.Value) > 0 Then
        valor = Me.txtBox_nv1_1.Value

   ' Definir o nome da tabela, a coluna e o valor a ser pesquisado
    strTabela = "t_Nivel1"
    strColuna = "descricaoNv1"
    strValorPesquisa = Me.txtBox_nv1_1.Value
    sql = "t_Nivel1"
    ' Abrir a conexão com o banco de dados
    'Set db = CurrentDb
     'Set conexao = CurrentDb
     'OU
     Set conexao = New ADODB.Connection
 ConectarBanco conexao

      sql = "SELECT * FROM " & strTabela & " WHERE " & strColuna & "='" & strValorPesquisa & "'"
    rs.Open sql, conexao
     
   End If
    

    On Error GoTo here
    If Not rs.EOF Then
        MsgBox "Valor encontrado na coluna " & strColuna & " da tabela " & strTabela, vbInformation, "Pesquisa"
        MsgBox "Dado NÃO inserido no DB"
        'Me.txtResultado.Value = True
    Else
        MsgBox "Valor não encontrado na coluna " & strColuna & " da tabela " & strTabela, vbInformation, "Pesquisa"
        MsgBox "Dado inserido no DB"
        'Me.txtResultado.Value = False
    End If
    
    ' Fechar o objeto Recordset e a conexão com o banco de dados
    rs.Close
    conexao.Close
    Set rs = Nothing
    Set db = Nothing
    
    GoTo jumpit
here:
MsgBox "Um dos campo do nível 1 precisa ser preenchido!"
Exit Sub
    
jumpit:

End Sub



Sub submit()


End Sub


'NIVEL 2
Private Sub txtComboBox_nv2_2_Change()

End Sub

'NIVEL 3
Private Sub txtComboBox_nv3_3_Change()

End Sub

Private Sub Menu_Adicionar_Click()
Me.MultiPage2(1).Visible = True
Me.MultiPage2(2).Visible = False
Me.MultiPage2(3).Visible = False
Me.MultiPage2.Value = 1

s_principais.Visible = True
s_diversos.Visible = True
s_terceiros.Visible = True
tituloGeral = "Adicionar Serviço Principal"

sPrincipais = True



edicao = False
remocao = False
adicao = True

Me.Editar.Visible = False
Me.Enviar.Visible = True

startAts_Principal



Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.EditarNiveis.Visible = False
Me.CheckBoxGeneric.Visible = False


    'Botão "Avançar" do Menu Adição
    CommandButton15.Visible = True

    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True

End Sub


Sub startAts_Principal()
'NIVEL1

Me.s_principais.BackColor = RGB(142, 162, 219)

Me.s_terceiros.ForeColor = &H80000008
Me.s_principais.Font.Bold = True





Me.s_diversos.BackColor = &H8000000F
Me.s_diversos.ForeColor = &H80000008
Me.s_diversos.Font.Bold = False

Me.s_terceiros.BackColor = &H8000000F
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = False



Me.txtBox_nv1_1.Value = ""
Me.txtComboBox_nv1_1.Value = ""
Me.txtBox_nv2_2.Value = ""
Me.txtComboBox_nv2_2.Value = ""
Me.txtBox_nv3_3.Value = ""
Me.txtComboBox_nv3_3.Value = ""
Me.txtBox_nv4_4.Value = ""
Me.txtComboBox_nv4_4.Value = ""
Me.txtBox_un.Value = ""
Me.txtBox_rendimento.Value = ""
Me.txtBox_custoInsumo.Value = ""
Me.txtBox_pvs.Value = ""
Me.txtBox_cmo.Value = ""

sPrincipais = True
sTerceiros = False
sDiversos = False
Me.CarregarNivelUm
Adicionar.CarregarNivelDois
Me.CarregarNivelTres

Me.txtBox_nv1_1.Enabled = True
Me.txtBox_nv1_1.BackColor = RGB(255, 255, 255)
'Me.txtBox_nv1_1.Font.Bold = True

Me.txtBoxID_nv1.Enabled = True
Me.txtBoxID_nv1.BackColor = RGB(255, 255, 255)

Me.txtComboBox_nv1_1.Enabled = True
Me.txtComboBox_nv1_1.BackColor = RGB(255, 255, 255)

cbBox_nv1.optionButton_nv1_1.Value = False
cbBox_nv1.optionButton_nv1_2.Value = False

Me.cbBox_nv1.Enabled = True

'NIVEL 2
Me.txtBox_nv2_2.Enabled = True
Me.txtBox_nv2_2.BackColor = RGB(255, 255, 255)
'Me.txtBox_nv2_2.Font.Bold = True

Me.cbBox_nv2_2.Enabled = True
'Me.cbBox_nv2_2.BackColor = RGB(255, 255, 255)

Me.txtComboBox_nv2_2.Enabled = True
Me.txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)

cbBox_nv2_2.optionButton_nv2_4.Value = False
cbBox_nv2_2.optionButton_nv2_3.Value = False


'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = True
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016

'INSUMO
    Me.txtBox_nv4_4.Enabled = True
    Me.txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = True
    'Me.cbBox_nv4_4.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    
    Me.txtBox_rendimento.Enabled = False
    Me.txtBox_rendimento.BackColor = &H80000016
    
    Me.txtBox_custoInsumo.Enabled = False
    Me.txtBox_custoInsumo.BackColor = &H80000016
    
    Me.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    
    Me.txtBox_cmo.Enabled = False
    Me.txtBox_cmo.BackColor = &H80000016
    
    
    
Me.optionButton_nv1_2.Value = True
Me.optionButton_nv2_3.Value = True
Me.optionButton_nv3_6.Value = True
Me.optionButton_nv4_8.Value = True
End Sub

'SERVIÇO DE TERCEIRO BOTÃO
Private Sub s_terceiros_Click()

sTerceiros = True
sPrincipais = False
sDiversos = False
tituloGeral = "Adicionar Serviços de Terceiros"

If Visible = True Then

Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.EditarNiveis.Visible = False

EditarNiveis.Visible = False
CommandButton15.Visible = True


End If

Me.txtBoxID_nv1 = ""
Me.txtBoxID_nv2 = ""
Me.txtBoxID_nv3 = ""
Me.txtBoxID_nv4 = ""

'Me.s_terceiros.BackColor = RGB(255, 242, 204)
Me.s_terceiros.BackColor = RGB(255, 242, 204)
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = True
Me.MultiPage2.Value = 2
'Me.MultiPage1.BackColor = RGB(255, 249, 231)

Me.s_principais.BackColor = &H8000000F
Me.s_principais.ForeColor = &H80000008
Me.s_principais.Font.Bold = False

Me.s_diversos.BackColor = &H8000000F
Me.s_diversos.ForeColor = &H80000008
Me.s_diversos.Font.Bold = False



Me.txtBox_nv1_1.Value = ""
Me.txtComboBox_nv1_1.Value = ""
Me.txtBox_nv2_2.Value = ""
Me.txtComboBox_nv2_2.Value = ""
Me.txtBox_nv3_3.Value = ""
Me.txtComboBox_nv3_3.Value = ""
Me.txtBox_nv4_4.Value = ""
Me.txtComboBox_nv4_4.Value = ""
Me.txtBox_un.Value = ""
Me.txtBox_rendimento.Value = ""
Me.txtBox_custoInsumo.Value = ""
Me.txtBox_pvs.Value = ""
Me.txtBox_cmo.Value = ""



id_nv1 = 9
id_nv2 = 0
'Me.CarregarNivelUm
'NIVEL 1
    Me.txtBox_nv1_1.Enabled = False
    Me.txtBox_nv1_1.BackColor = &H80000016
    'Me.txtBox_nv1_1.Font.Bold = True
    
    Me.txtBoxID_nv1.Enabled = False
    Me.txtBoxID_nv1.BackColor = &H80000016
    
    Me.txtComboBox_nv1_1.Enabled = False
    Me.txtComboBox_nv1_1.BackColor = &H80000016
    
    Me.cbBox_nv1.Enabled = False

'NIVEL 2
    Me.txtBox_nv2_2.Enabled = False
    Me.txtBox_nv2_2.BackColor = &H80000016
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv2_2.Enabled = False
    'Me.cbBox_nv2_2.BackColor = &H80000016
    
    Me.txtComboBox_nv2_2.Enabled = False
    Me.txtComboBox_nv2_2.BackColor = &H80000016

'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = False
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016

'INSUMO
    Me.txtBox_nv4_4.Enabled = False
    Me.txtBox_nv4_4.BackColor = &H80000016
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = False
    'Me.cbBox_nv4_4.BackColor = &H80000016
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    
    Me.txtBox_rendimento.Enabled = False
    Me.txtBox_rendimento.BackColor = &H80000016
    
    Me.txtBox_custoInsumo.Enabled = False
    Me.txtBox_custoInsumo.BackColor = &H80000016
    
    Me.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    
    Me.txtBox_cmo.Enabled = False
    Me.txtBox_cmo.BackColor = &H80000016
    
    
'    Me.optionButton_nv3_5.Locked = True
    Me.optionButton_nv4_7 = False
    Me.optionButton_nv4_8 = False
    Me.optionButton_nv1_1 = False
    Me.optionButton_nv1_2 = False
    Me.optionButton_nv2_3 = False
    Me.optionButton_nv2_4 = False
    Me.optionButton_nv3_6 = True
    
    
    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True
    
    
    CommandButton15.Visible = False
    Enviar.Visible = True

End Sub

Private Sub GoMenu_Click()
'Módulo1.CloseForm
'Módulo1.OpenForm
Me.MultiPage2.Value = 0
End Sub

Private Sub CommandButton2_Click()



Me.MultiPage2(1).Visible = False
Me.MultiPage2(2).Visible = True
Me.MultiPage2(3).Visible = False
Me.MultiPage2.Value = 2

Me.MultiPage4.Value = 0
Me.MultiPage4(0).Visible = True
Me.MultiPage4(1).Visible = False
Me.MultiPage4(2).Visible = False



Me.ex_btnNv1.Enabled = True
Me.ex_btnNv2.Enabled = True
Me.ex_btnNv3.Enabled = True
Me.ex_btnNv4.Enabled = False
Me.ex_btnEstrutura.Enabled = True



ex_sPrincipais = True
ex_sDiversos = False
ex_sTerceiros = False
ex_sGeneralInsumo = False

Me.ex_s_principais.BackColor = RGB(142, 162, 219)
Me.ex_s_terceiros.ForeColor = &H80000008
Me.ex_s_principais.Font.Bold = True



Me.ex_s_diversos.BackColor = &H8000000F
Me.ex_s_diversos.ForeColor = &H80000008
Me.ex_s_diversos.Font.Bold = False

Me.ex_s_terceiros.BackColor = &H8000000F
Me.ex_s_terceiros.ForeColor = &H80000008
Me.ex_s_terceiros.Font.Bold = False

Me.ex_insumoGerenal.BackColor = &H8000000F
Me.ex_insumoGerenal.ForeColor = &H80000008
Me.ex_insumoGerenal.Font.Bold = False


Me.ed_txtBox_nv1_1.Value = ""
Me.ed_txtComboBox_nv1_1.Value = ""
Me.ed_txtBox_nv2_2.Value = ""
Me.ed_txtComboBox_nv2_2.Value = ""
Me.ed_txtBox_nv3_3.Value = ""
Me.ed_txtComboBox_nv3_3.Value = ""
Me.ed_txtBox_nv4_4.Value = ""
Me.ed_txtComboBox_nv4_4.Value = ""
Me.ed_txtBox_un.Value = ""
Me.ed_txtBox_rendimento.Value = ""
Me.ed_txtBox_custoInsumo.Value = ""
Me.ed_txtBox_pvs.Value = ""
Me.ed_txtBox_cmo.Value = ""

GlobalComboBoxValue = ""
GlobalServiceType = "Servicos Principais"
GlobalTable = ""
Me.ex_ListBox.Clear
End Sub

Private Sub Menu_Editar_Click()


Me.MultiPage2(1).Visible = False
Me.MultiPage2(2).Visible = False
Me.MultiPage2(3).Visible = True
Me.MultiPage2.Value = 3

Me.MultiPage3.Value = 0
Me.MultiPage3(0).Visible = True
Me.MultiPage3(1).Visible = False

Me.s_principais.Visible = False
Me.s_diversos.Visible = False
Me.s_terceiros.Visible = False

edicao = True
remocao = False
adicao = False
'



Me.Editar.Visible = True
Me.Enviar.Visible = False

Me.ed_s_principais = True
Me.s_principais.BackColor = RGB(142, 162, 219)
Me.s_terceiros.ForeColor = &H80000008
Me.s_principais.Font.Bold = True



Me.s_diversos.BackColor = &H8000000F
Me.s_diversos.ForeColor = &H80000008
Me.s_diversos.Font.Bold = False

Me.s_terceiros.BackColor = &H8000000F
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = False



saveObjectsPosition



End Sub

Private Sub CommandButton4_Click()
Me.MultiPage1(0).Visible = True
Me.MultiPage1(1).Visible = False
Me.MultiPage1(2).Visible = False
Me.MultiPage1(3).Visible = True
Me.MultiPage1.Value = 3

End Sub

Private Sub CommandButton5_Click()
Me.MultiPage1(0).Visible = True
Me.MultiPage1(1).Visible = False
Me.MultiPage1(2).Visible = True
Me.MultiPage1(3).Visible = False
Me.MultiPage1.Value = 2
End Sub

Private Sub CommandButton6_Click()
Me.MultiPage1(0).Visible = True
Me.MultiPage1(1).Visible = True
Me.MultiPage1(2).Visible = False
Me.MultiPage1(3).Visible = False
Me.MultiPage1.Value = 1
End Sub



'SERVIÇO PRINCIPAL BOTÃO
Private Sub s_principais_Click()
'NIVEL1

If sTerceiros = True Then
Me.Enviar.Visible = False
EditarNiveis.Visible = False
CommandButton15.Visible = True
End If

sPrincipais = True
sTerceiros = False
sDiversos = False
tituloGeral = "Adicionar Serviços Principais"

If Visible = True Then



Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.EditarNiveis.Visible = False

EditarNiveis.Visible = False
CommandButton15.Visible = True
End If

Me.txtBoxID_nv1 = ""
Me.txtBoxID_nv2 = ""
Me.txtBoxID_nv3 = ""
Me.txtBoxID_nv4 = ""

'Me.s_diversos.BackColor = RGB(35, 55, 100)
'Me.s_principais.BackColor = RGB(142, 162, 219)
Me.s_principais.BackColor = RGB(142, 162, 219)
Me.s_terceiros.ForeColor = &H80000008
Me.s_principais.Font.Bold = True




'Me.MultiPage1.BackColor = RGB(142, 169, 219)


Me.s_diversos.BackColor = &H8000000F
Me.s_diversos.ForeColor = &H80000008
Me.s_diversos.Font.Bold = False

Me.s_terceiros.BackColor = &H8000000F
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = False



Me.txtBox_nv1_1.Value = ""
Me.txtComboBox_nv1_1.Value = ""
Me.txtBox_nv2_2.Value = ""
Me.txtComboBox_nv2_2.Value = ""
Me.txtBox_nv3_3.Value = ""
Me.txtComboBox_nv3_3.Value = ""
Me.txtBox_nv4_4.Value = ""
Me.txtComboBox_nv4_4.Value = ""
Me.txtBox_un.Value = ""
Me.txtBox_rendimento.Value = ""
Me.txtBox_custoInsumo.Value = ""
Me.txtBox_pvs.Value = ""
Me.txtBox_cmo.Value = ""


Me.CarregarNivelUm
Adicionar.CarregarNivelDois
Me.CarregarNivelTres
Adicionar.CarregarInsumos

Me.txtBox_nv1_1.Enabled = True
Me.txtBox_nv1_1.BackColor = RGB(255, 255, 255)
'Me.txtBox_nv1_1.Font.Bold = True

Me.txtBoxID_nv1.Enabled = True
Me.txtBoxID_nv1.BackColor = RGB(255, 255, 255)

Me.txtComboBox_nv1_1.Enabled = True
Me.txtComboBox_nv1_1.BackColor = RGB(255, 255, 255)

cbBox_nv1.optionButton_nv1_1.Value = False
cbBox_nv1.optionButton_nv1_2.Value = False

Me.cbBox_nv1.Enabled = True

'NIVEL 2
Me.txtBox_nv2_2.Enabled = True
Me.txtBox_nv2_2.BackColor = RGB(255, 255, 255)
'Me.txtBox_nv2_2.Font.Bold = True

Me.cbBox_nv2_2.Enabled = True
'Me.cbBox_nv2_2.BackColor = RGB(255, 255, 255)

Me.txtComboBox_nv2_2.Enabled = True
Me.txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)

cbBox_nv2_2.optionButton_nv2_4.Value = False
cbBox_nv2_2.optionButton_nv2_3.Value = False


'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = True
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016

'INSUMO
    Me.txtBox_nv4_4.Enabled = True
    Me.txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = True
    'Me.cbBox_nv4_4.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    
    Me.txtBox_rendimento.Enabled = False
    Me.txtBox_rendimento.BackColor = &H80000016
    
    Me.txtBox_custoInsumo.Enabled = False
    Me.txtBox_custoInsumo.BackColor = &H80000016
    
    Me.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    
    Me.txtBox_cmo.Enabled = False
    Me.txtBox_cmo.BackColor = &H80000016
    
    
    
Me.optionButton_nv1_2.Value = True
Me.optionButton_nv2_3.Value = True
Me.optionButton_nv3_6.Value = True
Me.optionButton_nv4_8.Value = True


    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True
End Sub
'SERVIÇOS DIVERSOS
Private Sub s_diversos_Click()
'NIVEL 1


''\!/---BLOQUEIA O ACESSO DO USUÁRIO---\!/
'stopApplication
'If stop_Application Then Exit Sub
''\!/----------------------------------\!/



If sTerceiros = True Then
Me.Enviar.Visible = False
EditarNiveis.Visible = False
CommandButton15.Visible = True
End If



sDiversos = True
sTerceiros = False
sPrincipais = False
tituloGeral = "Adicionar Serviços Diversos"

If Visible = True Then
'MsgBox "Clique em Editar Niveis para poder prosseguir."
'Exit Sub
'============= teste =======
Me.Label27.Visible = False
Me.txtBox_un.Visible = False
Me.Label38.Visible = False
Me.txtBox_rendimento.Visible = False
Me.Label36.Visible = False
Me.txtBox_cmo.Visible = False
Me.Label44.Visible = False
Me.txtBox_custoInsumo.Visible = False
Me.Label35.Visible = False
Me.txtBox_pvs.Visible = False
Me.Enviar.Visible = False
Me.EditarNiveis.Visible = False
'=======================
EditarNiveis.Visible = False
CommandButton15.Visible = True
End If


Me.txtBoxID_nv1 = ""
Me.txtBoxID_nv2 = ""
Me.txtBoxID_nv3 = ""
Me.txtBoxID_nv4 = ""

'Me.s_diversos.BackColor = RGB(255, 230, 153)
Me.s_diversos.BackColor = RGB(255, 230, 153)
Me.s_terceiros.ForeColor = &H80000008
Me.s_diversos.Font.Bold = True
Me.MultiPage2.Value = 2
'Me.MultiPage1.BackColor = RGB(142, 162, 219)

Me.s_principais.BackColor = &H8000000F
Me.s_principais.ForeColor = &H80000008
Me.s_principais.Font.Bold = False

Me.s_terceiros.BackColor = &H8000000F
Me.s_terceiros.ForeColor = &H80000008
Me.s_terceiros.Font.Bold = False






Me.CarregarNivelTres
Adicionar.CarregarInsumos

id_nv1 = 7
id_nv2 = 0

Me.txtBox_nv1_1.Enabled = False
Me.txtBox_nv1_1.BackColor = &H80000016 '&H80000016
Me.txtBox_nv1_1.Value = ""
'Me.txtBox_nv1_1.Font.Bold = True

Me.txtBoxID_nv1.Enabled = False
Me.txtBoxID_nv1.BackColor = &H80000016

Me.txtComboBox_nv1_1.Enabled = False
Me.txtComboBox_nv1_1.BackColor = &H80000016
Me.txtComboBox_nv1_1.Value = ""
'NIVEL 2
Me.txtBox_nv2_2.Enabled = False
Me.txtBox_nv2_2.BackColor = &H80000016
Me.txtBox_nv2_2.Value = ""
'Me.txtBox_nv2_2.Font.Bold = True

Me.cbBox_nv2_2.Enabled = False
'Me.cbBox_nv2_2.BackColor = &H80000016

Me.txtComboBox_nv2_2.Enabled = False
Me.txtComboBox_nv2_2.BackColor = &H80000016
Me.txtComboBox_nv2_2.Value = ""

Me.cbBox_nv1.Enabled = False

'NIVEL 3

    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
    Me.txtBox_nv3_3.Value = ""
    'Me.txtBox_nv3_3.Font.Bold = True
    
    Me.cbBox_nv3_3.Enabled = True
    'Me.cbBox_nv3_3.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtComboBox_nv3_3.BackColor = &H80000016
    Me.txtComboBox_nv3_3.Value = ""
    


    

    'INSUMO
    Me.txtBox_nv4_4.Enabled = True
    Me.txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    Me.txtBox_nv4_4.Value = ""
    'Me.txtBox_nv2_2.Font.Bold = True
    
    Me.cbBox_nv4_4.Enabled = True
    'Me.cbBox_nv4_4.BackColor = RGB(255, 255, 255)
    
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    
    Me.txtBox_un.Enabled = False
    Me.txtBox_un.BackColor = &H80000016
    
    Me.txtBox_rendimento.Enabled = False
    Me.txtBox_rendimento.BackColor = &H80000016
    
    Me.txtBox_custoInsumo.Enabled = False
    Me.txtBox_custoInsumo.BackColor = &H80000016
    
    Me.txtBox_pvs.Enabled = False
    Me.txtBox_pvs.BackColor = &H80000016
    
    Me.txtBox_cmo.Enabled = False
    Me.txtBox_cmo.BackColor = &H80000016
    
    

Me.optionButton_nv1_2.Value = False
Me.optionButton_nv2_3.Value = False
Me.optionButton_nv3_6.Value = True
Me.optionButton_nv4_8.Value = True
Me.optionButton_nv1_1.Value = False
Me.optionButton_nv2_4.Value = False

    optionButton_nv1_1.Enabled = True
    optionButton_nv1_2.Enabled = True
    optionButton_nv2_3.Enabled = True
    optionButton_nv2_4.Enabled = True
    optionButton_nv3_5.Enabled = True
    optionButton_nv3_6.Enabled = True
    optionButton_nv4_7.Enabled = True
    optionButton_nv4_8.Enabled = True


End Sub

'nivel 1
Private Sub txtBoxID_nv1_Click()

End Sub

'NIVEL 2
Private Sub cbBox_nv2_2_Click()

End Sub


'|!|!|!|!| NIVEL 2 |!|!|!|!|
Private Sub optionButton_nv2_4_Click()
    Me.txtBox_nv2_2.Value = ""
    Me.txtComboBox_nv2_2.Enabled = True
    Me.txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)
    Me.txtBox_nv2_2.BackColor = &H80000016
    Me.txtBox_nv2_2.Enabled = False
    
    
End Sub
'|!|!|!|!| NIVEL 2 |!|!|!|!|
Private Sub optionButton_nv2_3_Click()
    Me.txtComboBox_nv2_2.Value = ""
    Me.txtComboBox_nv2_2.Enabled = False
    Me.txtComboBox_nv2_2.BackColor = &H80000016
    Me.txtBox_nv2_2.Enabled = True
    Me.txtBox_nv2_2.BackColor = RGB(255, 255, 255)
End Sub


'NIVEL 3
Private Sub cbBox_nv3_3_Click()

End Sub

'|!|!|!|!| NIVEL 3 |!|!|!|!|
Private Sub optionButton_nv3_5_Click()
    Me.txtBox_nv3_3.Value = ""
    Me.txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
    Me.txtComboBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.Enabled = False
    Me.txtBox_nv3_3.BackColor = &H80000016
End Sub
'|!|!|!|!| NIVEL 3 |!|!|!|!|
Private Sub optionButton_nv3_6_Click()
    Me.txtComboBox_nv3_3.Value = ""
    Me.txtComboBox_nv3_3.BackColor = &H80000016
    Me.txtComboBox_nv3_3.Enabled = False
    Me.txtBox_nv3_3.Enabled = True
    Me.txtBox_nv3_3.BackColor = RGB(255, 255, 255)
End Sub


'INSUMO
Private Sub cbBox_nv4_4_Click()

End Sub

'|!|!|!|!| NIVEL 4 |!|!|!|!|
Private Sub optionButton_nv4_7_Click()
    Me.txtBox_nv4_4.Value = ""
    Me.txtComboBox_nv4_4.Enabled = True
    Me.txtComboBox_nv4_4.BackColor = RGB(255, 255, 255)
    Me.txtBox_nv4_4.BackColor = &H80000016
    Me.txtBox_nv4_4.Enabled = False

End Sub
'|!|!|!|!| NIVEL 4 |!|!|!|!|
Private Sub optionButton_nv4_8_Click()
    Me.txtComboBox_nv4_4.Value = ""
    Me.txtComboBox_nv4_4.BackColor = &H80000016
    Me.txtComboBox_nv4_4.Enabled = False
    Me.txtBox_nv4_4.Enabled = True
    Me.txtBox_nv4_4.BackColor = RGB(255, 255, 255)
    
    
    
    Me.txtBox_un.Value = ""
    Me.txtBox_rendimento.Value = ""
    Me.txtBox_custoInsumo.Value = ""
    Me.txtBoxID_nv4.Value = ""
End Sub

Private Sub MultiPage1_Change()





End Sub



Private Sub MultiPage2_Change()


End Sub



Private Sub plusIt_Click()



 If Me.Label45.Visible = True Then
        Me.Label45.Visible = False
    Else
        Me.Label45.Visible = True
    End If
End Sub


Private Sub txtBox_nv2_2_Change()

End Sub

Private Sub txtBox_nv2_2_Enter()

End Sub

'NIVEL 3
Private Sub txtBox_nv3_3_Change()


End Sub

'NIVEL 3
Private Sub txtBox_nv3_3_Enter()
'    If Me.txtBox_nv3_3.Value = "Digite o serviço aqui" Then
'    Me.txtBox_nv3_3.Value = ""
'    End If
End Sub

'INSUMO
Private Sub txtBox_nv4_4_Change()

End Sub

'INSUMO
Private Sub txtBox_nv4_4_Enter()
'    If Me.txtBox_nv4_4.Value = "Digite o serviço aqui" Then
'    Me.txtBox_nv4_4.Value = ""
'    End If
End Sub



Private Sub ed_txtBox_pvs_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que 0,1 na caixa de texto'
    Select Case KeyAscii
    
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'
            If KeyAscii = 44 And InStr(Me.ed_txtBox_pvs.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else

                KeyAscii = 0
                Beep 'emite um som de aviso'
            'End If
    End Select
End Sub




Private Sub txtBox_pvs_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números maiores que zero na caixa de texto'
    Select Case KeyAscii

        Case 48 To 57, 44 'números de 0 a 9 e vírgula'
            If KeyAscii = 44 And InStr(Me.txtBox_pvs.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'

            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select
End Sub






Private Sub txtBox_custoInsumo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que zero na caixa de texto'
    Select Case KeyAscii
    
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'
'            If Len(Me.txtBox_custoInsumo.Text) = 0 And KeyAscii = 48 Then
'                'impede a entrada de zeros no início do número'
'                KeyAscii = 0
'                Beep 'emite um som de aviso'
            If KeyAscii = 44 And InStr(Me.txtBox_custoInsumo.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select
End Sub


Private Sub ed_txtBox_custoInsumo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que zero na caixa de texto'
    Select Case KeyAscii
    
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'
            'If Len(Me.ed_txtBox_custoInsumo.Text) = 0 And KeyAscii = 48 Then
                'impede a entrada de zeros no início do número'
                'KeyAscii = 0
                'Beep 'emite um som de aviso'
            If KeyAscii = 44 And InStr(Me.ed_txtBox_custoInsumo.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select

End Sub







Private Sub txtBox_rendimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que zero na caixa de texto'
    Select Case KeyAscii
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'
            If KeyAscii = 44 And InStr(Me.txtBox_rendimento.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select
End Sub



Private Sub ed_txtBox_rendimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que zero na caixa de texto'
    Select Case KeyAscii
    
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'
            If Len(Me.ed_txtBox_rendimento.Text) = 0 And KeyAscii = 48 Then
                'impede a entrada de zeros no início do número'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            ElseIf KeyAscii = 44 And InStr(Me.ed_txtBox_rendimento.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select
End Sub



Private Sub txtBox_cmo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que zero na caixa de texto'
    Select Case KeyAscii
    
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'

            If KeyAscii = 44 And InStr(Me.txtBox_cmo.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select
End Sub



Private Sub ed_txtBox_cmo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Permite apenas a entrada de números e vírgulas maiores que zero na caixa de texto'
    Select Case KeyAscii
    
        Case 48 To 57, 44 'números de 0 a 9 e vírgula'

            If KeyAscii = 44 And InStr(Me.ed_txtBox_cmo.Text, ",") > 0 Then
                'impede a entrada de mais de uma vírgula'
                KeyAscii = 0
                Beep 'emite um som de aviso'
            End If
        Case 8, 13, 27 'backspace, enter, escape'
            'Não faz nada'
        Case Else
            'impede a entrada de outros caracteres'
            KeyAscii = 0
            Beep 'emite um som de aviso'
    End Select
End Sub

'INSUMO
Private Sub txtBox_un_Change()

End Sub



Private Sub txtBox_un_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Verifica se o comprimento atual do texto na caixa de texto é igual a 3 '
    If Len(Me.txtBox_un.Text) = 3 Then
        ' Impede a inserção de novos caracteres '
        KeyAscii = 0
        Beep ' emite um som de aviso '
    Else
        ' Permite apenas a inserção de letras como primeiro caractere '
        If Len(Me.txtBox_un.Text) = 0 Then
            If Not IsLetter(Chr(KeyAscii)) Then
                KeyAscii = 0
                Beep ' emite um som de aviso '
            End If
        Else
            ' Permite letras e números '
            Select Case KeyAscii
                Case 48 To 57 ' números de 0 a 9 '
                    ' Permite números '
                Case 65 To 90, 97 To 122 ' letras '
                    ' Permite letras '
                Case 8, 13, 27 ' backspace, enter, escape '
                    ' Não faz nada '
                Case Else
                    KeyAscii = 0
                    Beep ' emite um som de aviso '
            End Select
        End If
    End If
End Sub

Private Function IsLetter(ByVal Char As String) As Boolean
    ' Verifica se o caractere é uma letra '
    IsLetter = Char Like "[a-zA-Z]"
End Function

Private Sub ed_txtBox_un_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Verifica se o comprimento atual do texto na caixa de texto é igual a 3 '
    If Len(Me.ed_txtBox_un.Text) = 3 Then
        ' Impede a inserção de novos caracteres '
        KeyAscii = 0
        Beep ' emite um som de aviso '
    Else
        ' Permite apenas a inserção de letras como primeiro caractere '
        If Len(Me.ed_txtBox_un.Text) = 0 Then
            If Not IsLetter(Chr(KeyAscii)) Then
                KeyAscii = 0
                Beep ' emite um som de aviso '
            End If
        Else
            ' Permite letras e números '
            Select Case KeyAscii
                Case 48 To 57 ' números de 0 a 9 '
                    ' Permite números '
                Case 65 To 90, 97 To 122 ' letras '
                    ' Permite letras '
                Case 8, 13, 27 ' backspace, enter, escape '
                    ' Não faz nada '
                Case Else
                    KeyAscii = 0
                    Beep ' emite um som de aviso '
            End Select
        End If
    End If
End Sub

'INSUMO
Private Sub txtBox_pvs_Change()

att_pvs

End Sub
'INSUMO
Private Sub txtBox_cmo_Change()

End Sub

''INSUMO
Private Sub txtBox_rendimento_Change()
   rendimentoValue = False
End Sub


Private Sub txtBox_custoInsumo_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.txtBox_custoInsumo.Value) Then
        Me.txtBox_custoInsumo.Value = Format(Me.txtBox_custoInsumo.Value, "R$ #,##0.00")
    End If
End Sub

Private Sub txtBox_pvs_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.txtBox_pvs.Value) Then
        Me.txtBox_pvs.Value = Format(Me.txtBox_pvs.Value, "R$ #,##0.00")
    End If
End Sub

Private Sub txtBox_cmo_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.txtBox_cmo.Value) Then
        Me.txtBox_cmo.Value = Format(Me.txtBox_cmo.Value, "R$ #,##0.00")
    End If
End Sub

Private Sub txtBox_rendimento_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.txtBox_rendimento.Value) Then
        Me.txtBox_rendimento.Value = Format(Me.txtBox_rendimento.Value, "#,##0.00")
    End If
End Sub


'===
Private Sub ed_txtBox_custoInsumo_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.ed_txtBox_custoInsumo.Value) Then
        Me.ed_txtBox_custoInsumo.Value = Format(Me.ed_txtBox_custoInsumo.Value, "R$ #,##0.00")
    End If
End Sub

Private Sub ed_txtBox_pvs_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.ed_txtBox_pvs.Value) Then
        Me.ed_txtBox_pvs.Value = Format(Me.ed_txtBox_pvs.Value, "R$ #,##0.00")
    End If
End Sub

Private Sub ed_txtBox_cmo_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.ed_txtBox_cmo.Value) Then
        Me.ed_txtBox_cmo.Value = Format(Me.ed_txtBox_cmo.Value, "R$ #,##0.00")
    End If
End Sub

Private Sub ed_txtBox_rendimento_AfterUpdate()
    ' Formata o conteúdo do TextBox como moeda
    If IsNumeric(Me.ed_txtBox_rendimento.Value) Then
        Me.ed_txtBox_rendimento.Value = Format(Me.ed_txtBox_rendimento.Value, "#,##0.00")
    End If
End Sub

Private Sub txtBoxID_nv1_Change()

End Sub

Private Sub txtBoxID_nv2_Change()

End Sub

Private Sub txtBoxID_nv3_Change()

End Sub

Private Sub txtBoxID_nv4_Change()

End Sub

'nivel 1
Private Sub txtBox_nv1_1_Change()

End Sub



Private Sub TextBox9_Change()

End Sub

'Private Sub txtComboBox_nv1_1_Change()
'
'End Sub

Private Sub UserForm_Activate()
Me.MultiPage2(1).Visible = True
Me.MultiPage2(1).Visible = False
Me.MultiPage2(2).Visible = False
Me.MultiPage2(3).Visible = False

'With frmTelaCheia
'    Width = Application.Width
'    Height = Application.Height
'    Left = Application.Left
'    Top = Application.Top
'End With

With frmTelaCheia
    Width = 650
    Height = 555
    Left = Application.Left
    Top = Application.Top
End With


End Sub






Sub CarregarNivelUm()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String


Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel1"


'If sPrincipais = True Then
rs.Open "select descricaoNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv1", conexao, 3, 3
'End If





Do Until rs.EOF
UserForm2.txtComboBox_nv1_1.AddItem rs!descricaoNv1

rs.MoveNext
Loop

conexao.Close



End Sub


Sub CarregarNivelTres()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel3"




If sPrincipais = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv3", conexao, 3, 3
End If


If sTerceiros = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS TERCEIROS' order BY idNv3", conexao, 3, 3
End If

If sDiversos = True Then

rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv3", conexao, 3, 3
End If


UserForm2.txtComboBox_nv3_3.Clear
Do Until rs.EOF

UserForm2.txtComboBox_nv3_3.AddItem rs!descricaoNv3

rs.MoveNext
Loop

conexao.Close



End Sub

Private Sub txtBox_nv1_1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim InputString As String
    Dim OutputString As String
    Dim i As Long
    Dim AccentedChars As String
    Dim UnaccentedChars As String
    
    ' Defina a string de entrada como o valor do TextBox
    InputString = txtBox_nv1_1.Value
    
    ' Crie uma lista de caracteres acentuados e uma lista de caracteres correspondentes sem acento
    AccentedChars = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïñòóôõöùúûüýÿ"
    UnaccentedChars = "AAAAAACEEEEIIIINOOOOOUUUUYaaaaaaceeeeiiiinooooouuuuyy"
    
    ' Converta a string de entrada para maiúsculas e remova acentos
    OutputString = UCase(InputString)
    For i = 1 To Len(AccentedChars)
        OutputString = Replace(OutputString, Mid(AccentedChars, i, 1), Mid(UnaccentedChars, i, 1))
    Next i
    
    ' Defina o valor do TextBox como a string resultante
    txtBox_nv1_1.Value = OutputString
End Sub


Sub changeIndex()

'(BOTOES) TIPOS DE SERVIÇO'
Me.s_principais.TabIndex = 0
Me.s_diversos.TabIndex = 1
Me.s_terceiros.TabIndex = 2
'(BOTAO) Voltar ao menu
Me.GoMenu.TabIndex = 3
'NÍVEL 1 - seleção serviços
Me.txtBox_nv1_1.TabIndex = 4
Me.txtComboBox_nv1_1.TabIndex = 5
'NÍVEL 1 - textos serviços
'NÍVEL 1

'                            ID = txtBoxID_nv1
'Frame1                  Ex: (cbBox_nv1)
'    OptionButton1 optionButton_nv1_1
'    OptionButton2 optionButton_nv1_2
'------------------------------------------------------------------------------------------------------------------------------
'NÍVEL 2
Me.txtComboBox_nv2_2.TabIndex = 6
Me.txtComboBox_nv2_2.TabIndex = 7
'                            ID = TextBox23              txtBoxID_nv2
'Frame2 cbBox_nv2_2
Me.optionButton_nv2_3.TabIndex = 8
Me.optionButton_nv2_4.TabIndex = 9
'-----------------------------------------------------------------------------------------------------------------------------
'NÍVEL 3
Me.txtBox_nv3_3.TabIndex = 10
Me.txtComboBox_nv3_3.TabIndex = 11
'                        ID = TextBox24                          txtBoxID_nv3
'Frame3 cbBox_nv3_3
'Me.optionButton_nv3_5.TabIndex = 9
'Me.optionButton_nv3_6.TabIndex = 10
'-----------------------------------------------------------------------------------------------------------------------------
'insumo
Me.txtBox_nv4_4.TabIndex = 12
Me.txtComboBox_nv4_4.TabIndex = 13
'                    ID = TextBox25      txtBoxID_nv4
'
'Frame4 cbBox_nv4_4
'    OptionButton7 optionButton_nv4_7
'    OptionButton8 optionButton_nv4_8
'
'
Me.txtBox_un.TabIndex = 14
Me.txtBox_rendimento.TabIndex = 15
Me.txtBox_custoInsumo.TabIndex = 16
'
'
Me.txtBox_pvs.TabIndex = 18
Me.txtBox_cmo.TabIndex = 17

End Sub
