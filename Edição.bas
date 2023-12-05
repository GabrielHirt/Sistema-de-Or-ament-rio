Attribute VB_Name = "Edição"
Public ed_id_nv1 As Integer
Public ed_id_nv2 As Integer
Public ed_id_nv3 As Integer
'Public Insumo_nv4 As String
Public ed_id_nv4 As Integer

Public ed_sPrincipais As Boolean

'Public addPrincipal As Boolean
Public ed_sDiversos As Boolean
'Public addDiversos As Boolean
Public ed_sTerceiros As Boolean
Public ed_c_rendimento As Boolean
Public ed_c_pvs As Boolean
Public ed_c_cmo As Boolean
Public ed_btnCmoPvsBoolean As Boolean
Public ed_txtBox_unPublic As String
Public ed_txtBox_custoInsumoPublic As String
Public ed_Pvs As Double
Public ed_Cmo As Double
Public ed_rendimento As Double
Public unit As Boolean
Public coustInput As Boolean
Public existe As Boolean
Public d_un As Boolean
Public d_desc As Boolean
Public d_price As Boolean
Public p_un As Boolean
Public p_desc As Boolean
Public p_price As Boolean
Public strTipo As String
Public counte1 As Long
Public counte As Long







'
Sub atualizarPvsCmo()



GetAdicionais

End Sub

Sub ed_GetPvs_Cmo()


Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command
Dim id_master As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
Dim insumo As String
Dim selectedRow As Integer
Dim coluna2Valor As String



If UserForm2.ed_txtBoxID_nv1.Value = "" Then
UserForm2.ed_txtBoxID_nv1.Value = 0
End If

If UserForm2.ed_txtBoxID_nv2.Value = "" Then
UserForm2.ed_txtBoxID_nv2.Value = 0
End If

ed_id_nv1 = UserForm2.ed_txtBoxID_nv1.Value
ed_id_nv2 = UserForm2.ed_txtBoxID_nv2.Value
If UserForm2.ed_txtBoxID_nv3.Value = "" Then
    UserForm2.ed_txtBoxID_nv3.Value = 0
    ed_id_nv3 = UserForm2.ed_txtBoxID_nv3.Value

Else
ed_id_nv3 = UserForm2.ed_txtBoxID_nv3.Value
End If

sql1 = "t_Servicos_Principais"
sql = "t_Servicos_Principais"

If ed_sDiversos = True Then
'id_master = 7 & 0 & ed_id_nv3
id_master = 7 & "-" & 0 & "-" & ed_id_nv3
ed_id_nv1 = 7
ed_id_nv2 = 0
ed_btnCmoPvsBoolean = True
Else
'id_master = ed_id_nv1 & ed_id_nv2 & ed_id_nv3
id_master = ed_id_nv1 & "-" & ed_id_nv2 & "-" & ed_id_nv3
End If

    i = 0
    While i <= 1
    ConectarBanco conexao
    If ed_sDiversos = True Then
        If i = 0 Then
            sql1 = "SELECT precoVendaSugerido FROM t_Servicos_Diversos WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
        Else
            sql = "SELECT CustoMaoObra FROM t_Servicos_Diversos WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
            
           
        End If
    Else
        If i = 0 Then
            sql1 = "SELECT precoVendaSugerido FROM t_Servicos_Principais WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
        Else
            sql = "SELECT CustoMaoObra FROM t_Servicos_Principais WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
        End If
    End If




If i = 0 Then
    rs.Open sql1, conexao
Else
    rs.Open sql, conexao
End If
    If sql1 <> "" Then
        If i = 0 Then
        On Error GoTo here
        precoVendaSugerido = rs.Fields("precoVendaSugerido").Value
        pvs = precoVendaSugerido
        UserForm2.ed_txtBox_pvs.Value = Format(pvs, "R$ #,##0.00")
        ed_Pvs = Format(pvs, "R$ #,##0.00")
        UserForm2.ed_txtBox_pvs.BackColor = RGB(255, 255, 255)

        UserForm2.ed_txtBox_pvs.Enabled = True
        ed_c_pvs = False
        

        End If
            On Error GoTo here







    Else
        If i = 1 Then
        On Error GoTo -1
        On Error GoTo there
        CustoMaoObra = rs.Fields("CustoMaoObra").Value
        cmo = CustoMaoObra
        ed_c_cmo = False
        UserForm2.ed_txtBox_cmo.Value = Format(cmo, "R$ #,##0.00")
        ed_Cmo = Format(cmo, "R$ #,##0.00")
        UserForm2.ed_txtBox_cmo.BackColor = RGB(255, 255, 255)
        UserForm2.ed_txtBox_cmo.Enabled = True
        End If
        GoTo theEnd
        If i = 0 Then
here:
'IF de teste
            If UserForm2.ed_btnNv3 <> False And UserForm2.ed_s_diversos = True Then
            MsgBox "Não há preço de venda surgerido."
            End If
            c_pvs = True
            UserForm2.ed_txtBox_pvs.BackColor = &H80000016
            UserForm2.ed_txtBox_pvs.Enabled = False
            UserForm2.ed_txtBox_pvs.Value = ""
            
        Else
there:
'IF de teste
            If UserForm2.ed_btnNv3 <> False And UserForm2.ed_s_diversos = True Then
            MsgBox "Não há preço de custo para mão de obra."
            End If
            c_cmo = True
            UserForm2.ed_txtBox_cmo.BackColor = &H80000016
            UserForm2.ed_txtBox_cmo.Enabled = False
            UserForm2.ed_txtBox_cmo.Value = ""
        End If

    End If

    rs.Close
    conexao.Close

    i = i + 1
    
sql1 = ""
sql = ""
Wend
theEnd:
End Sub
'


Sub ed_EnviarDiversos()

Dim selectedRow As Integer
Dim i As Integer
Dim serv_atual As String
Dim input1 As String
Dim input2 As String
Dim ed_new_id1 As Integer
Dim ed_new_id2 As Double
Dim ed_new_id3 As Double
Dim ed_new_id4 As Double
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset





If UserForm2.ed_btnNv3Boolean = True And ed_sDiversos Then


    If UserForm2.ed_txtComboBox_nv3_3.Value = "" Or UserForm2.ed_txtBox_nv3_3.Value = "" Then
    Exit Sub
    End If
    
'AVALIAR SE HOUVE MUDANÇA
input1 = UCase(UserForm2.ed_txtComboBox_nv3_3.Value)
input2 = UCase(UserForm2.ed_txtBox_nv3_3.Value)

        If input1 = input2 Then
        MsgBox "Não existem valores para serem alterados, o valor selecionado no campo de listagem e de digitação são iguais."
        Exit Sub
        End If
 
ed_VerificarServico

If existe = True Then
Exit Sub
End If

        'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
        sql = "t_Nivel3"

        selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
        ConectarBanco conexao
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DIVERSOS';"
        rs.Open sql1, conexao
    
        ed_new_id3 = rs.Fields("idNv3").Value
        rs.Close
        conexao.Close


        'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE
        'Editar Log Insumo
        
        ConectarBanco conexao
                sql = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & ed_new_id3 & " AND grupo = 'SERVICOS DIVERSOS';"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!descricaoNv3 = input2
        rs.Update
        rs.Close
        conexao.Close
        

        Edição.ed_CarregarNivelTres
   
    

'(SERVIÇOS PRINCIPAIS) NÍVEL 4 EDIÇÃO
ElseIf UserForm2.ed_btnNv4Boolean = True And ed_sDiversos Then




        
        If UserForm2.ed_txtBox_un.Value <> ed_txtBox_unPublic And UserForm2.ed_txtBox_un.Value <> "" Then
    
    input1 = ""
    input2 = ""
    input1 = UCase(UserForm2.ed_txtBox_un.Value)
    input2 = UCase(ed_txtBox_unPublic)
    
            If input1 <> input2 Then
            unit = True
    
            'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
            sql = "t_Insumo"
    
            selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
            ConectarBanco conexao
            sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            rs.Open sql1, conexao
        
            ed_new_id4 = rs.Fields("idInsumo").Value
            rs.Close
            conexao.Close
    

            
            ConectarBanco conexao

                    sql = "SELECT Unidade FROM t_Insumos WHERE idInsumo = " & ed_new_id4 & ";"
    
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
 
            rs!unidade = UserForm2.ed_txtBox_un.Value
            rs.Update
            rs.Close
            conexao.Close
            d_un = True

    
            'Edição.ed_CarregarInsumos
            
            MsgBox "Alteração unidade de insumo realizada.", vbInformation
            End If
        End If
        
        If UserForm2.ed_txtBox_custoInsumo.Value <> ed_txtBox_unPublic And UserForm2.ed_txtBox_custoInsumo.Value <> "" Then
        
        SetTipo
    
        UserForm2.ed_txtBox_un.Value = Trim(Replace(UserForm2.ed_txtBox_un.Value, "  ", " "))
    UserForm2.ed_txtBox_rendimento.Value = Trim(Replace(UserForm2.ed_txtBox_rendimento.Value, "  ", " "))
    UserForm2.ed_txtBox_custoInsumo.Value = Trim(Replace(UserForm2.ed_txtBox_custoInsumo.Value, "  ", " "))
    UserForm2.ed_txtBox_pvs.Value = Trim(Replace(UserForm2.ed_txtBox_pvs.Value, "  ", " "))
    UserForm2.ed_txtBox_cmo.Value = Trim(Replace(UserForm2.ed_txtBox_cmo.Value, "  ", " "))


        If Len(Trim(UserForm2.ed_txtBox_un.Value)) <> 0 And Len(Trim(UserForm2.ed_txtBox_custoInsumo.Value)) <> 0 Then
             If UserForm2.ed_txtBox_custoInsumo.Value <> "R$ 0,00" And UserForm2.ed_txtBox_custoInsumo.Value <> "R$ 0,0" Then
                If UserForm2.ed_txtBox_un.Value <> "R$ 0,00" And UserForm2.ed_txtBox_un.Value <> "R$" And UserForm2.ed_txtBox_un.Value <> "R" And UserForm2.ed_txtBox_un.Value <> "$" And UserForm2.ed_txtBox_un.Value <> "," Then
                    If UserForm2.ed_txtBox_custoInsumo.Value <> "R$ 0,00" And UserForm2.ed_txtBox_custoInsumo.Value <> "R$" And UserForm2.ed_txtBox_custoInsumo.Value <> "R" And UserForm2.ed_txtBox_custoInsumo.Value <> "$" And UserForm2.ed_txtBox_custoInsumo.Value <> "," Then
                     GoTo KeepIt:

                    End If
                End If
            End If
        End If
                MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        Exit Sub
KeepIt:
    
    
    
    
            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_custoInsumo.Value)
            input2 = UCase(ed_txtBox_custoInsumoPublic)

    
            If input1 <> input2 Then
            coustInput = True
    

            sql = "t_Insumo"
    
            selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
            ConectarBanco conexao
            sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            rs.Open sql1, conexao
        
            ed_new_id4 = rs.Fields("idInsumo").Value
            rs.Close
            conexao.Close
    
    
            'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE
    
            ConectarBanco conexao

                    sql = "SELECT Custo FROM t_Insumos WHERE idInsumo = " & ed_new_id4 & ";"
    
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!CUSTO = input1
            rs.Update
            rs.Close
            conexao.Close
            d_price = True
    
            'Edição.ed_CarregarInsumos
            
            MsgBox "Alteração em custo de insumo realizada.", vbInformation
            End If
            
        End If
        
        If unit = True Or coustInput = True Then
        'Edição.ed_CarregarInsumos
        Else

        End If
        
        unit = False
        coustInput = False
        
        
        If UserForm2.ed_ComboBoxTipo = "PRINCIPAL" Then
        varToCompare = "SERVICOS PRINCIPAIS"
        ElseIf UserForm2.ed_ComboBoxTipo = "DIVERSO" Then
        varToCompare = "SERVICOS DIVERSOS"
        ElseIf UserForm2.ed_ComboBoxTipo = "GENÉRICO" Then
        varToCompare = "GENERICO"
        End If
        
        If strTipo <> varToCompare Then
        UserForm2.ed_btnNv4_carregar
        End If
        
        
        

            If UserForm2.ed_txtComboBox_nv4_4.Value = "" Or UserForm2.ed_txtBox_nv4_4.Value = "" And UserForm2.ed_txtBox_un = "" And UserForm2.ed_txtBox_custoInsumo = "" Or UserForm2.ed_txtBox_custoInsumo.Value = "R$ 0,00" Then
    
MsgBox "Não será possível realizar a atualização!" & vbCrLf & "" & vbCrLf & "Apenas valores maiores que zero serão aceitos." & vbCrLf & "Todos os campos de listagem devem estar selecionados.", vbExclamation
    Exit Sub
    End If
    
'AVALIAR SE HOUVE MUDANÇA
input1 = UCase(UserForm2.ed_txtComboBox_nv4_4.Value)
input2 = UCase(UserForm2.ed_txtBox_nv4_4.Value)

    If input1 <> input2 Then
    
ed_VerificarServico

If existe = True Then
Exit Sub
End If


'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
        sql = "t_Insumo"

        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
    
        ed_new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close


'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE

            
        ConectarBanco conexao

                sql = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & ed_new_id4 & ";"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!insumo = input2
        rs.Update
        rs.Close
        conexao.Close
        d_desc = True

 '       Edição.ed_CarregarInsumos
        
        MsgBox "Alteração em descrição de insumo realizada.", vbInformation
        End If
        If d_desc = True Or d_price = True Or d_un = True Then
        'Editar Log Insumo
          log_edicao_insumo
          Edição.ed_CarregarInsumos
          d_desc = False
          d_price = False
          d_un = False
        Else
                        MsgBox "Não será possível realizar a atualização dos campos Insumo, Unidade e Custo Insumo!" & vbCrLf & "" & vbCrLf & "Apenas valores maiores que zero serão aceitos." & vbCrLf & "Todos os campos de listagem devem estar selecionados.", vbExclamation
        End If



ElseIf UserForm2.ed_btnCmoPvsBoolean = True And ed_sDiversos Then




        If UserForm2.ed_txtComboBox_nv3_3 <> "" And UserForm2.ed_txtBox_cmo <> "" And UserForm2.ed_txtBox_cmo <> "R$ 0,00" Then
        
        
    UserForm2.ed_txtBox_un.Value = Trim(Replace(UserForm2.ed_txtBox_un.Value, "  ", " "))
    UserForm2.ed_txtBox_rendimento.Value = Trim(Replace(UserForm2.ed_txtBox_rendimento.Value, "  ", " "))
    UserForm2.ed_txtBox_custoInsumo.Value = Trim(Replace(UserForm2.ed_txtBox_custoInsumo.Value, "  ", " "))
    UserForm2.ed_txtBox_pvs.Value = Trim(Replace(UserForm2.ed_txtBox_pvs.Value, "  ", " "))
    UserForm2.ed_txtBox_cmo.Value = Trim(Replace(UserForm2.ed_txtBox_cmo.Value, "  ", " "))

    
        If Len(Trim(UserForm2.ed_txtBox_cmo.Value)) <> 0 Then
             If UserForm2.ed_txtBox_cmo.Value <> "R$ 0,00" Or UserForm2.ed_txtBox_cmo.Value <> "R$ 0,0" Then
                If UserForm2.ed_txtBox_cmo.Value <> "R$ 0,00" And UserForm2.ed_txtBox_cmo.Value <> "R$" And UserForm2.ed_txtBox_cmo.Value <> "R" And UserForm2.ed_txtBox_cmo.Value <> "$" And UserForm2.ed_txtBox_cmo.Value <> "," Then
                                GoTo KeepItCmo
                End If
             End If
         End If
        


                MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        Exit Sub
KeepItCmo:
        
            i = 0
''    PVS -> CHECAR SE PRECISO ATUALIZAR

        
        
        '    CMO -> CHECAR SE PRECISO ATUALIZAR

            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_cmo.Value)
            input2 = UCase(ed_Cmo)
            input2 = Format(input2, "R$ #,##0.00")
            input1 = Format(input1, "R$ #,##0.00")

            If input1 <> input2 Then
            i = i + 1

            
            ConectarBanco conexao

            'sql = "SELECT precoVendaSugerido FROM t_Servicos_Principais WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
            sql = "SELECT CustoMaoObra FROM t_Servicos_Diversos WHERE idNv1 = 7 AND idNv2 = 0 AND idNv3 = " & ed_id_nv3
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic


            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!CustoMaoObra = input1
            rs.Update
            rs.Close
            conexao.Close




            MsgBox "Alteração em custo de mão de obra realizada.", vbInformation


            End If
            
            If i <> 0 Then
            'Criar Log cmo
            log_edicao_cmo_pvs
            Edição.ed_CarregarNivelUm
            Edição.ed_CarregarNivelDois
            Edição.ed_CarregarNivelTres
            
            UserForm2.ed_txtBox_cmo.Value = ""
            UserForm2.ed_txtBox_cmo.Enabled = False
            UserForm2.ed_txtBox_cmo.BackColor = &H80000016
            UserForm2.ed_txtBox_pvs.Value = ""
            UserForm2.ed_txtBox_pvs.Enabled = False
            UserForm2.ed_txtBox_pvs.BackColor = &H80000016
            
            Else
            MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero.", vbExclamation
            End If
            

            
        Else
        MsgBox "Não será possível realizar a atualização!" & vbCrLf & "" & vbCrLf & "Apenas valores maiores que zero serão aceitos." & vbCrLf & "Todos os campos de listagem devem estar selecionados.", vbExclamation
        Exit Sub
        End If
        
        
    

ElseIf UserForm2.ed_btnRendimentoBoolean = True And ed_sDiversos Then

        If UserForm2.ed_txtComboBox_nv3_3 <> "" And UserForm2.ed_txtComboBox_nv4_4 <> "" And UserForm2.ed_txtBox_rendimento <> "" Then
            
    UserForm2.ed_txtBox_un.Value = Trim(Replace(UserForm2.ed_txtBox_un.Value, "  ", " "))
    UserForm2.ed_txtBox_rendimento.Value = Trim(Replace(UserForm2.ed_txtBox_rendimento.Value, "  ", " "))
    UserForm2.ed_txtBox_custoInsumo.Value = Trim(Replace(UserForm2.ed_txtBox_custoInsumo.Value, "  ", " "))
    UserForm2.ed_txtBox_pvs.Value = Trim(Replace(UserForm2.ed_txtBox_pvs.Value, "  ", " "))
    UserForm2.ed_txtBox_cmo.Value = Trim(Replace(UserForm2.ed_txtBox_cmo.Value, "  ", " "))

    
        If Len(Trim(UserForm2.ed_txtBox_rendimento.Value)) <> 0 Then
             If UserForm2.ed_txtBox_rendimento.Value <> "R$ 0,00" Or UserForm2.ed_txtBox_rendimento.Value <> "R$ 0,0" Then
                If UserForm2.ed_txtBox_rendimento.Value <> "R$ 0,00" And UserForm2.ed_txtBox_rendimento.Value <> "R$" And UserForm2.ed_txtBox_rendimento.Value <> "R" And UserForm2.ed_txtBox_rendimento.Value <> "$" And UserForm2.ed_txtBox_rendimento.Value <> "," Then
                                GoTo KeepItRendimento
                End If
             End If
         End If
        


        MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        Exit Sub
KeepItRendimento:

            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_rendimento.Value)
            input2 = UCase(ed_rendimento)
            input2 = Format(input2, "#,##0.00")
            
            If input1 <> input2 Then

            
            ConectarBanco conexao
            sql = "SELECT rendimento FROM t_Servicos_Diversos_Rendimento WHERE idNv1 = 7 AND idNv2 = 0 AND idNv3 = " & ed_id_nv3 & " AND idInsumo = " & ed_id_nv4
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
            
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!rendimento = input1
            rs.Update
            rs.Close
            conexao.Close
            
            'Editar Log Rendimento
            log_edicao_rendimento
            
            MsgBox "Alteração de rendimento realizada.", vbInformation
            
            Edição.ed_CarregarNivelUm
            Edição.ed_CarregarNivelDois
            Edição.ed_CarregarNivelTres
            Edição.ed_CarregarInsumos
            UserForm2.ed_txtBox_rendimento.Value = ""
            UserForm2.ed_txtBox_rendimento.Enabled = False
            UserForm2.ed_txtBox_rendimento.BackColor = &H80000016
            
            Exit Sub
            End If
            
            MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero.", vbExclamation
       
     
        


            
        Else
        MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero.", vbExclamation
       
        End If
        








End If

End Sub

Sub ed_EnviarTerceiros()
Dim selectedRow As Integer
Dim i As Integer
Dim serv_atual As String
Dim input1 As String
Dim input2 As String
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

If UserForm2.ed_btnNv3Boolean = True And ed_sTerceiros Then


    If UserForm2.ed_txtComboBox_nv3_3.Value = "" Or UserForm2.ed_txtBox_nv3_3.Value = "" Then
    Exit Sub
    End If
    
        'AVALIAR SE HOUVE MUDANÇA
        input1 = UCase(UserForm2.ed_txtComboBox_nv3_3.Value)
        input2 = UCase(UserForm2.ed_txtBox_nv3_3.Value)

        If input1 = input2 Then
        MsgBox "Não existem valores para serem alterados, o valor selecionado no campo de listagem e de digitação são iguais.", vbExclamation
        Exit Sub
        End If
        
        
        
        'VERIFICA SE O VALOR A SER EDITADO JÁ EXISTE
        ed_VerificarServico
        
        If existe = True Then
        Exit Sub
        End If

        'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
        sql = "t_Nivel3"

        selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
        ConectarBanco conexao
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DE TERCEIROS';"
        rs.Open sql1, conexao
    
        ed_new_id3 = rs.Fields("idNv3").Value
        rs.Close
        conexao.Close


        log_edicao_nivel
        
        ConectarBanco conexao
                sql = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & ed_new_id3 & " AND grupo = 'SERVICOS DE TERCEIROS';"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!descricaoNv3 = input2
        rs.Update
        rs.Close
        conexao.Close
        

        Edição.ed_CarregarNivelTres
End If


End Sub



Sub ed_VerificarServico()

Dim selectedRow As Integer
Dim i As Integer
Dim serv_atual As String
Dim input1 As String
Dim input2 As String
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

If ed_sPrincipais = True Then
    
    'VERIFY NV 1
    If UserForm2.ed_btnNv1Boolean = True Then


        ConectarBanco conexao
'        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & UserForm2.ed_txtBox_nv1_1.Value & "'"
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & UserForm2.ed_txtBox_nv1_1.Value & "' AND grupo = 'SERVICOS PRINCIPAIS'"

        rs.Open sql1, conexao
    
        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close
        
    End If
    'VERIFY NV 2
    If UserForm2.ed_btnNv2Boolean = True Then
    
        ConectarBanco conexao
        sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & UserForm2.ed_txtBox_nv2_2.Value & "' AND grupo = 'SERVICOS PRINCIPAIS'"
        rs.Open sql1, conexao
    
        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close
  
    End If
    'VERIFY NV 3
    If UserForm2.ed_btnNv3Boolean = True Then
    
        ConectarBanco conexao
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS PRINCIPAIS'"
        rs.Open sql1, conexao
    
        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close

        
    End If
    'VERIFY NV 4
    If UserForm2.ed_btnNv4Boolean = True Then
    

        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtBox_nv4_4.Value & "'"
        rs.Open sql1, conexao

        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close
    
    End If


ElseIf ed_sDiversos = True Then

    If UserForm2.ed_btnNv3Boolean = True Then
    
        ConectarBanco conexao
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS DIVERSOS'"
        rs.Open sql1, conexao
    
        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close

        
    End If
    'VERIFY NV 4
    If UserForm2.ed_btnNv4Boolean = True Then
    

        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtBox_nv4_4.Value & "'"
        rs.Open sql1, conexao

        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close
    
    End If

ElseIf ed_sTerceiros = True Then

    If UserForm2.ed_btnNv3Boolean = True Then
    
        ConectarBanco conexao
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtBox_nv3_3.Value & "' AND grupo = 'SERVICOS DE TERCEIROS'"
        rs.Open sql1, conexao
    
        If rs.EOF Then
            existe = False
            Exit Sub
        Else
        MsgBox "Serviço já existente na base de dados."
            existe = True
            Exit Sub
        End If
        
        rs.Close
        conexao.Close

        
    End If

End If


End Sub



Sub ed_Enviar()
Dim selectedRow As Integer
Dim i As Integer
Dim serv_atual As String
Dim input1 As String
Dim input2 As String
Dim ed_new_id1 As Integer
Dim ed_new_id2 As Double
Dim ed_new_id3 As Double
Dim ed_new_id4 As Double
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

If ed_sDiversos = True Then
ed_EnviarDiversos

Exit Sub
ElseIf ed_sTerceiros = True Then
ed_EnviarTerceiros
Exit Sub
End If


If UserForm2.ed_btnNv1Boolean = True And ed_sPrincipais Then
    If UserForm2.ed_txtComboBox_nv1_1.Value = "" Or UserForm2.ed_txtBox_nv1_1.Value = "" Then
    MsgBox "Não existem valores para serem alterados ou valor selecionado em lista igual ao valor digitado.", vbExclamation
    Exit Sub
    End If
'AVALIAR SE HOUVE MUDANÇA
input1 = UCase(UserForm2.ed_txtComboBox_nv1_1.Value)
input2 = UCase(UserForm2.ed_txtBox_nv1_1.Value)

    If input1 = input2 Then
    MsgBox "Não existem valores para serem alterados ou valor selecionado em lista igual ao valor digitado.", vbExclamation
    Exit Sub
    End If
    
'VALIDAR SE DESCRIÇÃO DE NÍVEL JÁ EXISTE DENTRO DESTE NÍVEL


ed_VerificarServico

If existe = True Then
Exit Sub
End If



'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
        sql = "t_Nivel1"

        selectedRow = UserForm2.ed_txtComboBox_nv1_1.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & UserForm2.ed_txtComboBox_nv1_1.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
    
        ed_new_id1 = rs.Fields("idNv1").Value
        rs.Close
        conexao.Close


'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE

        'Editar Log Nivel
        log_edicao_nivel

        ConectarBanco conexao

                sql = "SELECT descricaoNv1 FROM t_Nivel1 WHERE idNv1 = " & ed_new_id1 & " AND grupo = 'SERVICOS PRINCIPAIS';"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!descricaoNv1 = input2
        rs.Update
        rs.Close
        conexao.Close
        

        Edição.ed_CarregarNivelUm

ElseIf UserForm2.ed_btnNv2Boolean = True And ed_sPrincipais Then

    If UserForm2.ed_txtComboBox_nv2_2.Value = "" Or UserForm2.ed_txtBox_nv2_2.Value = "" Then
    MsgBox "Não existem valores para serem alterados ou valor selecionado em lista igual ao valor digitado.", vbExclamation
    Exit Sub
    End If
'AVALIAR SE HOUVE MUDANÇA
input1 = UCase(UserForm2.ed_txtComboBox_nv2_2.Value)
input2 = UCase(UserForm2.ed_txtBox_nv2_2.Value)

    If input1 = input2 Then
    MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero.", vbExclamation
    Exit Sub
    End If
    
'VALIDAR SE DESCRIÇÃO DE NÍVEL JÁ EXISTE DENTRO DESTE NÍVEL


ed_VerificarServico

If existe = True Then
Exit Sub
End If

'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE

        sql = "t_Nivel2"

        selectedRow = UserForm2.ed_txtComboBox_nv2_2.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & UserForm2.ed_txtComboBox_nv2_2.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
    
        ed_new_id2 = rs.Fields("idNv2").Value
        rs.Close
        conexao.Close


'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE
        'Editar Log Nivel
        log_edicao_nivel

        ConectarBanco conexao

                sql = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & ed_new_id2 & " AND grupo = 'SERVICOS PRINCIPAIS';"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!descricaoNv2 = input2
        rs.Update
        rs.Close
        conexao.Close
        

        Edição.ed_CarregarNivelDois



ElseIf UserForm2.ed_btnNv3Boolean = True And ed_sPrincipais Then


    If UserForm2.ed_txtComboBox_nv3_3.Value = "" Or UserForm2.ed_txtBox_nv3_3.Value = "" Then
    MsgBox "Não existem valores para serem alterados ou valor selecionado em lista igual ao valor digitado.", vbExclamation
    Exit Sub
    End If
    
'AVALIAR SE HOUVE MUDANÇA
input1 = UCase(UserForm2.ed_txtComboBox_nv3_3.Value)
input2 = UCase(UserForm2.ed_txtBox_nv3_3.Value)

        If input1 = input2 Then
        MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero.", vbExclamation
        Exit Sub
        End If
        
'VALIDAR SE DESCRIÇÃO DE NÍVEL JÁ EXISTE DENTRO DESTE NÍVEL


ed_VerificarServico

If existe = True Then
Exit Sub
End If
 


        'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
        sql = "t_Nivel2"

        selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
    
        ed_new_id3 = rs.Fields("idNv3").Value
        rs.Close
        conexao.Close


        'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE
        
        'Editar Log Nivel
        log_edicao_nivel

        ConectarBanco conexao
                sql = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & ed_new_id3 & " AND grupo = 'SERVICOS PRINCIPAIS';"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!descricaoNv3 = input2
        rs.Update
        rs.Close
        conexao.Close
        

        Edição.ed_CarregarNivelTres
   
    

'(SERVIÇOS PRINCIPAIS) NÍVEL 4 EDIÇÃO
ElseIf UserForm2.ed_btnNv4Boolean = True And ed_sPrincipais = True Then
'Verifica e Edita Tipo do Insumo no DB
    SetTipo
'    MsgBox "Não será possível realizar a atualização!" & vbCrLf & "" & vbCrLf & "Apenas valores maiores que zero serão aceitos." & vbCrLf & "Todos os campos de listagem devem estar selecionados.", vbExclamation
'    Exit Sub
    


    UserForm2.ed_txtBox_un.Value = Trim(Replace(UserForm2.ed_txtBox_un.Value, "  ", " "))
    UserForm2.ed_txtBox_rendimento.Value = Trim(Replace(UserForm2.ed_txtBox_rendimento.Value, "  ", " "))
    UserForm2.ed_txtBox_custoInsumo.Value = Trim(Replace(UserForm2.ed_txtBox_custoInsumo.Value, "  ", " "))
    UserForm2.ed_txtBox_pvs.Value = Trim(Replace(UserForm2.ed_txtBox_pvs.Value, "  ", " "))
    UserForm2.ed_txtBox_cmo.Value = Trim(Replace(UserForm2.ed_txtBox_cmo.Value, "  ", " "))
    
'
        If Len(Trim(UserForm2.ed_txtBox_nv4_4.Value)) <> "" Then
             If UserForm2.ed_txtBox_un.Value <> "" Then
                If UserForm2.ed_txtBox_custoInsumo.Value <> "R$ 0,00" And UserForm2.ed_txtBox_custoInsumo.Value <> "R$" And UserForm2.ed_txtBox_rendimento.Value <> "R" And UserForm2.ed_txtBox_custoInsumo.Value <> "$" And UserForm2.ed_txtBox_custoInsumo.Value <> "," Then
                                GoTo KeepIt:
                End If
             End If
         End If
'
'

                MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        Exit Sub
KeepIt:
    If UserForm2.ed_txtComboBox_nv4_4.Value = "" Or UserForm2.ed_txtBox_nv4_4.Value <> "" And UserForm2.ed_txtBox_un <> "" And UserForm2.ed_txtBox_custoInsumo.Value <> "" Then

    

        If UserForm2.ed_txtBox_un.Value <> ed_txtBox_unPublic And UserForm2.ed_txtBox_un.Value <> "" Then
    
    input1 = ""
    input2 = ""
    input1 = UCase(UserForm2.ed_txtBox_un.Value)
    input2 = UCase(ed_txtBox_unPublic)
    
            If input1 <> input2 Then
            unit = True
    
            'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
            sql = "t_Insumo"
    
            selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
            ConectarBanco conexao
            sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            rs.Open sql1, conexao
        
            ed_new_id4 = rs.Fields("idInsumo").Value
            rs.Close
            conexao.Close
    
    

         

    
            ConectarBanco conexao

                    sql = "SELECT Unidade FROM t_Insumos WHERE idInsumo = " & ed_new_id4 & ";"
    
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!unidade = UserForm2.ed_txtBox_un.Value
            rs.Update
            rs.Close
            conexao.Close
            d_un = True
    
            'Edição.ed_CarregarInsumos
            
            MsgBox "Alteração unidade de insumo realizada.", vbInformation
            End If
        End If
        
        If UserForm2.ed_txtBox_custoInsumo.Value <> ed_txtBox_unPublic And UserForm2.ed_txtBox_custoInsumo.Value <> "" Then
    
            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_custoInsumo.Value)
            input2 = UCase(ed_txtBox_custoInsumoPublic)
            input2 = Format(input2, "R$ #,##0.00")
            input1 = Format(input1, "R$ #,##0.00")
            If input1 <> input2 Then
            coustInput = True
    
            'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
            sql = "t_Insumo"
    
            selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
            ConectarBanco conexao
            sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            rs.Open sql1, conexao
        
            ed_new_id4 = rs.Fields("idInsumo").Value
            rs.Close
            conexao.Close
    
    

            
            ConectarBanco conexao

                    sql = "SELECT Custo FROM t_Insumos WHERE idInsumo = " & ed_new_id4 & ";"
    
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!CUSTO = input1
            rs.Update
            rs.Close
            conexao.Close
            d_price = True
    
            'Edição.ed_CarregarInsumos
            
            MsgBox "Alteração em custo de insumo realizada.", vbInformation
            End If
            
        If unit = True Or coustInput = True Then
        'Edição.ed_CarregarInsumos
        Else
                'MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        End If
        
        unit = False
        coustInput = False
            
        End If
        
        'AVALIAR SE HOUVE MUDANÇA
        input1 = UCase(UserForm2.ed_txtComboBox_nv4_4.Value)
        input2 = UCase(UserForm2.ed_txtBox_nv4_4.Value)
        
            If input1 <> input2 Then


'VALIDAR SE DESCRIÇÃO DE NÍVEL JÁ EXISTE DENTRO DESTE NÍVEL

        
        ed_VerificarServico
        
        If existe = True Then
        Exit Sub
        End If

'ENCONTRAR ID DA DESCRIÇÃO NV1 EXISTENTE
        sql = "t_Insumo"

        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
    
        ed_new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close


'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE


        
        
        ConectarBanco conexao

                sql = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & ed_new_id4 & ";"

        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
 

        'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
        rs!insumo = input2
        rs.Update
        rs.Close
        conexao.Close
        d_desc = True

        'Edição.ed_CarregarInsumos
        
        MsgBox "Alteração em descrição de insumo realizada.", vbInformation
        End If
        
        If d_desc = True Or d_price = True Or d_un = True Then
        'Editar Log Insumo
          log_edicao_insumo
          Edição.ed_CarregarInsumos
          d_desc = False
          d_price = False
          d_un = False
        End If
        
        If UserForm2.ed_ComboBoxTipo = "PRINCIPAL" Then
        varToCompare = "SERVICOS PRINCIPAIS"
        ElseIf UserForm2.ed_ComboBoxTipo = "DIVERSO" Then
        varToCompare = "SERVICOS DIVERSOS"
        ElseIf UserForm2.ed_ComboBoxTipo = "GENÉRICO" Then
        varToCompare = "GENERICO"
        End If
        
        If strTipo <> varToCompare Then
        UserForm2.ed_btnNv4_carregar
        End If
        
        'SetTipo
End If
        

ElseIf UserForm2.ed_btnCmoPvsBoolean = True And ed_sPrincipais Then


        If UserForm2.ed_txtComboBox_nv1_1 <> "" And UserForm2.ed_txtComboBox_nv2_2 <> "" And UserForm2.ed_txtComboBox_nv3_3 <> "" And UserForm2.ed_txtBox_cmo <> "" And UserForm2.ed_txtBox_cmo <> "R$ 0,00" And UserForm2.ed_txtBox_pvs <> "" And UserForm2.ed_txtBox_pvs <> "R$ 0,00" Then
    
    UserForm2.ed_txtBox_un.Value = Trim(Replace(UserForm2.ed_txtBox_un.Value, "  ", " "))
    UserForm2.ed_txtBox_rendimento.Value = Trim(Replace(UserForm2.ed_txtBox_rendimento.Value, "  ", " "))
    UserForm2.ed_txtBox_custoInsumo.Value = Trim(Replace(UserForm2.ed_txtBox_custoInsumo.Value, "  ", " "))
    UserForm2.ed_txtBox_pvs.Value = Trim(Replace(UserForm2.ed_txtBox_pvs.Value, "  ", " "))
    UserForm2.ed_txtBox_cmo.Value = Trim(Replace(UserForm2.ed_txtBox_cmo.Value, "  ", " "))

    
        If Len(Trim(UserForm2.ed_txtBox_pvs.Value)) <> 0 And Len(Trim(UserForm2.ed_txtBox_cmo.Value)) <> 0 Then
             If UserForm2.ed_txtBox_pvs.Value <> "R$ 0,00" Or UserForm2.ed_txtBox_pvs.Value <> "R$ 0,0" And UserForm2.ed_txtBox_cmo.Value <> "R$ 0,00" Or UserForm2.ed_txtBox_cmo.Value <> "R$ 0,0" Then
                            If UserForm2.ed_txtBox_pvs.Value <> "R$ 0,00" And UserForm2.ed_txtBox_pvs.Value <> "R$" And UserForm2.ed_txtBox_pvs.Value <> "R" And UserForm2.ed_txtBox_pvs.Value <> "$" And UserForm2.ed_txtBox_pvs.Value <> "," Then
                                If UserForm2.ed_txtBox_cmo.Value <> "R$ 0,00" And UserForm2.ed_txtBox_cmo.Value <> "R$" And UserForm2.ed_txtBox_cmo.Value <> "R" And UserForm2.ed_txtBox_cmo.Value <> "$" And UserForm2.ed_txtBox_cmo.Value <> "," Then
                                GoTo KeepItCmoPsv
                                End If
                            End If
                        End If
                    End If


                MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        Exit Sub
KeepItCmoPsv:
            
            
            i = 0
'    PVS -> CHECAR SE PRECISO ATUALIZAR

            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_pvs.Value)
            input2 = UCase(ed_Pvs)
            input2 = Format(input2, "R$ #,##0.00")
                        
            If input1 <> input2 Then
            i = i + 1
            
            'Editar Log PVS
            
            ConectarBanco conexao
            sql = "SELECT precoVendaSugerido FROM t_Servicos_Principais WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!precoVendaSugerido = input1
            rs.Update
            rs.Close
            conexao.Close
            
    
            'Edição.ed_CarregarInsumos
            
            MsgBox "Alteração preço de venda sugerido realizada.", vbInformation

            
            End If
       
     
        
        
        
        '    CMO -> CHECAR SE PRECISO ATUALIZAR

            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_cmo.Value)
            input2 = UCase(ed_Cmo)
            input2 = Format(input2, "R$ #,##0.00")
            
            If input1 <> input2 Then
            i = i + 1

            'SUBSTITUIR DESCRIÇÃO USANDO O ID EXISTENTE
            
            'Editar Log CMO
            
            ConectarBanco conexao

            'sql = "SELECT precoVendaSugerido FROM t_Servicos_Principais WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
            sql = "SELECT CustoMaoObra FROM t_Servicos_Principais WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!CustoMaoObra = input1
            rs.Update
            rs.Close
            conexao.Close
            
    

            
            MsgBox "Alteração em custo de mão de obra realizada.", vbInformation
            

            End If
            
            If i <> 0 Then
            
            log_edicao_cmo_pvs
            
            Edição.ed_CarregarNivelUm
            Edição.ed_CarregarNivelDois
            Edição.ed_CarregarNivelTres
            
            UserForm2.ed_txtBox_cmo.Value = ""
            UserForm2.ed_txtBox_cmo.Enabled = False
            UserForm2.ed_txtBox_cmo.BackColor = &H80000016
            UserForm2.ed_txtBox_pvs.Value = ""
            UserForm2.ed_txtBox_pvs.Enabled = False
            UserForm2.ed_txtBox_pvs.BackColor = &H80000016
            
            Else
            MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero.", vbExclamation
            End If
            

            
        Else
        MsgBox "Não será possível realizar a atualização!" & vbCrLf & "" & vbCrLf & "Apenas valores maiores que zero serão aceitos." & vbCrLf & "Todos os campos de listagem devem estar selecionados.", vbExclamation
        Exit Sub
        End If
        
    

ElseIf UserForm2.ed_btnRendimentoBoolean = True And ed_sPrincipais Then

        If UserForm2.ed_txtComboBox_nv1_1 <> "" And UserForm2.ed_txtComboBox_nv2_2 <> "" And UserForm2.ed_txtComboBox_nv3_3 <> "" And UserForm2.ed_txtComboBox_nv4_4 <> "" And UserForm2.ed_txtBox_rendimento <> "" And UserForm2.ed_txtComboBox_nv4_4 <> "R$ 0,00" And UserForm2.ed_txtBox_rendimento <> "R$ 0,00" Then
            

    UserForm2.ed_txtBox_un.Value = Trim(Replace(UserForm2.ed_txtBox_un.Value, "  ", " "))
    UserForm2.ed_txtBox_rendimento.Value = Trim(Replace(UserForm2.ed_txtBox_rendimento.Value, "  ", " "))
    UserForm2.ed_txtBox_custoInsumo.Value = Trim(Replace(UserForm2.ed_txtBox_custoInsumo.Value, "  ", " "))
    UserForm2.ed_txtBox_pvs.Value = Trim(Replace(UserForm2.ed_txtBox_pvs.Value, "  ", " "))
    UserForm2.ed_txtBox_cmo.Value = Trim(Replace(UserForm2.ed_txtBox_cmo.Value, "  ", " "))

    
        If Len(Trim(UserForm2.ed_txtBox_rendimento.Value)) <> 0 Then
             If UserForm2.ed_txtBox_rendimento.Value <> "R$ 0,00" Or UserForm2.ed_txtBox_rendimento.Value <> "R$ 0,0" Then
                If UserForm2.ed_txtBox_rendimento.Value <> "R$ 0,00" And UserForm2.ed_txtBox_rendimento.Value <> "R$" And UserForm2.ed_txtBox_rendimento.Value <> "R" And UserForm2.ed_txtBox_rendimento.Value <> "$" And UserForm2.ed_txtBox_rendimento.Value <> "," Then
                                GoTo KeepItRendimento
                End If
             End If
         End If
        


        MsgBox "Não existem valores para serem alterados ou o valor presente é igual a zero, R, R$ ou $.", vbExclamation
        Exit Sub
KeepItRendimento:

'    RENDIMENTO -> CHECAR SE PRECISO ATUALIZAR

            input1 = ""
            input2 = ""
            input1 = UCase(UserForm2.ed_txtBox_rendimento.Value)
            input2 = UCase(ed_rendimento)
            input2 = Format(input2, "#,##0.00")
           
            
            If input1 <> input2 Then
        
            'Editar Log Rendimento
            log_edicao_rendimento
            
            ConectarBanco conexao
            sql = "SELECT rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & ed_id_nv1 & " AND idNv2 = " & ed_id_nv2 & " AND idNv3 = " & ed_id_nv3 & " AND idInsumo = " & ed_id_nv4
            rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
     
    
            'MsgBox "Apenas " & rs.RecordCount & " item encontrado para substituição."
            rs!rendimento = input1
            rs.Update
            rs.Close
            conexao.Close
            
            MsgBox "Alteração de rendimento realizada.", vbInformation
            
            Edição.ed_CarregarNivelUm
            Edição.ed_CarregarNivelDois
            Edição.ed_CarregarNivelTres
            Edição.ed_CarregarInsumos
            UserForm2.ed_txtBox_rendimento.Value = ""
            UserForm2.ed_txtBox_rendimento.Enabled = False
            UserForm2.ed_txtBox_rendimento.BackColor = &H80000016
            
            Exit Sub
            End If
            
            MsgBox "Não existem valores para serem alterados.", vbExclamation
       
     
        


            
        Else
        MsgBox "Não será possível realizar a atualização!" & vbCrLf & "" & vbCrLf & "Apenas valores maiores que zero serão aceitos." & vbCrLf & "Todos os campos de listagem devem estar selecionados.", vbExclamation
        Exit Sub
        End If
        
        
        
        If nv1 = False Then
        'montar frase de: "Nenhuma modificação realizada para"
        ElseIf nv2 = False Then
        ElseIf Nv3 = False Then
        ElseIf Nv3 = False Then
        ElseIf Nv4 = False Then
        '[...]
        End If
        







End If







End Sub


Sub ed_GetIdNv1()


Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ed_id_nv1 = 0
ed_id_nv2 = 0
ed_id_nv3 = 0

ConectarBanco conexao

sql = "t_Nivel1"
sql1 = "t_Nivel1"


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = UserForm2.ed_txtComboBox_nv1_1.ListIndex
Dim selectedRow1 As String
selectedRow1 = UserForm2.ed_txtBox_nv1_1.Value


   Dim valor As String
   'AREAS TERREAS
    If Not IsNull(UserForm2.ed_txtComboBox_nv1_1.Value) And Len(UserForm2.ed_txtComboBox_nv1_1.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv1_1.Value
        selectedRow = UserForm2.ed_txtComboBox_nv1_1.ListIndex
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & UserForm2.ed_txtComboBox_nv1_1.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(UserForm2.ed_txtBox_nv1_1.Value) And Len(UserForm2.ed_txtBox_nv1_1.Value) > 0 Then
   On Error GoTo here
   'On Error GoTo -1

        selectedRow1 = UserForm2.ed_txtBox_nv1_1.Value
        'sql1 = "SELECT * FROM t_Nivel1 WHERE idNv1 = '" & selectedRow1 & "';"
        sql1 = "SELECT idNv1 FROM t_Nivel1 WHERE descricaoNv1 = '" & selectedRow1 & "';"
        'selectedRow = UserForm2.ed_txtBox_nv1_1.Value
        valor = UserForm2.ed_txtBox_nv1_1.Value


   End If

'===============================================================


'====
If selectedRow = -1 Then
idNv1 = 0
UserForm2.ed_txtBoxID_nv1.Value = idNv1


If (selectedRow = -1 And UserForm2.ed_txtBox_nv1_1.Value <> "Digite o serviço aqui") And (selectedRow = -1 And UserForm2.ed_txtBox_nv1_1.Value <> "") Then

rs.Open sql1, conexao

idNv1 = rs.Fields("idNv1").Value
UserForm2.ed_txtBoxID_nv1.Value = idNv1

ed_id_nv1 = idNv1
rs.Close
conexao.Close

End If

Else


rs.Open sql1, conexao
'grupo = rs.Fields("grupo").Value
idNv1 = rs.Fields("idNv1").Value

UserForm2.ed_txtBoxID_nv1.Value = idNv1
ed_id_nv1 = idNv1
rs.Close
conexao.Close



End If

    GoTo jumpit
here: Exit Sub
jumpit:

End Sub
'
'
'
Sub ed_GetIdNv2()

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
selectedRow = UserForm2.ed_txtComboBox_nv2_2.ListIndex
Dim selectedRow1 As String
selectedRow1 = UserForm2.ed_txtBox_nv2_2.Value



   Dim valor As String
    If Not IsNull(UserForm2.ed_txtComboBox_nv2_2.Value) And Len(UserForm2.ed_txtComboBox_nv2_2.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv2_2.Value
        selectedRow = UserForm2.ed_txtComboBox_nv2_2.ListIndex
        sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & UserForm2.ed_txtComboBox_nv2_2.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(UserForm2.ed_txtBox_nv2_2.Value) And Len(UserForm2.ed_txtBox_nv2_2.Value) > 0 Then


            
            On Error GoTo here
            selectedRow1 = UserForm2.ed_txtBox_nv2_2.Value
            sql1 = "SELECT idNv2 FROM t_Nivel2 WHERE descricaoNv2 = '" & selectedRow1 & "';"
            

   End If


            If selectedRow = -1 Then
            idNv2 = 0
            UserForm2.ed_txtBoxID_nv2.Value = idNv2


            If (selectedRow = -1 And UserForm2.ed_txtBox_nv2_2.Value <> "Digite o serviço aqui") And (selectedRow = -1 And UserForm2.ed_txtBox_nv2_2.Value <> "") Then

            rs.Open sql1, conexao

            idNv2 = rs.Fields("idNv2").Value
            UserForm2.ed_txtBoxID_nv2.Value = idNv2
            ed_id_nv2 = idNv2
            rs.Close
            conexao.Close

            End If

            Else


            rs.Open sql1, conexao
            'grupo = rs.Fields("grupo").Value
            idNv2 = rs.Fields("idNv2").Value

            UserForm2.ed_txtBoxID_nv2.Value = idNv2
            ed_id_nv2 = idNv2
            rs.Close
            conexao.Close



            End If

                GoTo jumpit
here:             Exit Sub
jumpit:


            '----


End Sub
'
'
'
'
Sub ed_GetIdNv3()

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
selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
Dim selectedRow1 As String
selectedRow1 = UserForm2.ed_txtBox_nv3_3.Value


If ed_sPrincipais = True Then

    If Not IsNull(UserForm2.ed_txtComboBox_nv3_3.Value) And Len(UserForm2.ed_txtComboBox_nv3_3.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv3_3.Value
        selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS PRINCIPAIS'"
   ElseIf Not IsNull(UserForm2.ed_txtBox_nv3_3.Value) And Len(UserForm2.ed_txtBox_nv3_3.Value) > 0 Then

            On Error GoTo here
            selectedRow1 = UserForm2.ed_txtBox_nv3_3.Value
'            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "';"
            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "' AND grupo = 'SERVICOS PRINCIPAIS';"
         


   End If
ElseIf ed_sDiversos = True Then

    If Not IsNull(UserForm2.ed_txtComboBox_nv3_3.Value) And Len(UserForm2.ed_txtComboBox_nv3_3.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv3_3.Value
        selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DIVERSOS'"
   ElseIf Not IsNull(UserForm2.ed_txtBox_nv3_3.Value) And Len(UserForm2.ed_txtBox_nv3_3.Value) > 0 Then

            On Error GoTo here
            selectedRow1 = UserForm2.ed_txtBox_nv3_3.Value

            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "' AND grupo = 'SERVICOS DIVERSOS';"
            '---


   End If
ElseIf ed_sTerceiros = True Then

    If Not IsNull(UserForm2.ed_txtComboBox_nv3_3.Value) And Len(UserForm2.ed_txtComboBox_nv3_3.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv3_3.Value
        selectedRow = UserForm2.ed_txtComboBox_nv3_3.ListIndex
'        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "'"
        sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Column(0, selectedRow) & "' AND grupo = 'SERVICOS DE TERCEIROS'"
   ElseIf Not IsNull(UserForm2.ed_txtBox_nv3_3.Value) And Len(UserForm2.ed_txtBox_nv3_3.Value) > 0 Then

            On Error GoTo here
            selectedRow1 = UserForm2.ed_txtBox_nv3_3.Value
'            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "';"
            sql1 = "SELECT idNv3 FROM t_Nivel3 WHERE descricaoNv3 = '" & selectedRow1 & "' AND grupo = 'SERVICOS DE TERCEIROS';"
        


   End If
End If






            If selectedRow = -1 Then
            idNv3 = 0
            UserForm2.ed_txtBoxID_nv3.Value = idNv3


            If (selectedRow = -1 And UserForm2.ed_txtBox_nv3_3.Value <> "Digite o serviço aqui") And (selectedRow = -1 And UserForm2.ed_txtBox_nv3_3.Value <> "") Then

            rs.Open sql1, conexao

            idNv3 = rs.Fields("idNv3").Value
            UserForm2.ed_txtBoxID_nv3.Value = idNv3
            ed_id_nv3 = idNv3
            rs.Close
            conexao.Close

            End If

            Else

 
            rs.Open sql1, conexao
            'grupo = rs.Fields("grupo").Value
            On Error GoTo notFound
            idNv3 = rs.Fields("idNv3").Value

            UserForm2.ed_txtBoxID_nv3.Value = idNv3
            ed_id_nv3 = idNv3
            rs.Close
            conexao.Close



            End If

                GoTo jumpit
notFound:
MsgBox "Item não encontrado na base de dados.", vbExclamation

If ed_sDiversos = True And UserForm2.ed_btnRendimentoBoolean = True Then
UserForm2.ed_txtComboBox_nv4_4.Clear
UserForm2.ed_txtComboBox_nv4_4.Enabled = False
UserForm2.ed_txtComboBox_nv4_4.BackColor = &H8000000F

UserForm2.ed_txtBox_rendimento.Enabled = False
UserForm2.ed_txtBox_rendimento.BackColor = &H8000000F
UserForm2.ed_txtBox_rendimento.Value = ""
Exit Sub
End If

Exit Sub
here:
 ed_id_nv3 = 0
Exit Sub
jumpit:


            '----


End Sub
'

'
'
Sub ed_GetCustoInsumo()

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
selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
Dim selectedRow1 As String
selectedRow1 = UserForm2.txtBox_nv4_4.Value



   Dim valor As String
    If Not IsNull(UserForm2.ed_txtComboBox_nv4_4.Value) And Len(UserForm2.ed_txtComboBox_nv4_4.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv4_4.Value
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        sql1 = "SELECT Custo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(UserForm2.txtBox_nv4_4.Value) And Len(UserForm2.txtBox_nv4_4.Value) > 0 Then

            selectedRow1 = UserForm2.txtBox_nv4_4.Value
            sql1 = "SELECT Custo FROM t_Insumos WHERE Insumo = '" & selectedRow1 & "';"
            '---

   End If





        rs.Open sql1, conexao
''grupo = rs.Fields("grupo").Value
        On Error GoTo here
        CUSTO = rs.Fields("Custo").Value

        If selectedRow = -1 And selectedRow1 = "" Then
        UserForm2.ed_txtBox_custoInsumo.Enabled = True
'        UserForm2.ed_txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
        UserForm2.ed_txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
        UserForm2.ed_txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
        ed_txtBox_custoInsumoPublic = ""
        '===== teste =====
            If selectedRow = -1 And selectedRow1 = "" And UserForm2.txtBox_nv4_4 = "" Then
           UserForm2.ed_txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
'            UserForm2.ed_txtBox_custoInsumo.Value = ""
            ed_txtBox_custoInsumoPublic = Format("", "R$ #,##0.00")
            UserForm2.ed_txtBox_custoInsumo.Enabled = False
            UserForm2.ed_txtBox_custoInsumo.BackColor = &H80000016
            End If
        '==============
        Else
        UserForm2.ed_txtBox_custoInsumo.Value = Format(CUSTO, "R$ #,##0.00")
'        UserForm2.ed_txtBox_custoInsumo.Value = Custo
        ed_txtBox_custoInsumoPublic = Format(CUSTO, "R$ #,##0.00")
'       ed_txtBox_custoInsumoPublic = Format(CUSTO, "R$ #,##0.00")
'        UserForm2.ed_txtBox_custoInsumo.Enabled = False
'        UserForm2.ed_txtBox_custoInsumo.BackColor = &H80000016
        UserForm2.ed_txtBox_custoInsumo.Enabled = True
        UserForm2.ed_txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
        End If
        rs.Close
        conexao.Close


GoTo jumpit

If (i > 1) Then
here:
UserForm2.ed_txtBox_custoInsumo.Enabled = True
UserForm2.ed_txtBox_custoInsumo.BackColor = RGB(255, 255, 255)
'MsgBox "Insira um novo custo!"
'UserForm2.ed_txtBox_custoInsumo.Value = Format("", "R$ #,##0.00")
UserForm2.ed_txtBox_custoInsumo.Value = ""
ed_txtBox_custoInsumoPublic = ""
End If
Exit Sub
jumpit:



End Sub
'
Sub ed_GetUnidade()

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
selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
Dim selectedRow1 As String
selectedRow1 = UserForm2.txtBox_nv4_4.Value


   Dim valor As String
    If Not IsNull(UserForm2.ed_txtComboBox_nv4_4.Value) And Len(UserForm2.ed_txtComboBox_nv4_4.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv4_4.Value
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        sql1 = "SELECT Unidade FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(UserForm2.txtBox_nv4_4.Value) And Len(UserForm2.txtBox_nv4_4.Value) > 0 Then

            selectedRow1 = UserForm2.txtBox_nv4_4.Value
            sql1 = "SELECT Unidade FROM t_Insumos WHERE Insumo = '" & selectedRow1 & "';"

   End If


rs.Open sql1, conexao
'grupo = rs.Fields("grupo").Value
On Error GoTo here
unidade = rs.Fields("Unidade").Value






If unidade <> "" Then
    If selectedRow1 = "" And selectedRow = -1 Then
    UserForm2.ed_txtBox_un.Value = ""
    UserForm2.ed_txtBox_un.Enabled = True
    UserForm2.ed_txtBox_un.BackColor = RGB(255, 255, 255)
    
        '=======TESTE===
        If selectedRow1 = "" And selectedRow = -1 And UserForm2.txtBox_nv4_4 = "" Then
        UserForm2.ed_txtBox_un.Value = ""
        UserForm2.ed_txtBox_un.Enabled = False
        UserForm2.ed_txtBox_un.BackColor = &H80000016
        End If

    Else
    UserForm2.ed_txtBox_un.Value = unidade
    ed_txtBox_unPublic = unidade
'    UserForm2.ed_txtBox_un.Enabled = False
'    UserForm2.ed_txtBox_un.BackColor = &H80000016
    UserForm2.ed_txtBox_un.Enabled = True
    UserForm2.ed_txtBox_un.BackColor = RGB(255, 255, 255)
    End If

Else

    '
    UserForm2.ed_txtBox_un.Value = "-"
    ed_txtBox_unPublic = "-"
'    UserForm2.ed_txtBox_un.Enabled = False
'    UserForm2.ed_txtBox_un.BackColor = &H80000016
    UserForm2.ed_txtBox_un.Enabled = True
    UserForm2.ed_txtBox_un.BackColor = RGB(255, 255, 255)
End If

rs.Close
conexao.Close

GoTo jumpit
here:
    UserForm2.ed_txtBox_un.Value = ""
    UserForm2.ed_txtBox_un.Enabled = True
    UserForm2.ed_txtBox_un.BackColor = RGB(255, 255, 255)
Exit Sub

jumpit:

End Sub



Sub ed_GetTipoInsumo()



Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim i As Integer
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

 

Dim selectedRow As Integer
selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex

ConectarBanco conexao
If Not IsNull(UserForm2.ed_txtComboBox_nv4_4.Value) And Len(UserForm2.ed_txtComboBox_nv4_4.Value) > 0 Then
     selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
     sql1 = "SELECT Tipo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
End If

On Error GoTo exitSub
rs.Open sql1, conexao
strTipo = rs.Fields("Tipo").Value

UserForm2.ed_ComboBoxTipo.Clear




If ed_sPrincipais = True And strTipo = "SERVICOS PRINCIPAIS" Then
UserForm2.ed_ComboBoxTipo.AddItem "PRINCIPAL"
UserForm2.ed_ComboBoxTipo.AddItem "GENÉRICO"
UserForm2.ed_ComboBoxTipo.ListIndex = 0

ElseIf ed_sPrincipais = True And strTipo = "GENERICO" Then
UserForm2.ed_ComboBoxTipo.AddItem "GENÉRICO"
UserForm2.ed_ComboBoxTipo.AddItem "PRINCIPAL"
UserForm2.ed_ComboBoxTipo.ListIndex = 0

ElseIf ed_sDiversos = True And strTipo = "SERVICOS DIVERSOS" Then
UserForm2.ed_ComboBoxTipo.AddItem "DIVERSO"
UserForm2.ed_ComboBoxTipo.AddItem "GENÉRICO"
UserForm2.ed_ComboBoxTipo.ListIndex = 0

ElseIf ed_sDiversos = True And strTipo = "GENERICO" Then
UserForm2.ed_ComboBoxTipo.AddItem "GENÉRICO"
UserForm2.ed_ComboBoxTipo.AddItem "DIVERSO"
UserForm2.ed_ComboBoxTipo.ListIndex = 0

End If

GoTo jumpit

exitSub:
    Exit Sub
    
jumpit:
End Sub


Sub SetTipo()
Dim result As VbMsgBoxResult
Dim varToCompare As String

If UserForm2.ed_ComboBoxTipo = "PRINCIPAL" Then
varToCompare = "SERVICOS PRINCIPAIS"
ElseIf UserForm2.ed_ComboBoxTipo = "DIVERSO" Then
varToCompare = "SERVICOS DIVERSOS"
ElseIf UserForm2.ed_ComboBoxTipo = "GENÉRICO" Then
varToCompare = "GENERICO"
End If

If strTipo <> varToCompare Then
'Trará a quantidade de diversos e principais atrelados ao insumo
GetQuantidadePorTipoDeServico
GetDefinirAcaoEdicaoTipo
'UserForm2.ed_btnNv4_carregar

Else
MsgBox "O tipo de insumo foi mantido.", vbInformation
Exit Sub

End If




End Sub


Sub GetQuantidadePorTipoDeServico()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim insumo As String
Dim nivel1 As String
Dim nivel2 As String
Dim nivel3 As String
Dim nivel4 As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim id As String
'Dim rs As Recordset

Dim db As Object
Dim idNv1 As Long
Dim idNv2 As Long
Dim idNv3 As Long
Dim idInsumo As Long
Dim result As VbMsgBoxResult


'UserForm2.ex_ListBox.Clear


    ConectarBanco conexao
    sql = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4 & "';"
    rs.Open sql, conexao
    idInsumo = rs.Fields("idInsumo").Value
    id = idInsumo
    conexao.Close

    ConectarBanco conexao
    sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idInsumo = " & id & ";"
    rs.Open sql, conexao
    counte = rs.Fields("counte").Value
    counte = rs("counte")
    rs.Close
    conexao.Close
    

    
    ConectarBanco conexao
    sql = "SELECT COUNT(*) AS counte1 FROM t_Servicos_Diversos_Rendimento WHERE idInsumo = " & id & ";"
    rs.Open sql, conexao
    counte1 = rs.Fields("counte1").Value
    counte1 = rs("counte1")
    rs.Close
    conexao.Close
    Dim counter As Integer
    counter = counte + counte1
    
    
    
    
End Sub

Sub GetDefinirAcaoEdicaoTipo()
Dim selectedRow As Integer
    
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
'counte1 = diversos
'counte = principais
If ed_sPrincipais = True Then

    If strTipo = "GENERICO" Then
        If counte1 <> 0 And counte <> 0 Then
    
        MsgBox "A mudança para o tipo Principal não é permitida." & vbNewLine & _
               "Excluir extruturas que possuam o insumo em Serviços Diversos, em:" & vbNewLine & _
               "" & vbNewLine & _
               "- Menu Inicial" & vbNewLine & _
               "- Exclusão" & vbNewLine & _
               "- Insumo Geral" & vbNewLine & _
               "- Estrutura" & vbNewLine & _
               "- Serviços Diversos" & vbNewLine & _
               "", vbInformation
        Exit Sub

        ElseIf counte <> 0 And counte1 = 0 Or counte = 0 And counte1 = 0 Then
        'MsgBox "A mudança para o tipo Principal permitida! Deseja prosseguir?"
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
        
        new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close
        
        ConectarBanco conexao
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs!tipo = "SERVICOS PRINCIPAIS"
        
        rs.Update
        rs.Close
        conexao.Close
        
        log_edicao_insumo
        MsgBox "Edição de tipo de insumo realizada.", vbInformation
        Else

         MsgBox "A mudança para o tipo Principal não é permitida." & vbNewLine & _
        "Excluir extruturas que possuam o insumo em Serviços Diversos, em:" & vbNewLine & _
        "" & vbNewLine & _
        "- Menu Inicial" & vbNewLine & _
        "- Exclusão" & vbNewLine & _
        "- Insumo Geral" & vbNewLine & _
        "- Estrutura" & vbNewLine & _
        "- Serviços Diversos" & vbNewLine & _
        "", vbInformation
        
        End If
        
    ElseIf strTipo = "SERVICOS PRINCIPAIS" Then
    
    

        If counte <> 0 And counte1 = 0 Then
        'MsgBox "Mudar para tipo genérico"
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
        
        new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close
        
        ConectarBanco conexao
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs!tipo = "GENERICO"
        
        rs.Update
        rs.Close
        conexao.Close
        
        log_edicao_insumo
        MsgBox "Edição de tipo de insumo realizada.", vbInformation
        ElseIf counte1 <> 0 Then
        'Existe no banco de dados o mesmo insumo sendo usado em estrutura de diversos e principais, o código não deve permitir isso. O tipo presente em ambos será sempre e apenas o genérico
        MsgBox "ERRO de tipagem de insumo! O item existe em Principais e Diversos! Deve ser utilizado o tipo Genérico", vbCritical
        End If
        
    End If
    

ElseIf ed_sDiversos = True Then


If strTipo = "GENERICO" Then

    If counte1 <> 0 And counte <> 0 Then
    
        MsgBox "A mudança para o tipo Diverso não é permitida." & vbNewLine & _
               "Excluir extruturas que possuam o insumo em Serviços Principais, em:" & vbNewLine & _
               "" & vbNewLine & _
               "- Menu Inicial" & vbNewLine & _
               "- Exclusão" & vbNewLine & _
               "- Insumo Geral" & vbNewLine & _
               "- Estrutura" & vbNewLine & _
               "- Serviços Principais" & vbNewLine & _
               "- Exclusão de Estrutura" & vbNewLine & _
               "", vbInformation
        Exit Sub

        ElseIf counte1 <> 0 And counte = 0 Then
        'MsgBox "A mudança para o tipo Diverso permitido! Deseja prosseguir?"
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
        
        new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close
        
        ConectarBanco conexao
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs!tipo = "SERVICOS DIVERSOS"
        
        rs.Update
        rs.Close
        conexao.Close
        
        log_edicao_insumo
        MsgBox "Edição de tipo de insumo realizada.", vbInformation
        
        ElseIf counte1 = 0 And counte = 0 Then
        'MsgBox "A mudança para o tipo Diverso permitido! Deseja prosseguir?"
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
        
        new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close
        
        ConectarBanco conexao
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs!tipo = "SERVICOS DIVERSOS"
        
        rs.Update
        rs.Close
        conexao.Close
        
        log_edicao_insumo
        MsgBox "Edição de tipo de insumo realizada.", vbInformation
        Else

         MsgBox "A mudança para o tipo Diverso não é permitida." & vbNewLine & _
        "Excluir extruturas que possuam o insumo em Serviços Principais, em:" & vbNewLine & _
        "" & vbNewLine & _
        "- Menu Inicial" & vbNewLine & _
        "- Exclusão" & vbNewLine & _
        "- Insumo Geral" & vbNewLine & _
        "- Estrutura" & vbNewLine & _
        "- Serviços Principais" & vbNewLine & _
        "", vbInformation
        
        End If
        
    ElseIf strTipo = "SERVICOS DIVERSOS" Then
    
    

        If counte1 <> 0 And counte = 0 Then
        'MsgBox "Mudar para tipo genérico"
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        ConectarBanco conexao
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
        rs.Open sql1, conexao
        
        new_id4 = rs.Fields("idInsumo").Value
        rs.Close
        conexao.Close
        
        ConectarBanco conexao
        sql = "SELECT Tipo FROM t_Insumos WHERE idInsumo = " & new_id4
        rs.Open sql, conexao, adOpenKeyset, adLockOptimistic
        rs!tipo = "GENERICO"
        
        rs.Update
        rs.Close
        conexao.Close
        
        log_edicao_insumo
        MsgBox "Edição de tipo de insumo realizada.", vbInformation
        ElseIf counte1 <> 0 Then
        'Existe no banco de dados o mesmo insumo sendo usado em estrutura de diversos e principais, o código não deve permitir isso. O tipo presente em ambos será sempre e apenas o genérico
        MsgBox "ERRO de tipagem de insumo! O item existe em Principais e Diversos! Deve ser utilizado o tipo Genérico", vbCritical

        End If
        
    End If

End If


End Sub


Sub emStandBy()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim insumo As String
Dim nivel1 As String
Dim nivel2 As String
Dim nivel3 As String
Dim nivel4 As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim id As String
'Dim rs As Recordset
Dim counte As Integer
Dim db As Object
Dim idNv1 As Long
Dim idNv2 As Long
Dim idNv3 As Long
Dim idInsumo As Long
Dim result As VbMsgBoxResult


'---------------------- LISTAR VALORES -------------------------------
        UserForm2.ex_ListBox.Clear
        UserForm2.ex_ListBox.ColumnCount = 4
        UserForm2.ex_ListBox.AddItem ""

        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = "Insumo"
        
If counte <> 0 Then

        ConectarBanco conexao
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idInsumo = " & id & ";"

        rs.Open sql, conexao
        rs.MoveFirst
            
    
        If counte > 0 Then

    ' Ajusta o tamanho das colunas
    UserForm2.ex_ListBox.columnWidths = "100;100;400;200" ' Ajuste o tamanho conforme necessário
            
        Do Until rs.EOF
            
        Dim strNv1 As Integer
       strNv1 = rs("idNv1")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nive1SQL As String
        nive1SQL = "SELECT descricaoNv1 FROM t_Nivel1 WHERE idNv1 = " & strNv1 & ";"
        Dim rsDescricaoNv1 As ADODB.Recordset
        Set rsDescricaoNv1 = New ADODB.Recordset
        rsDescricaoNv1.Open nive1SQL, conexao
        
    
       Dim strNv2 As Integer
       strNv2 = rs("idNv2")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nivel2SQL As String
        nivel2SQL = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & strNv2 & ";"
        Dim rsDescricaoNv2 As ADODB.Recordset
        Set rsDescricaoNv2 = New ADODB.Recordset
        rsDescricaoNv2.Open nivel2SQL, conexao
        
      Dim strNv3P As Integer
       strNv3P = rs("idNv3")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nivel3SQLP As String
        nivel3SQLP = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & strNv3P & ";"
        Dim rsDescricaoNv3P As ADODB.Recordset
        Set rsDescricaoNv3P = New ADODB.Recordset
        rsDescricaoNv3P.Open nivel3SQLP, conexao
               
    
       Dim strInsumoP As Integer
       strInsumoP = rs("idInsumo")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim insumoSQLP As String
        insumoSQLP = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & strInsumoP & ";"
        Dim rsPInsumo As ADODB.Recordset
       Set rsPInsumo = New ADODB.Recordset
        rsPInsumo.Open insumoSQLP, conexao
        
                  
        UserForm2.ex_ListBox.ColumnCount = 4
        ' Verifica se há um valor correspondente na tabela t_Insumos
        If Not rsDescricaoNv1.EOF Then
             'Obtém o valor de Insumo encontrado na tabela t_Insumos

'            insumo = rsPInsumo("Insumo")
            nivel3 = rsDescricaoNv3P("DescricaoNv3")
            nivel2 = rsDescricaoNv2("DescricaoNv2")
            nivel1 = rsDescricaoNv1("DescricaoNv1")
            nivel4 = rsPInsumo("Insumo")

            
    ' Adiciona uma nova linha ao ListBox
        UserForm2.ex_ListBox.AddItem ""
    
        'Define os valores das colunas para a nova linha
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = nivel1
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = nivel2
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = nivel3
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = nivel4

    
        ' Ajusta o tamanho das colunas
        UserForm2.ex_ListBox.columnWidths = "100;150;300;300" ' Ajuste o tamanho conforme necessário
    
        End If
            rsPInsumo.Close
            rs.MoveNext
        Loop
 
    End If
    rs.Close
    conexao.Close
 End If
 
End Sub
'
''Verifica se há rendimento p serviços principais
Sub ed_GetRendimento()

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


If UserForm2.ed_txtBoxID_nv1.Value = "" Then
ed_id_nv1 = 0
UserForm2.ed_txtBoxID_nv1.Value = 0
Else
ed_id_nv1 = UserForm2.ed_txtBoxID_nv1.Value
End If

If UserForm2.ed_txtBoxID_nv2.Value = "" Then
    ed_id_nv2 = 0
    UserForm2.ed_txtBoxID_nv2.Value = 0
Else
ed_id_nv2 = UserForm2.ed_txtBoxID_nv2.Value
End If

If UserForm2.ed_txtBoxID_nv3.Value = "" Then
    UserForm2.ed_txtBoxID_nv3.Value = 0
    ed_id_nv3 = UserForm2.ed_txtBoxID_nv3.Value
Else
ed_id_nv3 = UserForm2.ed_txtBoxID_nv3.Value
End If

'sql = "t_Servicos_Principais_Insumos"
sql1 = "t_Servicos_Principais_Insumos"

id_master = ed_id_nv1 & "-" & ed_id_nv2 & "-" & ed_id_nv3


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex

   Dim valor As String
    If Not IsNull(UserForm2.ed_txtComboBox_nv4_4.Value) And Len(UserForm2.ed_txtComboBox_nv4_4.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv4_4.Value
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        'FUNCIONANDO'sql1 = "SELECT rendiUserForm2nto FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & "ACRILICA SEMI-BRILHO 18LT" & "'"
        sql1 = "SELECT rendimento FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            'sql1 = "SELECT * FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "';"

   ElseIf Not IsNull(UserForm2.txtBox_nv4_4.Value) And Len(UserForm2.txtBox_nv4_4.Value) > 0 Then
 
        sql1 = "SELECT rendimento FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & UserForm2.txtBox_nv4_4.Value & "'"

   'End If
    Else
    Exit Sub
    End If
'===============================================================


Dim insumo As String



rs.Open sql1, conexao

On Error GoTo here
rendimento = rs.Fields("rendimento").Value
If (rendimento <> "") And (selectedRow <> -1 Or UserForm2.ed_txtBox_nv4_4 <> "") Then
    'rendiUserForm2nto = rs.Fields("rendiUserForm2nto").Value
    'UserForm2.ed_txtBox_un.Value = idNV1
    If UserForm2.ed_txtComboBox_nv4_4.Value <> "" Then
        insumo = UserForm2.ed_txtComboBox_nv4_4.Value
        ed_c_rendimento = True
    Else
        insumo = UserForm2.ed_txtBox_nv4_4.Value
        ed_c_rendimento = True
    End If

    'MsgBox "A soma dos IDs é existente, os valores existem na tabela, sendo ele = " & id_master & " / De valor de rendiUserForm2nto = " & RENDIMENTO & " Com noUserForm2 de insumo = " & insumo
    UserForm2.ed_txtBox_rendimento.Value = Format(rendimento, "#,##0.00")
    ed_rendimento = Format(rendimento, "#,##0.00")

    UserForm2.ed_txtBox_rendimento.BackColor = RGB(255, 255, 255)
    UserForm2.ed_txtBox_rendimento.Enabled = True

    ed_c_rendimento = True

Else
    'UserForm2.ed_txtBox_un.Value = "-"
here:
ed_c_rendimento = False

    UserForm2.ed_txtBox_rendimento.BackColor = &H80000016
    UserForm2.ed_txtBox_rendimento.Enabled = False
    ed_txtBox_rendimento = True
    'UserForm2.ed_txtBox_un.Value = ""
    UserForm2.ed_txtBox_rendimento.Value = ""

End If
rs.Close
conexao.Close


End Sub
'Verifica se há rendiUserForm2ntos para serviços diversos
Sub ed_GetRendimentoDiversos()

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


If UserForm2.ed_txtBoxID_nv1.Value = "" Then
ed_id_nv1 = 0
UserForm2.ed_txtBoxID_nv1.Value = 0
Else
ed_id_nv1 = UserForm2.ed_txtBoxID_nv1.Value
End If

If UserForm2.ed_txtBoxID_nv2.Value = "" Then
    ed_id_nv2 = 0
    UserForm2.ed_txtBoxID_nv2.Value = 0
Else
ed_id_nv2 = UserForm2.ed_txtBoxID_nv2.Value
End If

If UserForm2.ed_txtBoxID_nv3.Value = "" Then
    UserForm2.ed_txtBoxID_nv3.Value = 0
    ed_id_nv3 = UserForm2.ed_txtBoxID_nv3.Value
Else
ed_id_nv3 = UserForm2.ed_txtBoxID_nv3.Value
End If

'sql = "t_Servicos_Principais_Insumos"
sql1 = "t_Servicos_Principais_Insumos"

id_master = 7 & "-" & 0 & "-" & ed_id_nv3


Dim selectedRow As Integer
Dim coluna2Valor As String
'Dim sql As String

'Obter o índice da linha selecionada no listbox
selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex



   Dim valor As String
    If Not IsNull(UserForm2.ed_txtComboBox_nv4_4.Value) And Len(UserForm2.ed_txtComboBox_nv4_4.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv4_4.Value
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        'FUNCIONANDO'sql1 = "SELECT rendiUserForm2nto FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & "ACRILICA SEMI-BRILHO 18LT" & "'"
        sql1 = "SELECT rendimento FROM t_Servicos_Diversos_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
            'sql1 = "SELECT * FROM t_Servicos_Principais_Insumos WHERE ID = '" & id_master & "';"

   ElseIf Not IsNull(UserForm2.txtBox_nv4_4.Value) And Len(UserForm2.txtBox_nv4_4.Value) > 0 Then
 
        sql1 = "SELECT rendimento FROM t_Servicos_Diversos_Insumos WHERE ID = '" & id_master & "' AND Insumo = '" & UserForm2.txtBox_nv4_4.Value & "'"

   'End If
    Else
    Exit Sub
    End If
'===============================================================


Dim insumo As String


rs.Open sql1, conexao

On Error GoTo here
rendimento = rs.Fields("rendimento").Value
If (rendimento <> "") And (selectedRow <> -1 Or UserForm2.txtBox_nv4_4 <> "") Then
    'rendiUserForm2nto = rs.Fields("rendiUserForm2nto").Value
    'UserForm2.ed_txtBox_un.Value = idNV1
    If UserForm2.ed_txtComboBox_nv4_4.Value <> "" Then
        insumo = UserForm2.ed_txtComboBox_nv4_4.Value
        ed_c_rendimento = True
    Else
        insumo = UserForm2.txtBox_nv4_4.Value
        ed_c_rendimento = True
    End If

    'MsgBox "A soma dos IDs é existente, os valores existem na tabela, sendo ele = " & id_master & " / De valor de rendiUserForm2nto = " & RENDIMENTO & " Com noUserForm2 de insumo = " & insumo
    UserForm2.ed_txtBox_rendimento.Value = Format(rendimento, "#,##0.00")
    ed_rendimento = Format(rendimento, "#,##0.00")
'    UserForm2.ed_txtBox_rendimento.BackColor = &H80000016
'    UserForm2.ed_txtBox_rendimento.Enabled = False
    UserForm2.ed_txtBox_rendimento.BackColor = RGB(255, 255, 255)
    UserForm2.ed_txtBox_rendimento.Enabled = True
    ed_rendimentoValue = True

Else
    'UserForm2.ed_txtBox_un.Value = "-"
here:
ed_c_rendimento = False
    'MsgBox "Serviço inválido"
'    UserForm2.ed_txtBox_rendimento.BackColor = RGB(255, 255, 255)
'    UserForm2.ed_txtBox_rendimento.Enabled = True
    UserForm2.ed_txtBox_rendimento.BackColor = &H80000016
    UserForm2.ed_txtBox_rendimento.Enabled = False
    UserForm2.ed_txtBox_rendimento.Value = ""
    'ed_rendimento = ""
    ed_rendimentoValue = True
    'UserForm2.ed_txtBox_un.Value = ""

    'UserForm2.txtBox_custoInsumo.Value = ""
End If
rs.Close
conexao.Close


End Sub
'
'
'
Sub ed_CarregarNivelDois()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel2"

rs.Open "select descricaoNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv2", conexao, 3, 3

UserForm2.ed_txtComboBox_nv2_2.Clear
Do Until rs.EOF
UserForm2.ed_txtComboBox_nv2_2.AddItem rs!descricaoNv2

rs.MoveNext
Loop

conexao.Close



End Sub
'
'
Sub ed_CarregarNivelUm()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel1"


'
'If sPrincipais = True Then
rs.Open "select descricaoNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv1", conexao, 3, 3
'End If


UserForm2.ed_txtComboBox_nv1_1.Clear
Do Until rs.EOF
UserForm2.ed_txtComboBox_nv1_1.AddItem rs!descricaoNv1

rs.MoveNext
Loop

conexao.Close


End Sub
'
'
Sub ed_CarregarNivelTres()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel3"



If ed_sPrincipais = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv3", conexao, 3, 3
End If


If ed_sTerceiros = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DE TERCEIROS' order BY idNv3", conexao, 3, 3
End If

If ed_sDiversos = True Then

rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv3", conexao, 3, 3
End If


UserForm2.ed_txtComboBox_nv3_3.Clear
Do Until rs.EOF

UserForm2.ed_txtComboBox_nv3_3.AddItem rs!descricaoNv3

rs.MoveNext
Loop

conexao.Close




End Sub




Sub ed_CarregarInsumos()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Insumos"

'rs.Open "select Insumo from t_Insumos order BY idInsumo", conexao, 3, 3
If ed_sPrincipais = True Then
rs.Open "SELECT Insumo FROM t_Insumos WHERE (Tipo = 'SERVICOS PRINCIPAIS' OR Tipo = 'GENERICO') ORDER BY idInsumo", conexao, 3, 3
ElseIf ed_sDiversos = True Then
rs.Open "SELECT Insumo FROM t_Insumos WHERE (Tipo = 'SERVICOS DIVERSOS' OR Tipo = 'GENERICO') ORDER BY idInsumo", conexao, 3, 3
End If

UserForm2.ed_txtComboBox_nv4_4.Clear
Do Until rs.EOF
UserForm2.ed_txtComboBox_nv4_4.AddItem rs!insumo

rs.MoveNext
Loop

conexao.Close



End Sub




Sub ed_GetIdInsumo()

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
selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
Dim selectedRow1 As String
selectedRow1 = ""
'


   Dim valor As String
    If Not IsNull(UserForm2.ed_txtComboBox_nv4_4.Value) And Len(UserForm2.ed_txtComboBox_nv4_4.Value) > 0 Then
        valor = UserForm2.ed_txtComboBox_nv4_4.Value
        selectedRow = UserForm2.ed_txtComboBox_nv4_4.ListIndex
        sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Column(0, selectedRow) & "'"
   ElseIf Not IsNull(UserForm2.ed_txtBox_nv4_4.Value) And Len(UserForm2.ed_txtBox_nv4_4.Value) > 0 Then
'        valor = UserForm2.txtBox_nv4_4.Value
'        selectedRow = UserForm2.txtBox_nv4_4.Value
'        sql1 = "SELECT * FROM t_Insumos WHERE idInsumo = '" & selectedRow & "';"
            '---TESTE
            On Error GoTo here
            selectedRow1 = UserForm2.ed_txtBox_nv4_4.Value
            sql1 = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & selectedRow1 & "';"
            '---
   End If


            If selectedRow = -1 Then
            idInsumo = 0
            UserForm2.ed_txtBoxID_nv4.Value = idInsumo
            ed_id_nv4 = idInsumo
            
            If selectedRow = -1 And UserForm2.ed_txtBox_nv4_4.Value <> "" Then
            
            rs.Open sql1, conexao
            
            idInsumo = rs.Fields("idInsumo").Value
            UserForm2.ed_txtBoxID_nv4.Value = idInsumo
            ed_id_nv4 = idInsumo
            Insumo_nv4 = selectedRow1
            If Insumo_nv4 <> "" Then
            insumoBolean = True
            End If
            
            rs.Close
            conexao.Close
            
            End If
            
            Else
            

            rs.Open sql1, conexao
            'grupo = rs.Fields("grupo").Value
            idInsumo = rs.Fields("idInsumo").Value

            UserForm2.ed_txtBoxID_nv4.Value = idInsumo
            idInsumo = idInsumo
            ed_id_nv4 = idInsumo
            rs.Close
            conexao.Close
            
            
            
            End If
            
                GoTo jumpit
here:
            idInsumo = 0
            UserForm2.ed_txtBoxID_nv4.Value = idInsumo
            ed_id_nv4 = idInsumo
Exit Sub
jumpit:
            
            
            '----

End Sub




Sub ed_carregarEstruturaPrincipais2()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
        Dim verifica_repitidos
        Dim numero_item1
        Dim numero_item2
Dim idNv1 As Long
Dim id As Long

    'If UserForm2.ed_txtComboBox_nv2_2.Enabled = True And UserForm2.ed_txtComboBox_nv1_1.Value <> "" And UserForm2.ed_txtComboBox_nv3_3.Enabled = False Then
    If UserForm2.ed_txtComboBox_nv2_2.Enabled = True And UserForm2.ed_txtComboBox_nv1_1.Value <> "" Then



       UserForm2.ed_txtComboBox_nv2_2.Clear
        ConectarBanco conexao
'        If UserForm2.ed_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' and descricaoNv1 = '" & UserForm2.ed_txtComboBox_nv1_1.Value & "'"

        rs.Open sql, conexao
        idNv1 = rs.Fields("idNv1").Value
        id = idNv1
        conexao.Close
    
        ConectarBanco conexao
        
'        If UserForm2.ed_sPrincipais = True Or UserForm2.ed_BtnSelectionPrincipaisBoolean = True Then
        rs.Open "SELECT idNv2 FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & " ORDER BY idNv2", conexao, 3, 3


        Do Until rs.EOF
        UserForm2.ed_txtComboBox_nv2_2.AddItem rs!idNv2
        rs.MoveNext
        Loop
        conexao.Close
        

        ConectarBanco conexao
        
        
        For verifica_repitidos = 0 To UserForm2.ed_txtComboBox_nv2_2.ListCount - 1
         sql = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & UserForm2.ed_txtComboBox_nv2_2.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ed_txtComboBox_nv2_2.List(verifica_repitidos) = rs!descricaoNv2
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         

        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ed_txtComboBox_nv2_2.ListCount - 1
         For numero_item2 = 0 To UserForm2.ed_txtComboBox_nv2_2.ListCount - 1
             If numero_item1 > UserForm2.ed_txtComboBox_nv2_2.ListCount - 1 Or numero_item2 > UserForm2.ed_txtComboBox_nv2_2.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ed_txtComboBox_nv2_2.List(numero_item1) = UserForm2.ed_txtComboBox_nv2_2.List(numero_item2) Then
                         UserForm2.ed_txtComboBox_nv2_2.RemoveItem (numero_item2)
                     Else
                     End If
                 End If
             End If
         Next numero_item2
         Next numero_item1
        
        Next verifica_repitidos
        
    End If




End Sub



Sub ed_carregarEstruturaPrincipais4()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
        Dim verifica_repitidos
        Dim numero_item1
        Dim numero_item2
Dim idNv3 As Long
Dim idNv2 As Long
Dim idNv1 As Long
Dim id3 As Long
Dim id2 As Long
Dim id1 As Long



    If UserForm2.ed_txtComboBox_nv1_1.Value <> "" And UserForm2.ed_txtComboBox_nv2_2.Value <> "" And UserForm2.ed_txtComboBox_nv3_3.Value <> "" Or ed_sDiversos = True Then
    
    
'ID1
            ConectarBanco conexao

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
        If ed_sPrincipais = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv1 = '" & UserForm2.ed_txtComboBox_nv1_1.Value & "'"
        rs.Open sql, conexao
        On Error GoTo here:
        idNv1 = rs.Fields("idNv1").Value
        id1 = idNv1
        conexao.Close
        ElseIf ed_sDiversos = True Then
        id1 = 7
        conexao.Close
        End If

 
'ID2
            ConectarBanco conexao
        If ed_sPrincipais = True Then
        sql = "select idNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv2 = '" & UserForm2.ed_txtComboBox_nv2_2.Value & "'"
        rs.Open sql, conexao
        On Error GoTo here:
        idNv2 = rs.Fields("idNv2").Value
        id2 = idNv2
        conexao.Close
        ElseIf ed_sDiversos = True Then
        id2 = 0
        conexao.Close
        End If




'ID3
       UserForm2.ed_txtComboBox_nv4_4.Clear
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        sql = "select idNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Value & "'"
        ElseIf ed_sDiversos = True Then
        sql = "select idNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' And descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Value & "'"
        End If

  
        rs.Open sql, conexao
        On Error GoTo here:
        idNv3 = rs.Fields("idNv3").Value
        id3 = idNv3
        conexao.Close
        
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        rs.Open "SELECT idInsumo FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id3 & " AND idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idInsumo", conexao, 3, 3
        ElseIf ed_sDiversos = True Then
        rs.Open "SELECT idInsumo FROM t_Servicos_Diversos_Rendimento WHERE idNv3 = " & id3 & " AND idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idInsumo", conexao, 3, 3
        End If

        Do Until rs.EOF
        UserForm2.ed_txtComboBox_nv4_4.AddItem rs!idInsumo
        rs.MoveNext
        Loop
        conexao.Close
        

        ConectarBanco conexao
        For verifica_repitidos = 0 To UserForm2.ed_txtComboBox_nv4_4.ListCount - 1
         sql = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & UserForm2.ed_txtComboBox_nv4_4.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ed_txtComboBox_nv4_4.List(verifica_repitidos) = rs!insumo
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         

        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ed_txtComboBox_nv4_4.ListCount - 1
         For numero_item2 = 0 To UserForm2.ed_txtComboBox_nv4_4.ListCount - 1
             If numero_item1 > UserForm2.ed_txtComboBox_nv4_4.ListCount - 1 Or numero_item2 > UserForm2.ed_txtComboBox_nv4_4.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ed_txtComboBox_nv4_4.List(numero_item1) = UserForm2.ed_txtComboBox_nv4_4.List(numero_item2) Then
                         UserForm2.ed_txtComboBox_nv4_4.RemoveItem (numero_item2)
                     Else
                     End If
                 End If
             End If
         Next numero_item2
         Next numero_item1
        
        Next verifica_repitidos
    End If
    GoTo jumpOverIt
here:
If id1 = 0 Or id2 = 0 Or id3 = 0 Then

UserForm2.ex_ComboBoxNv4.Enabled = False
UserForm2.ex_ComboBoxNv4.BackColor = &H8000000F

Exit Sub
Else
MsgBox "ed_carregarEstruturaPrincipais4"
End If

jumpOverIt:
    
End Sub



Sub ed_carregarEstruturaPrincipais3()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
        Dim verifica_repitidos
        Dim numero_item1
        Dim numero_item2
Dim idNv2 As Long
Dim id2 As Long
Dim idNv1 As Long
Dim id1 As Long

'    If UserForm2.ed_txtComboBox_nv1_1.Value <> "" And UserForm2.ed_txtComboBox_nv2_2 <> "" And UserForm2.ed_txtComboBox_nv3_3.Enabled = True Or UserForm2.ex_BtnSelectionDiversosBoolean = True And UserForm2.ed_txtComboBox_nv1_1.Enabled = False And UserForm2.ed_txtComboBox_nv2_2.Enabled = False Then
     If UserForm2.ed_txtComboBox_nv1_1.Value <> "" And UserForm2.ed_txtComboBox_nv2_2 <> "" And UserForm2.ed_txtComboBox_nv3_3.Enabled = True Or ed_sDiversos = True Then
'ID1




        
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv1 = '" & UserForm2.ed_txtComboBox_nv1_1.Value & "'"
        rs.Open sql, conexao
        idNv1 = rs.Fields("idNv1").Value
        id1 = idNv1
        conexao.Close
        ElseIf ed_sDiversos = True Then
        id1 = 7
        conexao.Close
        ElseIf ed_sTerceiros = True Then
        id1 = 9
        conexao.Close
        End If
  

    
'ID2
       UserForm2.ed_txtComboBox_nv3_3.Clear
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        sql = "select idNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv2 = '" & UserForm2.ed_txtComboBox_nv2_2.Value & "'"
        rs.Open sql, conexao
        idNv2 = rs.Fields("idNv2").Value
        id2 = idNv2
        conexao.Close
        ElseIf ed_sDiversos = True Then
        id2 = 0
        conexao.Close
        End If
  

        
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        rs.Open "SELECT idNv3 FROM t_Servicos_Principais_Rendimento WHERE idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idNv3", conexao, 3, 3
        ElseIf ed_sDiversos = True Then
        rs.Open "SELECT idNv3 FROM t_Servicos_Diversos_Rendimento WHERE idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idNv3", conexao, 3, 3
        End If
        

        Do Until rs.EOF
        UserForm2.ed_txtComboBox_nv3_3.AddItem rs!idNv3
        rs.MoveNext
        Loop
        conexao.Close
        


        ConectarBanco conexao
        For verifica_repitidos = 0 To UserForm2.ed_txtComboBox_nv3_3.ListCount - 1
         sql = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & UserForm2.ed_txtComboBox_nv3_3.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ed_txtComboBox_nv3_3.List(verifica_repitidos) = rs!descricaoNv3
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         


        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ed_txtComboBox_nv3_3.ListCount - 1
         For numero_item2 = 0 To UserForm2.ed_txtComboBox_nv3_3.ListCount - 1
             If numero_item1 > UserForm2.ed_txtComboBox_nv3_3.ListCount - 1 Or numero_item2 > UserForm2.ed_txtComboBox_nv3_3.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ed_txtComboBox_nv3_3.List(numero_item1) = UserForm2.ed_txtComboBox_nv3_3.List(numero_item2) Then
                         UserForm2.ed_txtComboBox_nv3_3.RemoveItem (numero_item2)
                     Else
                     End If
                 End If
             End If
         Next numero_item2
         Next numero_item1
        
        Next verifica_repitidos
    End If
    
End Sub

'Utilizado na guia exclusão "lista dinâmica"
Sub ed_carregarEstruturaPrincipais()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
        Dim verifica_repitidos
        Dim numero_item1
        Dim numero_item2

If UserForm2.ed_txtComboBox_nv1_1.Enabled = True And UserForm2.ed_txtComboBox_nv2_2.Enabled = False Then
        UserForm2.ed_txtComboBox_nv1_1.Clear
        ConectarBanco conexao
        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        
        
'        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        rs.Open "select idNv1 from t_Servicos_Principais_Rendimento order BY idNv1", conexao, 3, 3
        
'        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
'        rs.Open "select idNv1 from t_Servicos_Diversos_Rendimento order BY idNv1", conexao, 3, 3
'        End If
            
        Do Until rs.EOF
        UserForm2.ed_txtComboBox_nv1_1.AddItem rs!idNv1
        rs.MoveNext
        Loop
        conexao.Close
        
'
'        Dim verifica_repitidos
'        Dim numero_item1
'        Dim numero_item2
        ConectarBanco conexao
        For verifica_repitidos = 0 To UserForm2.ed_txtComboBox_nv1_1.ListCount - 1
         sql = "SELECT descricaoNv1 FROM t_Nivel1 WHERE idNv1 = " & UserForm2.ed_txtComboBox_nv1_1.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ed_txtComboBox_nv1_1.List(verifica_repitidos) = rs!descricaoNv1
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         


        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ed_txtComboBox_nv1_1.ListCount - 1
         For numero_item2 = 0 To UserForm2.ed_txtComboBox_nv1_1.ListCount - 1
             If numero_item1 > UserForm2.ed_txtComboBox_nv1_1.ListCount - 1 Or numero_item2 > UserForm2.ed_txtComboBox_nv1_1.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ed_txtComboBox_nv1_1.List(numero_item1) = UserForm2.ed_txtComboBox_nv1_1.List(numero_item2) Then
                         UserForm2.ed_txtComboBox_nv1_1.RemoveItem (numero_item2)
                     Else
                     End If
                 End If
             End If
         Next numero_item2
         Next numero_item1
        
        Next verifica_repitidos
End If

End Sub



Sub ed_verify_estructure()


    
If ed_sPrincipais = True And UserForm2.ed_btnRendimentoBoolean = True Then
    If UserForm2.ed_txtComboBox_nv1_1 <> "" Then
    UserForm2.ed_txtComboBox_nv2_2.Enabled = True
    UserForm2.ed_txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)
            If UserForm2.ed_txtComboBox_nv2_2.Enabled = True And UserForm2.ed_txtComboBox_nv2_2.Value <> "" Then
            UserForm2.ed_txtComboBox_nv3_3.Enabled = True
            UserForm2.ed_txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
            
            If UserForm2.ed_txtComboBox_nv2_2.Enabled = True And UserForm2.ed_txtComboBox_nv2_2.Value <> "" And UserForm2.ed_txtComboBox_nv3_3.Value = "" And UserForm2.ed_txtComboBox_nv3_3.Enabled = True Then
            UserForm2.ed_txtComboBox_nv4_4.Enabled = False
            UserForm2.ed_txtComboBox_nv4_4.BackColor = &H8000000F
            UserForm2.ed_txtComboBox_nv4_4.Clear
            UserForm2.ed_txtBox_rendimento.Enabled = False
            UserForm2.ed_txtBox_rendimento.BackColor = &H8000000F
            UserForm2.ed_txtBox_rendimento = ""
            End If
            
            If UserForm2.ed_txtComboBox_nv2_2.Enabled = True And UserForm2.ed_txtComboBox_nv2_2.Value <> "" And UserForm2.ed_txtComboBox_nv3_3.Value <> "" And UserForm2.ed_txtComboBox_nv4_4.Value = "" Then

            UserForm2.ed_txtBox_rendimento.BackColor = &H8000000F
            UserForm2.ed_txtBox_rendimento.Enabled = False
            UserForm2.ed_txtBox_rendimento = ""
            End If
            
            Else
            UserForm2.ed_txtComboBox_nv3_3.Enabled = False
            UserForm2.ed_txtComboBox_nv3_3.BackColor = &H8000000F
            
            UserForm2.ed_txtComboBox_nv4_4.Enabled = False
            UserForm2.ed_txtComboBox_nv4_4.BackColor = &H8000000F
            End If
    Else
            UserForm2.ed_txtComboBox_nv2_2.Enabled = False
            UserForm2.ed_txtComboBox_nv2_2.BackColor = &H8000000F
    End If
        If UserForm2.ed_txtComboBox_nv1_1 <> "" And UserForm2.ed_txtComboBox_nv2_2 <> "" And UserForm2.ed_txtComboBox_nv3_3 <> "" Then
        UserForm2.ed_txtComboBox_nv4_4.Enabled = True
        'UserForm2.ed_txtComboBox_nv4_4.Clear
        UserForm2.ed_txtComboBox_nv4_4.BackColor = RGB(255, 255, 255)
        End If
ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then

    If ed_ComboBoxNv3.Enabled = True And ed_ComboBoxNv3.Value <> "" Then
        ed_ComboBoxNv4.Enabled = True
        ed_ComboBoxNv4.Clear
        ed_ComboBoxNv4.BackColor = RGB(255, 255, 255)
    Else
        ed_ComboBoxNv4.Enabled = False
        ed_ComboBoxNv4.Clear
        ed_ComboBoxNv4.BackColor = &H8000000F
    End If
    
ElseIf ed_sPrincipais = True And UserForm2.ed_btnCmoPvsBoolean = True Then
    If UserForm2.ed_txtComboBox_nv1_1 <> "" Then
        UserForm2.ed_txtComboBox_nv2_2.Enabled = True
        UserForm2.ed_txtComboBox_nv2_2.BackColor = RGB(255, 255, 255)
            If UserForm2.ed_txtComboBox_nv2_2.Enabled = True And UserForm2.ed_txtComboBox_nv2_2.Value <> "" Then
            UserForm2.ed_txtComboBox_nv3_3.Enabled = True
            UserForm2.ed_txtComboBox_nv3_3.BackColor = RGB(255, 255, 255)
            
            Else
            UserForm2.ed_txtComboBox_nv3_3.Enabled = False
            UserForm2.ed_txtComboBox_nv3_3.BackColor = &H8000000F
            
'            ex_ComboBoxNv4.Enabled = False
'            ex_ComboBoxNv4.BackColor = &H8000000F
            End If
    Else
            UserForm2.ed_txtComboBox_nv2_2.Enabled = False
            UserForm2.ed_txtComboBox_nv2_2.BackColor = &H8000000F
    End If
ElseIf ed_sDiversos = True And UserForm2.ed_btnRendimentoBoolean = True Then
    If UserForm2.ed_txtComboBox_nv3_3.Enabled = True And UserForm2.ed_txtComboBox_nv3_3.Value <> "" Then
        UserForm2.ed_txtComboBox_nv4_4.Enabled = True
        UserForm2.ed_txtComboBox_nv4_4.Clear
        UserForm2.ed_txtComboBox_nv4_4.BackColor = RGB(255, 255, 255)
    Else
        UserForm2.ed_txtComboBox_nv4_4.Enabled = False
        UserForm2.ed_txtComboBox_nv4_4.Clear
        UserForm2.ed_txtComboBox_nv4_4.BackColor = &H8000000F
    End If


End If
    'ed_txtBox_rendimento
End Sub



Sub ed_carregarRendimento()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
        Dim verifica_repitidos
        Dim numero_item1
        Dim numero_item2
Dim idNv3 As Long
Dim idNv2 As Long
Dim idNv1 As Long
Dim id3 As Long
Dim id2 As Long
Dim id1 As Long
Dim idNv4 As Long
Dim id4 As Long


    If UserForm2.ed_txtComboBox_nv1_1.Value <> "" And UserForm2.ed_txtComboBox_nv2_2.Value <> "" And UserForm2.ed_txtComboBox_nv3_3.Value <> "" And UserForm2.ed_txtComboBox_nv4_4.Value <> "" Then
    
    
'ID1
            ConectarBanco conexao

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
        If ed_sPrincipais = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv1 = '" & UserForm2.ed_txtComboBox_nv1_1.Value & "'"
        rs.Open sql, conexao
        On Error GoTo here:
        idNv1 = rs.Fields("idNv1").Value
        id1 = idNv1
        conexao.Close
        ElseIf ed_sDiversos = True Then
        id1 = 7
        conexao.Close
        End If

 
'ID2
            ConectarBanco conexao
        If ed_sPrincipais = True Then
        sql = "select idNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv2 = '" & UserForm2.ed_txtComboBox_nv2_2.Value & "'"
        rs.Open sql, conexao
        On Error GoTo here:
        idNv2 = rs.Fields("idNv2").Value
        id2 = idNv2
        conexao.Close
        ElseIf ed_sDiversos = True Then
        id2 = 0
        conexao.Close
        End If

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
  


'ID3
 
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        sql = "select idNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Value & "'"
        ElseIf ed_sDiversos = True Then
        sql = "select idNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' And descricaoNv3 = '" & UserForm2.ed_txtComboBox_nv3_3.Value & "'"
        End If

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
  
        rs.Open sql, conexao
        On Error GoTo here:
        idNv3 = rs.Fields("idNv3").Value
        id3 = idNv3
        conexao.Close
        
'ID4
     
        ConectarBanco conexao

        sql = "select idInsumo from t_Insumo WHERE And Insumo = '" & UserForm2.ed_txtComboBox_nv4_4.Value & "'"


        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
  
        rs.Open sql, conexao
        On Error GoTo here:
        idNv4 = rs.Fields("idInsumo").Value
        id4 = idNv4
        conexao.Close
        
        ConectarBanco conexao
        If ed_sPrincipais = True Then
        rs.Open "SELECT rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id3 & " AND idNv2 = " & id2 & " AND idNv1 = " & id1 & " AND idNv4 = " & id4 & " ORDER BY idInsumo", conexao, 3, 3
        ElseIf ed_sDiversos = True Then
        rs.Open "SELECT rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id3 & " AND idNv2 = 0 AND idNv1 = 7 AND idNv4 = " & id4 & " ORDER BY idInsumo", conexao, 3, 3
        End If
        
        rs.Open sql, conexao
        On Error GoTo here:
        idNv4 = rs.Fields("idInsumo").Value
        ed_txtBox_rendimento = idNv4
        conexao.Close
End If
here:
End Sub
