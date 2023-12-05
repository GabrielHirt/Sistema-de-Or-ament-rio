Attribute VB_Name = "Exclusão"
Option Explicit

Public GlobalComboBoxValue As String
Public GlobalServiceType As String
Public GlobalTable As String
Public coluna1 As Long
Public coluna2 As Long
Public coluna3 As Long
Public coluna4 As Long


Sub calculatePixelsAmount()
    Dim LargerValue1 As String
    Dim LargerValue2 As String
    Dim LargerValue3 As String
    Dim linha As Integer
    Dim valor As Integer
    
    ' Inicializa as variáveis com zero
    coluna1 = 0
    coluna2 = 0
    coluna3 = 0
    coluna4 = 0
    
    ' Percorre as linhas da ListBox
    For linha = 0 To UserForm2.ex_ListBox.ListCount - 1
        ' Verifica o tamanho do valor da coluna 1
        valor = Len(UserForm2.ex_ListBox.Column(0, linha))
        If valor > coluna1 Then
            coluna1 = valor
            LargerValue1 = UserForm2.ex_ListBox.Column(0, linha)
        End If
        
        If IsNull(Len(UserForm2.ex_ListBox.Column(1, linha))) Then
        GoTo nextLine
        End If

        ' Verifica o tamanho do valor da coluna 2
        valor = Len(UserForm2.ex_ListBox.Column(1, linha))
        If valor > coluna2 Then
            coluna2 = valor
            LargerValue2 = UserForm2.ex_ListBox.Column(1, linha)
        End If
        
        ' Verifica o tamanho do valor da coluna 3
        valor = Len(UserForm2.ex_ListBox.Column(2, linha))
        If valor > coluna3 Then
            coluna3 = valor
            LargerValue3 = UserForm2.ex_ListBox.Column(2, linha)
        End If
        
        If UserForm2.ex_sGeneralInsumo = True Then
            'Verifica o tamanho do valor da coluna 3
            valor = Len(UserForm2.ex_ListBox.Column(3, linha))
            If valor > coluna4 Then
                coluna4 = valor
                LargerValue3 = UserForm2.ex_ListBox.Column(3, linha)
            End If
        End If
nextLine:
    Next linha
    
    ' Exibe os resultados
    'MsgBox "Maior valor de caracteres na coluna 1: " & coluna1 & vbCrLf & _
     '      "Maior valor de caracteres na coluna 2: " & coluna2 & vbCrLf & _
    '       "Maior valor de caracteres na coluna 3: " & coluna3
    
    
    'Pixels Column Size One
    coluna1 = (coluna1 * 50 / 8) + 5
    'Pixels Column Size Two
    coluna2 = (coluna2 * 50 / 8) + 5
    'Pixels Column Size Tree
    coluna3 = (coluna3 * 50 / 8) + 5
    
    If UserForm2.ex_sGeneralInsumo = True Then
    coluna4 = (coluna4 * 50 / 8) + 5
    End If
    
End Sub

Sub SetLayout()

If UserForm2.ex_sPrincipais = True Then
UserForm2.ex_estruturaNv1Title = "Nível 1"
UserForm2.ex_estruturaNv2Title.Visible = True
UserForm2.ex_estruturaNv3Title.Visible = True
UserForm2.ex_ComboBoxNv2.Visible = True
UserForm2.ex_ComboBoxNv3.Visible = True
ElseIf UserForm2.ex_sDiversos = True Then
UserForm2.ex_estruturaNv2Title.Visible = False
UserForm2.ex_estruturaNv3Title.Visible = False
UserForm2.ex_ComboBoxNv2.Visible = False
UserForm2.ex_ComboBoxNv3.Visible = False
UserForm2.ex_estruturaNv1Title = "Nível 3"
ElseIf UserForm2.ex_sTerceiros = True Then
UserForm2.ex_estruturaNv2Title.Visible = False
UserForm2.ex_estruturaNv3Title.Visible = False
UserForm2.ex_estruturaNv1Title = "Nível 3"
UserForm2.ex_ComboBoxNv2.Visible = False
UserForm2.ex_ComboBoxNv3.Visible = False
End If

End Sub

Sub carregarEstruturaPrincipais2()
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

    If UserForm2.ex_ComboBoxNv2.Enabled = True And UserForm2.ex_ComboBoxNv1.Value <> "" And UserForm2.ex_ComboBoxNv3.Enabled = False Then



       UserForm2.ex_ComboBoxNv2.Clear
        ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' and descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1.Value & "'"
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS DIVERSOS' and descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1.Value & "'"
        End If
        
        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
  
        rs.Open sql, conexao
        idNv1 = rs.Fields("idNv1").Value
        id = idNv1
        conexao.Close
    
        ConectarBanco conexao
        
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        rs.Open "SELECT idNv2 FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & " ORDER BY idNv2", conexao, 3, 3
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        rs.Open "SELECT idNv2 FROM t_Servicos_Diversos_Rendimento WHERE idNv1 = " & id & " ORDER BY idNv2", conexao, 3, 3
        End If

        Do Until rs.EOF
        UserForm2.ex_ComboBoxNv2.AddItem rs!idNv2
        rs.MoveNext
        Loop
        conexao.Close
        

'        Dim verifica_repitidos
'        Dim numero_item1
'        Dim numero_item2
        ConectarBanco conexao
        
        
        For verifica_repitidos = 0 To UserForm2.ex_ComboBoxNv2.ListCount - 1
         sql = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & UserForm2.ex_ComboBoxNv2.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ex_ComboBoxNv2.List(verifica_repitidos) = rs!descricaoNv2
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         
         '=====================================================================

        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ex_ComboBoxNv2.ListCount - 1
         For numero_item2 = 0 To UserForm2.ex_ComboBoxNv2.ListCount - 1
             If numero_item1 > UserForm2.ex_ComboBoxNv2.ListCount - 1 Or numero_item2 > UserForm2.ex_ComboBoxNv2.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ex_ComboBoxNv2.List(numero_item1) = UserForm2.ex_ComboBoxNv2.List(numero_item2) Then
                         UserForm2.ex_ComboBoxNv2.RemoveItem (numero_item2)
                     Else
                     End If
                 End If
             End If
         Next numero_item2
         Next numero_item1
        
        Next verifica_repitidos
        
    End If




End Sub



Sub carregarEstruturaPrincipais4()
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



    If UserForm2.ex_ComboBoxNv1.Value <> "" And UserForm2.ex_ComboBoxNv2 <> "" And UserForm2.ex_ComboBoxNv3.Enabled = True Or UserForm2.ex_BtnSelectionDiversosBoolean = True And UserForm2.ex_ComboBoxNv1.Enabled = False And UserForm2.ex_ComboBoxNv2.Enabled = False And UserForm2.ex_ComboBoxNv3.Enabled = True Then
    
    
'ID1
            ConectarBanco conexao

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1.Value & "'"
        rs.Open sql, conexao
        On Error GoTo here:
        idNv1 = rs.Fields("idNv1").Value
        id1 = idNv1
        conexao.Close
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        id1 = 7
        conexao.Close
        End If

 
'ID2
            ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv2 = '" & UserForm2.ex_ComboBoxNv2.Value & "'"
        rs.Open sql, conexao
        On Error GoTo here:
        idNv2 = rs.Fields("idNv2").Value
        id2 = idNv2
        conexao.Close
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        id2 = 0
        conexao.Close
        End If

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
  


'ID3
       UserForm2.ex_ComboBoxNv4.Clear
        ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv3 = '" & UserForm2.ex_ComboBoxNv3.Value & "'"
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        sql = "select idNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' And descricaoNv3 = '" & UserForm2.ex_ComboBoxNv3.Value & "'"
        End If

        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        'rs.Open "select idNv2 from t_Servicos_Principais_Rendimento WHERE idNv1 = " & UserForm2.ex_ComboBoxNv1.Value & "order BY idNv1", conexao, 3, 3
  
        rs.Open sql, conexao
        On Error GoTo here:
        idNv3 = rs.Fields("idNv3").Value
        id3 = idNv3
        conexao.Close
        
        ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        rs.Open "SELECT idInsumo FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id3 & " AND idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idInsumo", conexao, 3, 3
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        rs.Open "SELECT idInsumo FROM t_Servicos_Diversos_Rendimento WHERE idNv3 = " & id3 & " AND idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idInsumo", conexao, 3, 3
        End If

        Do Until rs.EOF
        UserForm2.ex_ComboBoxNv4.AddItem rs!idInsumo
        rs.MoveNext
        Loop
        conexao.Close
        

'        Dim verifica_repitidos
'        Dim numero_item1
'        Dim numero_item2
        ConectarBanco conexao
        For verifica_repitidos = 0 To UserForm2.ex_ComboBoxNv4.ListCount - 1
         sql = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & UserForm2.ex_ComboBoxNv4.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ex_ComboBoxNv4.List(verifica_repitidos) = rs!insumo
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         
         '=====================================================================

        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ex_ComboBoxNv4.ListCount - 1
         For numero_item2 = 0 To UserForm2.ex_ComboBoxNv4.ListCount - 1
             If numero_item1 > UserForm2.ex_ComboBoxNv4.ListCount - 1 Or numero_item2 > UserForm2.ex_ComboBoxNv4.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ex_ComboBoxNv4.List(numero_item1) = UserForm2.ex_ComboBoxNv4.List(numero_item2) Then
                         UserForm2.ex_ComboBoxNv4.RemoveItem (numero_item2)
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
'    If UserForm2.menuEstrutura = False Then
'    MsgBox "Um dos valores da estrutura não está presente no banco de dados!", vbExclamation
'    End If
Exit Sub
Else
MsgBox "carregarEstruturaPrincipais4"
End If

jumpOverIt:
    
End Sub



Sub carregarEstruturaPrincipais3()
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

    If UserForm2.ex_ComboBoxNv1.Value <> "" And UserForm2.ex_ComboBoxNv2 <> "" And UserForm2.ex_ComboBoxNv3.Enabled = True Or UserForm2.ex_BtnSelectionDiversosBoolean = True And UserForm2.ex_ComboBoxNv1.Enabled = False And UserForm2.ex_ComboBoxNv2.Enabled = False Then
    
'ID1




        
        ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1.Value & "'"
        
        rs.Open sql, conexao
        idNv1 = rs.Fields("idNv1").Value
        id1 = idNv1
        conexao.Close
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        id1 = 7
        conexao.Close
        End If
  

    
'ID2
       UserForm2.ex_ComboBoxNv3.Clear
        ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        sql = "select idNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' And descricaoNv2 = '" & UserForm2.ex_ComboBoxNv2.Value & "'"
        rs.Open sql, conexao
        idNv2 = rs.Fields("idNv2").Value
        id2 = idNv2
        conexao.Close
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        id2 = 0
        conexao.Close
        End If
  

        
        ConectarBanco conexao
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        rs.Open "SELECT idNv3 FROM t_Servicos_Principais_Rendimento WHERE idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idNv3", conexao, 3, 3
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        rs.Open "SELECT idNv3 FROM t_Servicos_Diversos_Rendimento WHERE idNv2 = " & id2 & " AND idNv1 = " & id1 & " ORDER BY idNv3", conexao, 3, 3
        End If
        

        Do Until rs.EOF
        UserForm2.ex_ComboBoxNv3.AddItem rs!idNv3
        rs.MoveNext
        Loop
        conexao.Close
        

'        Dim verifica_repitidos
'        Dim numero_item1
'        Dim numero_item2
        ConectarBanco conexao
        For verifica_repitidos = 0 To UserForm2.ex_ComboBoxNv3.ListCount - 1
         sql = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & UserForm2.ex_ComboBoxNv3.List(verifica_repitidos) & ""
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ex_ComboBoxNv3.List(verifica_repitidos) = rs!descricaoNv3
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         
         '=====================================================================

        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ex_ComboBoxNv3.ListCount - 1
         For numero_item2 = 0 To UserForm2.ex_ComboBoxNv3.ListCount - 1
             If numero_item1 > UserForm2.ex_ComboBoxNv3.ListCount - 1 Or numero_item2 > UserForm2.ex_ComboBoxNv3.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ex_ComboBoxNv3.List(numero_item1) = UserForm2.ex_ComboBoxNv3.List(numero_item2) Then
                         UserForm2.ex_ComboBoxNv3.RemoveItem (numero_item2)
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
Sub carregarEstruturaPrincipais()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer
        Dim verifica_repitidos
        Dim numero_item1
        Dim numero_item2

If UserForm2.ex_ComboBoxNv1.Enabled = True And UserForm2.ex_ComboBoxNv2.Enabled = False Then
        UserForm2.ex_ComboBoxNv1.Clear
        ConectarBanco conexao
        'PEGA O ID NV1 DE TODOS AS ESTRUTURAS
        
        
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        rs.Open "select idNv1 from t_Servicos_Principais_Rendimento order BY idNv1", conexao, 3, 3
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        rs.Open "select idNv1 from t_Servicos_Diversos_Rendimento order BY idNv1", conexao, 3, 3
        End If
            
        Do Until rs.EOF
        UserForm2.ex_ComboBoxNv1.AddItem rs!idNv1
        rs.MoveNext
        Loop
        conexao.Close
        
'
'        Dim verifica_repitidos
'        Dim numero_item1
'        Dim numero_item2
        ConectarBanco conexao
        For verifica_repitidos = 0 To UserForm2.ex_ComboBoxNv1.ListCount - 1
         'sql = "SELECT descricaoNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' AND idNv1 = " & UserForm2.ex_ComboBoxNv1.List(verifica_repitidos) & ""
         sql = "SELECT descricaoNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' AND idNv1 = " & UserForm2.ex_ComboBoxNv1.List(verifica_repitidos) & ""

         
         rs.Open sql, conexao, 3, 3
         
         If Not rs.EOF Then
             UserForm2.ex_ComboBoxNv1.List(verifica_repitidos) = rs!descricaoNv1
         End If
         
         rs.Close
        Next verifica_repitidos
         
         conexao.Close
         
         '=====================================================================

        For verifica_repitidos = 0 To 5
        
         For numero_item1 = 0 To UserForm2.ex_ComboBoxNv1.ListCount - 1
         For numero_item2 = 0 To UserForm2.ex_ComboBoxNv1.ListCount - 1
             If numero_item1 > UserForm2.ex_ComboBoxNv1.ListCount - 1 Or numero_item2 > UserForm2.ex_ComboBoxNv1.ListCount - 1 Then
             Exit For
             Else
                 If numero_item1 <> numero_item2 Then
                     If UserForm2.ex_ComboBoxNv1.List(numero_item1) = UserForm2.ex_ComboBoxNv1.List(numero_item2) Then
                         UserForm2.ex_ComboBoxNv1.RemoveItem (numero_item2)
                     Else
                     End If
                 End If
             End If
         Next numero_item2
         Next numero_item1
        
        Next verifica_repitidos
End If









End Sub




Sub ex_CarregarNivelPrincipal()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim i As Integer

'O trecho do código abaixo carrega individualmente as listas para as listas de nível 1,2,3 dentro do botão estrutura
i = 1
If UserForm2.ex_btnEstruturaBoolean = True Then
    If UserForm2.ex_sPrincipais = True Or UserForm2.ex_sGeneralInsumo Then
       SetLayout
  
    While i <= 3
    
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
            If i = 1 Then
            ConectarBanco conexao
            UserForm2.ex_ComboBoxNv1.Clear
            rs.Open "select descricaoNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv1", conexao, 3, 3
            ElseIf i = 2 Then
            UserForm2.ex_ComboBoxNv2.Clear
            ConectarBanco conexao
            rs.Open "select descricaoNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv2", conexao, 3, 3
            ElseIf i = 3 Then
            UserForm2.ex_ComboBoxNv3.Clear
            ConectarBanco conexao
            rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv3", conexao, 3, 3
            End If
        End If
        
        If UserForm2.ex_sGeneralInsumo = True And UserForm2.ex_BtnSelectionDiversosBoolean = True Then
            If i = 1 Then
            ConectarBanco conexao
            UserForm2.ex_ComboBoxNv1.Clear
            rs.Open "select descricaoNv1 from t_Nivel1 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv1", conexao, 3, 3
            ElseIf i = 2 Then
            UserForm2.ex_ComboBoxNv2.Clear
            ConectarBanco conexao
            rs.Open "select descricaoNv2 from t_Nivel2 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv2", conexao, 3, 3
            ElseIf i = 3 Then
            UserForm2.ex_ComboBoxNv3.Clear
            ConectarBanco conexao
            rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv3", conexao, 3, 3
            End If
        End If
    
            
    
        Do Until rs.EOF
        If i = 1 Then

        UserForm2.ex_ComboBoxNv1.AddItem rs!descricaoNv1
        ElseIf i = 2 Then

        UserForm2.ex_ComboBoxNv2.AddItem rs!descricaoNv2
        ElseIf i = 3 Then

        UserForm2.ex_ComboBoxNv3.AddItem rs!descricaoNv3
        End If
        rs.MoveNext
        Loop
        
        conexao.Close
    
        i = i + 1
    Wend
        If UserForm2.ex_sPrincipais = True Or UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        UserForm2.ex_ComboBoxNv2.Enabled = False
        UserForm2.ex_ComboBoxNv2.BackColor = &H8000000F
        UserForm2.ex_ComboBoxNv3.Enabled = False
        UserForm2.ex_ComboBoxNv3.BackColor = &H8000000F
        End If
        
        If UserForm2.ex_sGeneralInsumo = True And UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        UserForm2.ex_ComboBoxNv2.Enabled = False
        UserForm2.ex_ComboBoxNv2.BackColor = &H8000000F
        End If

    
    
    ElseIf UserForm2.ex_sDiversos = True Then
    SetLayout
    ConectarBanco conexao
    UserForm2.ex_ComboBoxNv1.Clear
    rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv3", conexao, 3, 3
      
    UserForm2.ex_ComboBoxNv1.Clear
    Do Until rs.EOF
    UserForm2.ex_ComboBoxNv1.AddItem rs!descricaoNv3

    rs.MoveNext
    Loop

    conexao.Close
    
    ElseIf UserForm2.ex_sTerceiros = True Then
    SetLayout
    ConectarBanco conexao
    UserForm2.ex_ComboBoxNv1.Clear
    rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DE TERCEIROS' order BY idNv3", conexao, 3, 3

    UserForm2.ex_ComboBoxNv1.Clear
    Do Until rs.EOF
    UserForm2.ex_ComboBoxNv1.AddItem rs!descricaoNv3
    rs.MoveNext
    Loop

    conexao.Close

    End If
      GoTo theEnd
End If

'O trecho do código abaixo carrega individual mentes as listas para os botoes de nível 1,2,3 e insumo
ConectarBanco conexao

If UserForm2.ex_btnNv1BooleanP = True Then
rs.Open "select descricaoNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv1", conexao, 3, 3
ElseIf UserForm2.ex_btnNv2BooleanP = True Then
rs.Open "select descricaoNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv2", conexao, 3, 3
ElseIf UserForm2.ex_btnNv3BooleanP = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv3", conexao, 3, 3
ElseIf UserForm2.ex_btnNv4BooleanP = True Then
rs.Open "select Insumo from t_Insumos order BY idInsumo", conexao, 3, 3
End If

UserForm2.ex_ComboBoxNiveis.Clear
Do Until rs.EOF

If UserForm2.ex_btnNv1BooleanP = True Then
UserForm2.ex_ComboBoxNiveis.AddItem rs!descricaoNv1
ElseIf UserForm2.ex_btnNv2BooleanP = True Then
UserForm2.ex_ComboBoxNiveis.AddItem rs!descricaoNv2
ElseIf UserForm2.ex_btnNv3BooleanP = True Then
UserForm2.ex_ComboBoxNiveis.AddItem rs!descricaoNv3
ElseIf UserForm2.ex_btnNv4BooleanP = True Then
UserForm2.ex_ComboBoxNiveis.AddItem rs!insumo
End If

rs.MoveNext
Loop

conexao.Close

theEnd:
   
End Sub
'Carrega listas de servicos diversos e terceiros
Sub ex_CarregarNivelTres()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
'Dim Db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao
If UserForm2.ex_sDiversos = True And UserForm2.ex_btnNv3BooleanD = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' order BY idNv3", conexao, 3, 3
ElseIf UserForm2.ex_sDiversos = True And UserForm2.ex_btnNv4BooleanP = True Then
rs.Open "select Insumo from t_Insumos order BY idInsumo", conexao, 3, 3
ElseIf UserForm2.ex_sTerceiros = True And UserForm2.ex_btnNv3BooleanD = True Then
rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS DE TERCEIROS' order BY idNv3", conexao, 3, 3
End If

UserForm2.ex_ComboBoxNiveis.Clear
If UserForm2.ex_sDiversos = True And UserForm2.ex_btnNv3BooleanD = True Or UserForm2.ex_sTerceiros = True And UserForm2.ex_btnNv3BooleanD = True Then
Do Until rs.EOF
UserForm2.ex_ComboBoxNiveis.AddItem rs!descricaoNv3
rs.MoveNext
Loop
ElseIf UserForm2.ex_sDiversos = True And UserForm2.ex_btnNv4BooleanP = True Then
Do Until rs.EOF
UserForm2.ex_ComboBoxNiveis.AddItem rs!insumo
rs.MoveNext
Loop
End If
'rs.Update
conexao.Close

End Sub



'A função abaixo irá lista em uma listbox os serviços que estão pendentes para que possa ser prosseguido a exclusão do nível
Sub listarPendentes()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim insumo As String
Dim nivel1 As String
Dim nivel2 As String
Dim nivel3 As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim id As String
'Dim rs As Recordset
Dim counte As Integer
Dim db As Object
Dim idNv01 As Long
Dim idNv02 As Long
Dim idNv03 As Long
Dim idInsumo As Long
Dim result As VbMsgBoxResult


Dim teste As String



 
    Dim totalWidth As Double
    Dim columnWidths As String






'Set db = CurrentDb
'Set rs = db.OpenRecordset(sql)

If UserForm2.ex_sPrincipais = True Then
    If UserForm2.ex_ComboBoxNiveis.Value = "" Then
        MsgBox "Selecione um valor, campo vazio!", vbExclamation
        Exit Sub
    End If
    If UserForm2.ex_btnNv1BooleanP = True Or UserForm2.ex_btnNv2BooleanP = True Or UserForm2.ex_btnNv3BooleanP = True Then
    ConectarBanco conexao
        If UserForm2.ex_btnNv1BooleanP = True Then
            sql = "SELECT idNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv1 = '" & UserForm2.ex_ComboBoxNiveis & "';"
        ElseIf UserForm2.ex_btnNv2BooleanP = True Then
            sql = "SELECT idNv2 FROM t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv2 = '" & UserForm2.ex_ComboBoxNiveis & "';"
        ElseIf UserForm2.ex_btnNv3BooleanP = True Then
            sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNiveis & "';"
        End If

        If UserForm2.ex_btnNv1BooleanP = True Then
        rs.Open sql, conexao
        idNv01 = rs.Fields("idNv1").Value
        id = idNv01
        conexao.Close
        ElseIf UserForm2.ex_btnNv2BooleanP = True Then
         rs.Open sql, conexao
        idNv02 = rs.Fields("idNv2").Value
        id = idNv02
        conexao.Close
        ElseIf UserForm2.ex_btnNv3BooleanP = True Then
        rs.Open sql, conexao
        idNv03 = rs.Fields("idNv3").Value
        id = idNv03
        conexao.Close
        End If
    
    ConectarBanco conexao
    
        If UserForm2.ex_btnNv1BooleanP = True Then
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
        ElseIf UserForm2.ex_btnNv2BooleanP = True Then
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv2 = " & id & ";"
        ElseIf UserForm2.ex_btnNv3BooleanP = True Then
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id & ";"
        End If
    
    '    sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
    rs.Open sql, conexao
    counte = rs.Fields("counte").Value
    counte = rs("counte")
    rs.Close
    conexao.Close
    
    If counte = 0 Then

        result = MsgBox("Valor passível de exclusão. Deseja prosseguir com a exclusão do nível selecionado?", vbYesNo + vbQuestion, "Confirmação")
        If result = vbYes Then
'Cria Log
        log_exclusao_nivel
        
            ConectarBanco conexao
            
            If UserForm2.ex_btnNv1BooleanP = True Then
            sql = "DELETE FROM t_Nivel1 Where idNv1 = " & id & ""
            ElseIf UserForm2.ex_btnNv2BooleanP = True Then
            sql = "DELETE FROM t_Nivel2 Where idNv2 = " & id & ""
            ElseIf UserForm2.ex_btnNv3BooleanP = True Then
            sql = "DELETE FROM t_Nivel3 Where idNv3 = " & id & ""
            End If
            
'            sql = "DELETE FROM t_Nivel1 Where idNv1 = " & id & ""
            conexao.Execute sql
            conexao.Close
            UserForm2.ex_ListBox.Clear
            MsgBox "Registro excluído com sucesso!", vbInformation
            Exclusão.ex_CarregarNivelPrincipal
            Exit Sub

        ElseIf result = vbNo Then
        MsgBox "Você optou por não realizar a operação de exclusão."
        Exit Sub
        End If
        
    End If
        
     ConectarBanco conexao
    ' Monta o código SQL para buscar as linhas com o idNv1 selecionado
        If UserForm2.ex_btnNv1BooleanP = True Then
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
        ElseIf UserForm2.ex_btnNv2BooleanP = True Then
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv2 = " & id & ";"
        ElseIf UserForm2.ex_btnNv3BooleanP = True Then
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id & ";"
        End If
'    sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
    rs.Open sql, conexao
    rs.MoveFirst

    UserForm2.ex_ListBox.Clear
'    UserForm2.ex_ListBox.AddItem "         Níve 1                    Nível 2                           Serviço                                                   Insumo"
        
    
    If counte > 0 Then
            UserForm2.ex_ListBox.ColumnCount = 3
            UserForm2.ex_ListBox.AddItem ""
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"


    ' Ajusta o tamanho das colunas
    UserForm2.ex_ListBox.columnWidths = "100;100;100;100" ' Ajuste o tamanho conforme necessário
            
        Do Until rs.EOF
'            UserForm2.ex_ListBox.AddItem rs("idNv1") & " | " & rs("idNv2") & " | " & rs("idNv3") & " | " & rs("idInsumo")
'            rs.MoveNext
            
            
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
        
                  
        UserForm2.ex_ListBox.ColumnCount = 3
        ' Verifica se há um valor correspondente na tabela t_Insumos
        If Not rsDescricaoNv1.EOF Then
            ' Obtém o valor de Insumo encontrado na tabela t_Insumos
            'Dim insumo As String
            'insumo = rsPInsumo("Insumo")
            nivel3 = rsDescricaoNv3P("DescricaoNv3")
            nivel2 = rsDescricaoNv2("DescricaoNv2")
            nivel1 = rsDescricaoNv1("DescricaoNv1")
' Adiciona uma nova linha ao ListBox
    UserForm2.ex_ListBox.AddItem ""
    ' Define os valores das colunas para a nova linha
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = nivel1
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = nivel2
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = nivel3



    ' Ajusta o tamanho das colunas
    '50 pixels = 8 caracteres
    'teste = "50"
    UserForm2.ex_ListBox.columnWidths = "300;300;800" ' Ajuste o tamanho conforme necessário"
    'UserForm2.ex_ListBox.columnWidths = teste & ";300;800"  ' Ajuste o tamanho conforme necessário"
    
    

        End If
        rsPInsumo.Close
        rs.MoveNext

        Loop
    End If
    rs.Close
    
    
    
    '====================================================================
    
    Dim Num_Linhas, i, X As Integer
    
    Num_Linhas = UserForm2.ex_ListBox.ListCount - 1
    
    For i = 0 To Num_Linhas
    For X = Num_Linhas To (i + 1) Step -1
        If UserForm2.ex_ListBox.List(i, 0) = UserForm2.ex_ListBox.List(X, 0) _
            And UserForm2.ex_ListBox.List(i, 1) = UserForm2.ex_ListBox.List(X, 1) _
            And UserForm2.ex_ListBox.List(i, 2) = UserForm2.ex_ListBox.List(X, 2) Then
            UserForm2.ex_ListBox.RemoveItem X
            Num_Linhas = Num_Linhas - 1
        End If
    Next X
Next i
    
    
    '====================================================================
    
    
8
    '============= AJUSTAR COLUNAS =========================================
    
    calculatePixelsAmount
    
    
    UserForm2.ex_ListBox.columnWidths = coluna1 & ";" & coluna2 & ";" & coluna3 ' Ajuste o tamanho conforme necessário"
    
    
    
    '=========================================================================


    SaveLogList
    MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & Num_Linhas & " serviço(s) registrados.", vbExclamation
    End If

    
    

ElseIf UserForm2.ex_btnNv2BooleanP = True Then
'
ElseIf UserForm2.ex_btnNv3BooleanP = True Then
'
ElseIf UserForm2.ex_btnNv4BooleanP = True And UserForm2.ex_sGeneralInsumo = True Then


    If UserForm2.ex_ComboBoxNiveis.Value = "" Then
        MsgBox "Selecione um valor, campo vazio!", vbExclamation
        Exit Sub
    End If
    If UserForm2.ex_btnNv1BooleanP = True Or UserForm2.ex_btnNv2BooleanP = True Or UserForm2.ex_btnNv3BooleanP = True Then

        ConectarBanco conexao
        sql = "SELECT idInsumo FROM t_Insumo WHERE Insumo = '" & UserForm2.ex_ComboBoxNiveis & "';"
    
    
        rs.Open sql, conexao
        idInsumo = rs.Fields("idInsumo").Value
        id = idInsumo
        conexao.Close
    
    
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idInsumo = " & id & ";"
        rs.Open sql, conexao
        counte = rs.Fields("counte").Value
        counte = rs("counte")
        rs.Close
        conexao.Close
        
        Dim counte1 As Long
        
        ConectarBanco conexao
        sql = "SELECT COUNT(*) AS counte1 FROM t_Servicos_Principais_Diversos WHERE idInsumo = " & id & ";"
        rs.Open sql, conexao
        counte1 = rs.Fields("counte1").Value
        counte1 = rs("counte1")
        rs.Close
        conexao.Close
        
        counte = counte + counte1
    
        If counte = 0 Then
    
            result = MsgBox("Valor passível de exclusão. Deseja prosseguir com a exclusão do nível selecionado?", vbYesNo + vbQuestion, "Confirmação")
            If result = vbYes Then

'Cria Log
        log_exclusao_nivel
        
                ConectarBanco conexao
                sql = "DELETE FROM t_Insumo Where idInsumo = " & id & ""
                conexao.Execute sql
                conexao.Close
                
                UserForm2.ex_ListBox.Clear
                MsgBox "Registro excluído com sucesso!", vbInformation
                Exclusão.ex_CarregarNivelPrincipal
                Exit Sub
    
            ElseIf result = vbNo Then
            MsgBox "Você optou por não realizar a operação de exclusão."
            Exit Sub
            End If
            
        End If
        
'---------------------- LISTAR VALORES -------------------------------

        ConectarBanco conexao
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idInsumo = " & id & ";"

        rs.Open sql, conexao
        rs.MoveFirst
        UserForm2.ex_ListBox.Clear
            
    
        If counte > 0 Then
            UserForm2.ex_ListBox.ColumnCount = 3
            UserForm2.ex_ListBox.AddItem ""
            UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
            UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
            UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"
 
            
            ' Ajusta o tamanho das colunas
            UserForm2.ex_ListBox.columnWidths = "100;100;100;100" ' Ajuste o tamanho conforme necessário
                
            Do Until rs.EOF
                
                Dim strNv001 As Integer
                strNv001 = rs("idNv1")
                ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
                Dim nive001SQL As String
                nive001SQL = "SELECT descricaoNv1 FROM t_Nivel1 WHERE idNv1 = " & strNv001 & ";"
                Dim rsDescricaoNv001 As ADODB.Recordset
                Set rsDescricaoNv001 = New ADODB.Recordset
                rsDescricaoNv001.Open nive1SQL, conexao
                
            
                Dim strNv002 As Integer
                strNv002 = rs("idNv2")
                ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
                Dim nivel002SQL As String
                nivel002SQL = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & strNv002 & ";"
                Dim rsDescricaoNv002 As ADODB.Recordset
                Set rsDescricaoNv002 = New ADODB.Recordset
                rsDescricaoNv002.Open nivel2SQL, conexao
                
                Dim strNv003P As Integer
                strNv003P = rs("idNv3")
                ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
                Dim nivel003SQLP As String
                nivel003SQLP = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & strNv003P & ";"
                Dim rsDescricaoNv003P As ADODB.Recordset
                Set rsDescricaoNv003P = New ADODB.Recordset
                rsDescricaoNv003P.Open nivel3SQLP, conexao
                       
            
                Dim strInsumo00P As Integer
                strInsumo00P = rs("idInsumo")
                ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
                Dim insumo00SQLP As String
                insumo00SQLP = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & strInsumo00P & ";"
                Dim rsPInsumo00 As ADODB.Recordset
                Set rsPInsumo00 = New ADODB.Recordset
                rsPInsumo00.Open insumoSQLP, conexao
                
                          
                UserForm2.ex_ListBox.ColumnCount = 4
                ' Verifica se há um valor correspondente na tabela t_Insumos
                If Not rsDescricaoNv001.EOF Then
                    nivel3 = rsDescricaoNv003P("DescricaoNv3")
                    nivel2 = rsDescricaoNv002("DescricaoNv2")
                    nivel1 = rsDescricaoNv001("DescricaoNv1")

        
                    
                    ' Adiciona uma nova linha ao ListBox
                     UserForm2.ex_ListBox.AddItem ""
            
                    ' Define os valores das colunas para a nova linha
                    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = nivel1
                    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = nivel2
                    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = nivel3

                    'UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = insumo
            
                    ' Ajusta o tamanho das colunas
                    UserForm2.ex_ListBox.columnWidths = "100;300;300" ' Ajuste o tamanho conforme necessário
            
            
                End If
                    '        rsDescricaoNv3P.Close
                    '        rsDescricaoNv2.Close
                    rsPInsumo.Close
                    rs.MoveNext
                
            Loop
        End If
    rs.Close

    MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & counte & " serviço(s) registrados.", vbExclamation
    End If





ElseIf UserForm2.ex_sDiversos = True Then

    If UserForm2.ex_btnNv3BooleanD = True Then
        If UserForm2.ex_ComboBoxNiveis = "" Then
        MsgBox "Selecione um valor presente na lista, campo vazio!", vbExclamation
        End If

    ConectarBanco conexao
    sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNiveis & "';"
    rs.Open sql, conexao
    Dim idNv3 As Long
    idNv3 = rs.Fields("idNv3").Value
    
    id = idNv3
    conexao.Close
  
'    ' Verifica se o valor selecionado no combobox está presente na tabela
'    If DCount("idNv1", "t_Servicos_Principais_Rendimento", "idNv1 = '" & id & "'") = 0 Then
'        MsgBox "Valor não encontrado na tabela t_Servicos_Principais_Rendimento.", vbInformation, "Aviso"
'        Exit Sub
'    End If

    ' Verifica se o valor selecionado no combobox está presente na tabela
    ConectarBanco conexao
    
    sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Diversos_Rendimento WHERE idNv3 = " & id & ";"
    'sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv1 = '1';"
'        sql = "SELECT COUNT(*) AS count FROM t_Servicos_Principais_Rendimento WHERE idNv1 = 1;"
'    Set rs = CurrentDb.OpenRecordset(sql)
    '     sql = "SELECT COUNT(*) FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
    rs.Open sql, conexao
    counte = rs.Fields("counte").Value
    
    
    
    counte = rs("counte")
    rs.Close
    conexao.Close
    
    If counte = 0 Then
        MsgBox "Valor passível de exclusão. Deseja prosseguir com a exclusão do nível selecionado?", vbInformation, "Aviso"
        Exit Sub
    End If
    
     
    
    
    ConectarBanco conexao
' Monta o código SQL para buscar as linhas com o idNv1 selecionado
sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Diversos_Rendimento WHERE idNv3 = " & id & ";"

' Executa o código SQL e adiciona os resultados na listbox
rs.Open sql, conexao
rs.MoveFirst

UserForm2.ex_ListBox.Clear
UserForm2.ex_ListBox.AddItem "                         Serviço                                                   Insumo"
    If counte > 0 Then
        Do Until rs.EOF
        
        
           Dim strInsumo As Integer
           strInsumo = rs("idInsumo")
            ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
            Dim insumoSQL As String
            insumoSQL = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & strInsumo & ";"
            Dim rsInsumo As ADODB.Recordset
            Set rsInsumo = New ADODB.Recordset
            rsInsumo.Open insumoSQL, conexao
                    
           Dim strNv3 As Integer
           strNv3 = rs("idNv3")
            ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
            Dim nivel3SQL As String
            nivel3SQL = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & strNv3 & ";"
            Dim rsDescricaoNv3 As ADODB.Recordset
            Set rsDescricaoNv3 = New ADODB.Recordset
            rsDescricaoNv3.Open nivel3SQL, conexao
                 
            ' Verifica se há um valor correspondente na tabela t_Insumos
            If Not rsInsumo.EOF Then
                ' Obtém o valor de Insumo encontrado na tabela t_Insumos
    '            Dim insumo As String
                insumo = rsInsumo("Insumo")
                nivel3 = rsDescricaoNv3("DescricaoNv3")
                ' Adiciona o item ao ListBox substituindo o valor de idNv3 pelo valor de Insumo
                UserForm2.ex_ListBox.AddItem nivel3 & "     -     " & insumo '& " | " & rs("rendimento")
            End If
            
            rsInsumo.Close
            rs.MoveNext
        Loop
    End If
rs.Close
    

    MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & counte & " serviço(s) registrados.", vbExclamation
    End If

    
ElseIf UserForm2.ex_sTerceiros = True Then

    If UserForm2.ex_btnNv3BooleanD = True Then
        If UserForm2.ex_ComboBoxNiveis = "" Then
        MsgBox "Selecione um valor presente na lista, campo vazio!", vbExclamation
        End If
        ConectarBanco conexao
        sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS DE TERCEIROS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNiveis & "';"
        rs.Open sql, conexao
        'ERRO se o registro não existir (TRATAR)
        idNv3 = rs.Fields("idNv3").Value
        id = idNv3
        conexao.Close
        
        ConectarBanco conexao
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Terceiros WHERE idNv3 = " & id & ";"
        rs.Open sql, conexao
        counte = rs.Fields("counte").Value
        
        counte = rs("counte")
        rs.Close
        conexao.Close
        
        If counte = 0 Then
            result = MsgBox("Nenhu resigstro encontrado. Deseja excluir o registor encontrado?", vbYesNo + vbQuestion, "Confirmação")
            If result = vbYes Then
'Cria Log
        log_exclusao_nivel
        
                ConectarBanco conexao
                sql = "DELETE FROM t_Nivel3 Where idNv3 = " & id & ""
                conexao.Execute sql
                conexao.Close
'               Set rs = Nothing
'               Set conexao = Nothing
                MsgBox "Registro excluído com sucesso", vbInformation
                ex_CarregarNivelTres
            End If
            
            Exit Sub
        End If
        
         ConectarBanco conexao
        ' Monta o código SQL para buscar as linhas com o idNv1 selecionado
        sql = "SELECT descricaoNv3, idNv1, idNv2, idNv3 FROM t_Servicos_Terceiros WHERE idNv3 = " & id & ";"
    
        rs.Open sql, conexao
    
        rs.MoveFirst
        
        UserForm2.ex_ListBox.Clear
        If counte > 0 Then
            UserForm2.ex_ListBox.AddItem "Descrição de Serviço"
            Do Until rs.EOF
                UserForm2.ex_ListBox.AddItem rs("descricaoNv3") '& " | " & rs("idNv2") & " | " & rs("idNv3") & " | " & rs("idNv1")
                rs.MoveNext
            Loop
        End If
        rs.Close
    
        MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & counte & " serviço(s) registrados.", vbExclamation
        
        
        GoTo juuump
        
        
        ConectarBanco conexao
        sql = "SELECT * FROM t_Servicos_Terceiros Where idNv3 = " & id & ""
        rs.Open sql, conexao
        If Not rs.EOF Then
            result = MsgBox("Deseja excluir o registor encontrado?", vbYesNo + vbQuestion, "Confirmação")
            If result = vbYes Then
'Cria Log
        log_exclusao_nivel
        
                sql = "DELETE FROM t_Servicos_Terceiros Where idNv3 = " & id & ""
                conexao.Execute sql
                
                rs.Close
                conexao.Close
                
'                Set rs = Nothing
'                Set conexao = Nothing
                    MsgBox "Registro excluído com sucesso", vbInformation

            End If
        Else
        MsgBox "Nenhum valor encontrado para o serviço selecionado.", vbInformation
        End If
ConectarBanco conexao
sql = "UPDATE t_Servicos_Terceiros SET Status='Inativo' WHERE idNv3 = " & id & ""
conexao.Execute sql

conexao.Close

Set conexao = Nothing

                    ex_CarregarNivelTres
    End If

End If

juuump:

End Sub


Sub ex_deleteEstrutura()
Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim id As String
Dim counte As Integer
Dim db As Object
Dim result As VbMsgBoxResult
Dim id1 As Long
Dim id2 As Long
Dim id3 As Long
Dim id4 As Long
Dim idNv1 As Long
Dim idNv2 As Long
Dim idNv3 As Long
Dim idNv4 As Long
Dim idInsumo  As String


If UserForm2.ex_sPrincipais = True Then
    If UserForm2.ex_btnEstruturaBoolean = True Then
        If UserForm2.ex_ComboBoxNv1 = "" Or UserForm2.ex_ComboBoxNv2 = "" Or UserForm2.ex_ComboBoxNv3 = "" Then
        MsgBox "Existe um ou mais campos vazios, selecione um valor!", vbExclamation
        Exit Sub
        End If
    ConectarBanco conexao
    sql = "SELECT idNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1 & "';"
    rs.Open sql, conexao
    On Error GoTo there
    idNv1 = rs.Fields("idNv1").Value
    id1 = idNv1
    conexao.Close
    
        ConectarBanco conexao
    sql = "SELECT idNv2 FROM t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv2 = '" & UserForm2.ex_ComboBoxNv2 & "';"
    rs.Open sql, conexao
    On Error GoTo there
    idNv2 = rs.Fields("idNv2").Value
    id2 = idNv2
    conexao.Close
    
    ConectarBanco conexao
    sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNv3 & "';"
    rs.Open sql, conexao
    On Error GoTo there
    idNv3 = rs.Fields("idNv3").Value
    id3 = idNv3
    conexao.Close
        
    result = MsgBox("Deseja excluir a estrutura selecionada?", vbYesNo + vbQuestion, "Confirmação")
        If result = vbYes Then
        kLogs.log_exclusao_estrutura
'Excluir de t_Servicos_Principais
            ConectarBanco conexao
            sql = "DELETE FROM t_Servicos_Principais Where idNv1 = " & id1 & " AND idNv2 = " & id2 & " AND idNv3 = " & id3 & ""
            conexao.Execute sql
            conexao.Close

'Excluir de t_Servicos_Principais_Rendimento
            ConectarBanco conexao
            sql = "DELETE FROM t_Servicos_Principais_Rendimento Where idNv1 = " & id1 & " AND idNv2 = " & id2 & " AND idNv3 = " & id3 & ""
            conexao.Execute sql
            conexao.Close

            
            
            MsgBox "Estrutura excluída com sucesso!", vbInformation
            
'SE LISTA VAZIA, PULAR TRECHO QUE ATUALIZA LISTA (A LISTA ESTARÁ VAZIA QUANDO O USUÁRIO EXCLUIR SEM UTILIZAR A VISUALIZAÇÃO DA LISATA)
            If UserForm2.ex_ListBox.ListCount = 0 Then
            
            ex_CarregarNivelPrincipal
            GoTo noList

            End If
            
            
            ChargeListByLog
            ex_CarregarNivelPrincipal
            calculatePixelsAmount
         UserForm2.ex_ListBox.columnWidths = coluna1 & ";" & coluna2 & ";" & coluna3 & ";" & coluna4 ' Ajuste o tamanho conforme necessário"
        End If
    End If
    GoTo jumpOverIt
there:

    MsgBox "O valor de um dos níveis não foi encontrado.", vbExclamation
jumpOverIt:
noList:

    
    
    
    
    
    

    

ElseIf UserForm2.ex_sDiversos = True Then
    If UserForm2.ex_btnNv3BooleanD = True Then
        If UserForm2.ex_ComboBoxNiveis = "" Then
        MsgBox "Selecione um valor presente na lista, campo vazio!", vbExclamation
        Exit Sub
        End If
    ConectarBanco conexao
    sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNiveis & "';"
    rs.Open sql, conexao
    On Error GoTo here
    idNv3 = rs.Fields("idNv3").Value
    id = idNv3
    conexao.Close
        
    result = MsgBox("Deseja excluir o serviço selecionado?", vbYesNo + vbQuestion, "Confirmação")
        If result = vbYes Then
'Cria Log
        kLogs.log_exclusao_estrutura
        kLogs.log_exclusao_nivel
'Excluir de t_Servicos_Diversos
            ConectarBanco conexao
            sql = "DELETE FROM t_Servicos_Diversos Where idNv3 = " & id & ""
            conexao.Execute sql
            conexao.Close

'Excluir de t_Servicos_Diversos_Rendimento
            ConectarBanco conexao
            sql = "DELETE FROM t_Servicos_Diversos_Rendimento Where idNv3 = " & id & ""
            conexao.Execute sql
            conexao.Close

'Excluir de t_Nivel2
            ConectarBanco conexao
            sql = "DELETE FROM t_Nivel3 Where idNv3 = " & id & ""
            conexao.Execute sql
            conexao.Close
            UserForm2.ex_ListBox.Clear
            MsgBox "Registro excluído com sucesso!", vbInformation

'Carrega o Nível 2 e 3 apesar do nome
            ex_CarregarNivelTres
            calculatePixelsAmount
            UserForm2.ex_ListBox.columnWidths = coluna1 & ";" & coluna2 & ";" & coluna3 & ";" & coluna4 ' Ajuste o tamanho conforme necessário"
        End If
    End If



ElseIf UserForm2.ex_sTerceiros = True Then
    If UserForm2.ex_btnNv3BooleanD = True Then
        If UserForm2.ex_ComboBoxNiveis = "" Then
        MsgBox "Selecione um valor presente na lista, campo vazio!", vbExclamation
        Exit Sub
        End If
    ConectarBanco conexao
    sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS DE TERCEIROS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNiveis & "';"
    rs.Open sql, conexao
    On Error GoTo here
    idNv3 = rs.Fields("idNv3").Value
    id = idNv3
    conexao.Close
        
    result = MsgBox("Deseja excluir o serviço selecionado?", vbYesNo + vbQuestion, "Confirmação")
        If result = vbYes Then
'Cria Log
        kLogs.log_exclusao_nivel
'Excluir de t_Servicos Terceiros
            ConectarBanco conexao
            sql = "DELETE FROM t_Servicos_Terceiros Where idNv3 = " & id & ""
            conexao.Execute sql
            conexao.Close

'Excluir de t_Nivel3
            ConectarBanco conexao
            sql = "DELETE FROM t_Nivel3 Where idNv3 = " & id & ""
            conexao.Execute sql
            conexao.Close
            UserForm2.ex_ListBox.Clear
            MsgBox "Registro excluído com sucesso!", vbInformation

            ex_CarregarNivelTres
            calculatePixelsAmount
            UserForm2.ex_ListBox.columnWidths = coluna1 & ";" & coluna2 & ";" & coluna3 & ";" & coluna4 ' Ajuste o tamanho conforme necessário"
        End If
    End If
GoTo jumpit
here:
    MsgBox "Valor selecionado não encontrado no banco de dados!", vbExclamation
    
jumpit:

    ElseIf UserForm2.ex_sGeneralInsumo Then
    
        If UserForm2.ex_BtnSelectionPrincipaisBoolean = True Then
        
        '======= ADICIONAR UMA VALIDAÇÃO ===========
        '
        'Apenas passará se todos os campos estiverem <> de vazio
        If UserForm2.ex_ComboBoxNv1 = "" Or UserForm2.ex_ComboBoxNv2 = "" Or UserForm2.ex_ComboBoxNv3 = "" Or UserForm2.ex_ComboBoxNv4 = "" Then
        MsgBox "Existe um ou mais campos vazios!", vbExclamation
        Exit Sub
        End If
        
        '==========================================
        
            If UserForm2.ex_btnEstruturaBoolean = True Then
                If UserForm2.ex_ComboBoxNv1 = "" Or UserForm2.ex_ComboBoxNv2 = "" Or UserForm2.ex_ComboBoxNv3 = "" Or UserForm2.ex_ComboBoxNv4 = "" Then
                MsgBox "Existe um ou mais campos vazios, selecione um valor!", vbExclamation
                Exit Sub
                End If
            ConectarBanco conexao
            sql = "SELECT idNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1 & "';"
            rs.Open sql, conexao
            On Error GoTo there
            idNv1 = rs.Fields("idNv1").Value
            id1 = idNv1
            conexao.Close
        
                ConectarBanco conexao
            sql = "SELECT idNv2 FROM t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv2 = '" & UserForm2.ex_ComboBoxNv2 & "';"
            rs.Open sql, conexao
            On Error GoTo there
            idNv2 = rs.Fields("idNv2").Value
            id2 = idNv2
            conexao.Close
        
            ConectarBanco conexao
            sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNv3 & "';"
            rs.Open sql, conexao
            On Error GoTo there
            idNv3 = rs.Fields("idNv3").Value
            id3 = idNv3
            conexao.Close
            
            
            '== NOVO === 13/09/2023
            
            ConectarBanco conexao
            sql = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ex_ComboBoxNv4 & "';"
            rs.Open sql, conexao
            On Error GoTo there
            idInsumo = rs.Fields("idInsumo").Value
            id4 = idInsumo
            conexao.Close
            
            '======================
        
            result = MsgBox("Deseja excluir o insumo da estrutura selecionada?", vbYesNo + vbQuestion, "Confirmação")
                If result = vbYes Then
'Criar Log
                log_exclusao_estrutura
                
'        'Excluir de t_Servicos_Principais
'                    ConectarBanco conexao
'                    sql = "DELETE FROM t_Servicos_Principais Where idNv1 = " & id1 & " AND idNv2 = " & id2 & " AND idNv3 = " & id3 & ""
'                    conexao.Execute sql
'                    conexao.Close
        
        'Excluir de t_Servicos_Principais_Rendimento
                    ConectarBanco conexao
                    sql = "DELETE FROM t_Servicos_Principais_Rendimento Where idNv1 = " & id1 & " AND idNv2 = " & id2 & " AND idNv3 = " & id3 & " AND idInsumo = " & id4 & ""
                    conexao.Execute sql
                    conexao.Close
        
        
                    MsgBox "Estrutura excluída com sucesso!", vbInformation
        Dim rows As Long
        
        rows = UserForm2.ex_ListBox.ListCount - 1
                    If Exclusão.GlobalTable <> "" And rows <> -1 Then
                    Exclusão.ChargeListByLogInsumo
                    End If
                    
                    ex_CarregarNivelPrincipal
                    calculatePixelsAmount
                    UserForm2.ex_ListBox.columnWidths = coluna1 & ";" & coluna2 & ";" & coluna3 & ";" & coluna4 ' Ajuste o tamanho conforme necessário"
                UserForm2.ex_ComboBoxNv4.Enabled = False
                UserForm2.ex_ComboBoxNv4.Clear
                UserForm2.ex_ComboBoxNv4.BackColor = &H8000000F
                End If
            End If
        
        
        
        
        
        ElseIf UserForm2.ex_BtnSelectionDiversosBoolean = True Then
        
         If UserForm2.ex_btnEstruturaBoolean = True Then
                ''If UserForm2.ex_ComboBoxNv1 = "" OR UserForm2.ex_ComboBoxNv3 = "" Then
                If UserForm2.ex_ComboBoxNv1.Enabled = True And UserForm2.ex_ComboBoxNv2.Enabled = True And UserForm2.ex_ComboBoxNv2.Enabled = False And UserForm2.ex_ComboBoxNv1.Enabled = False Then
                MsgBox "Existe um ou mais campos vazios, selecione um valor!", vbExclamation
                Exit Sub
                End If
                
'            ConectarBanco conexao
'            sql = "SELECT idNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS DIVERSOS' AND descricaoNv1 = '" & UserForm2.ex_ComboBoxNv1 & "';"
'            rs.Open sql, conexao
'            On Error GoTo there
'            idNv1 = rs.Fields("idNv1").Value
'            id1 = idNv1
'            conexao.Close
'
'                ConectarBanco conexao
'            sql = "SELECT idNv2 FROM t_Nivel2 WHERE grupo = 'SERVICOS DIVERSOS' AND descricaoNv2 = '" & UserForm2.ex_ComboBoxNv2 & "';"
'            rs.Open sql, conexao
'            On Error GoTo there
'            idNv2 = rs.Fields("idNv2").Value
'            id2 = idNv2
'            conexao.Close
        
            ConectarBanco conexao
            sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS DIVERSOS' AND descricaoNv3 = '" & UserForm2.ex_ComboBoxNv3 & "';"
            rs.Open sql, conexao
            On Error GoTo there
            idNv3 = rs.Fields("idNv3").Value
            id3 = idNv3
            conexao.Close
            
            ConectarBanco conexao
            sql = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ex_ComboBoxNv4 & "';"
            rs.Open sql, conexao
            On Error GoTo there
            idNv4 = rs.Fields("idInsumo").Value
            id4 = idNv4
            conexao.Close
        
            result = MsgBox("Deseja excluir a estrutura selecionada?", vbYesNo + vbQuestion, "Confirmação")
                If result = vbYes Then
'Criar Log
                log_exclusao_estrutura
'        'Excluir de t_Servicos_Principais
'                    ConectarBanco conexao
'                    sql = "DELETE FROM t_Servicos_Diversos Where idNv1 = 7 AND idNv2 = 0 AND idNv3 = " & id3 & ""
'                    conexao.Execute sql
'                    conexao.Close
        
        'Excluir de t_Servicos_Principais_Rendimento
                    ConectarBanco conexao
                    sql = "DELETE FROM t_Servicos_Diversos_Rendimento Where idNv1 = 7 AND idNv2 = 0 AND idNv3 = " & id3 & " AND idInsumo = " & id4 & ""
                    conexao.Execute sql
                    conexao.Close
        
        
                    MsgBox "Estrutura excluída com sucesso!", vbInformation
                    Dim rows1 As Long
                    rows1 = UserForm2.ex_ListBox.ListCount - 1
                    If Exclusão.GlobalTable <> "" And rows1 <> -1 Then
                    Exclusão.ChargeListByLogInsumo
                    End If
        
                    UserForm2.ex_ComboBoxNv1.Enabled = False
                    UserForm2.ex_ComboBoxNv1.BackColor = &H8000000F
                    UserForm2.ex_ComboBoxNv2.Enabled = False
                    UserForm2.ex_ComboBoxNv2.BackColor = &H8000000F
                    UserForm2.ex_ComboBoxNv3.Enabled = True
                    UserForm2.ex_ComboBoxNv3.BackColor = RGB(255, 255, 255)
                    UserForm2.ex_ComboBoxNv4.Enabled = False
                    UserForm2.ex_ComboBoxNv4.BackColor = &H8000000F
                    
                    
                    'Limpar ComboBoxes
                    UserForm2.ex_ComboBoxNv1.Clear
                    UserForm2.ex_ComboBoxNv2.Clear
                    UserForm2.ex_ComboBoxNv3.Clear
                    UserForm2.ex_ComboBoxNv4.Clear
                    'Carregar função de lista
                    'Exclusão.ex_CarregarNivelPrincipal
                    UserForm2.ex_BtnSelectionPrincipais.BackColor = &H8000000F
                    UserForm2.ex_BtnSelectionPrincipais.Font.Bold = False
                    UserForm2.ex_BtnSelectionDiversos.BackColor = RGB(255, 230, 153)
                    UserForm2.ex_BtnSelectionDiversos.Font.Bold = True
                    carregarEstruturaPrincipais3
                End If
            End If
        
        End If
    
    End If

End Sub

Sub verificarPendenteInsumo()

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


UserForm2.ex_ListBox.Clear
If UserForm2.ex_btnNv4BooleanP = True And UserForm2.ex_sGeneralInsumo = True Then

    If UserForm2.ex_ComboBoxNiveis.Value = "" Then
        MsgBox "Selecione um valor, campo vazio!", vbExclamation
        Exit Sub
    End If
    If UserForm2.ex_btnNv4BooleanP = True And UserForm2.ex_sGeneralInsumo = True Then

    ConectarBanco conexao
    sql = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & UserForm2.ex_ComboBoxNiveis & "';"


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
    
    Dim counte1 As Long
    
    ConectarBanco conexao
    sql = "SELECT COUNT(*) AS counte1 FROM t_Servicos_Diversos_Rendimento WHERE idInsumo = " & id & ";"
    rs.Open sql, conexao
    counte1 = rs.Fields("counte1").Value
    counte1 = rs("counte1")
    rs.Close
    conexao.Close
    Dim counter As Integer
    counter = counte + counte1
    
    If counter = 0 Then

        result = MsgBox("Valor passível de exclusão. Deseja prosseguir com a exclusão do nível selecionado?", vbYesNo + vbQuestion, "Confirmação")
        If result = vbYes Then
'Criar Log
kLogs.log_exclusao_nivel

            ConectarBanco conexao
            sql = "DELETE FROM t_Insumos Where idInsumo = " & id & ""
            conexao.Execute sql
            conexao.Close
            
            UserForm2.ex_ListBox.Clear
            MsgBox "Registro excluído com sucesso!", vbInformation
            Exclusão.ex_CarregarNivelPrincipal
            Exit Sub

        Else
        MsgBox "Você optou por não realizar a operação de exclusão."
        Exit Sub
        End If
        
    End If
        
'---------------------- LISTAR VALORES -------------------------------
        UserForm2.ex_ListBox.Clear
        UserForm2.ex_ListBox.ColumnCount = 4
        UserForm2.ex_ListBox.AddItem ""
        'UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Serviços Principais"

'        UserForm2.ex_ListBox.ColumnCount = 3
        'UserForm2.ex_ListBox.AddItem ""
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
'        UserForm2.ex_ListBox.ColumnCount = 4
'        UserForm2.ex_ListBox.AddItem ""
'        'UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Serviços Principais"
'
''        UserForm2.ex_ListBox.ColumnCount = 3
'        'UserForm2.ex_ListBox.AddItem ""
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = "Insumo"
                


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
    
 'LISTAR DIVERSOS
    If counte1 <> 0 Then
           ConectarBanco conexao
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Diversos_Rendimento WHERE idInsumo = " & id & ";"

        rs.Open sql, conexao
        rs.MoveFirst
'        UserForm2.ex_ListBox.Clear
            
    
        If counte1 > 0 Then
        UserForm2.ex_ListBox.ColumnCount = 4
        UserForm2.ex_ListBox.AddItem ""
'        'UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Serviços Diversos"
'
''        UserForm2.ex_ListBox.ColumnCount = 3
'        'UserForm2.ex_ListBox.AddItem ""
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"
'        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = "Insumo"
                


    ' Ajusta o tamanho das colunas
'    UserForm2.ex_ListBox.columnWidths = "100;150;400;200" ' Ajuste o tamanho conforme necessário
            
        Do Until rs.EOF
            
        Dim strNv1I As Integer
       strNv1I = rs("idNv1")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nive1ISQL As String
        nive1ISQL = "SELECT descricaoNv1 FROM t_Nivel1 WHERE idNv1 = " & strNv1I & ";"
        Dim rsDescricaoNv1I As ADODB.Recordset
        Set rsDescricaoNv1I = New ADODB.Recordset
        rsDescricaoNv1I.Open nive1ISQL, conexao
        
    
       Dim strNv2I As Integer
       strNv2I = rs("idNv2")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nivel2ISQL As String
        nivel2ISQL = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & strNv2I & ";"
        Dim rsDescricaoNv2I As ADODB.Recordset
        Set rsDescricaoNv2I = New ADODB.Recordset
        rsDescricaoNv2I.Open nivel2ISQL, conexao
        
      Dim strNv3IP As Integer
       strNv3IP = rs("idNv3")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nivel3SQLIP As String
        nivel3SQLIP = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & strNv3IP & ";"
        Dim rsDescricaoNv3IP As ADODB.Recordset
        Set rsDescricaoNv3IP = New ADODB.Recordset
        rsDescricaoNv3IP.Open nivel3SQLIP, conexao
               
    
       Dim strInsumoIP As Integer
       strInsumoIP = rs("idInsumo")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim insumoSQLIP As String
        insumoSQLIP = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & strInsumoIP & ";"
        Dim rsPInsumoI As ADODB.Recordset
       Set rsPInsumoI = New ADODB.Recordset
        rsPInsumoI.Open insumoSQLIP, conexao
        
                  
        UserForm2.ex_ListBox.ColumnCount = 4
        ' Verifica se há um valor correspondente na tabela t_Insumos
        If Not rsDescricaoNv1I.EOF Then
             'Obtém o valor de Insumo encontrado na tabela t_Insumos

            insumo = rsPInsumoI("Insumo")
            nivel3 = rsDescricaoNv3IP("DescricaoNv3")
            nivel1 = rsDescricaoNv1I("DescricaoNv1")

            
    ' Adiciona uma nova linha ao ListBox
        UserForm2.ex_ListBox.AddItem ""
    
        ' Define os valores das colunas para a nova linha
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = nivel1
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "-"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = nivel3
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = insumo
    
        ' Ajusta o tamanho das colunas
        UserForm2.ex_ListBox.columnWidths = "100;150;300" ' Ajuste o tamanho conforme necessário
    
    
        End If

            rsPInsumoI.Close
            rs.MoveNext

            


            
        Loop
 
    End If
    rs.Close
    
    
    

    
    End If
    
    Exclusão.SaveLogList
    calculatePixelsAmount
        UserForm2.ex_ListBox.columnWidths = coluna1 & ";" & coluna2 & ";" & coluna3 & ";" & coluna4 ' Ajuste o tamanho conforme necessário"
    End If
    If counte <> 0 And counte1 <> 0 Then
    MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & counte & " serviço(s). Principais(" & counte - counte1 & ") e Diversos(" & counte1 & ") registrados.", vbExclamation
    ElseIf counte + counte1 <> 0 And counte1 = 0 Then
    MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & counte - counte1 & " serviço(s) Principais registrados.", vbExclamation
    ElseIf counte + counte1 <> 0 And counte = 0 Then
    MsgBox "Não será  possível excluir o nível selecionado. Existe(m) um total de " & counte1 & " serviço(s) Diversos registrados.", vbExclamation
    
    End If
End If

End Sub


Sub SaveLogList()

GlobalComboBoxValue = UserForm2.ex_ComboBoxNiveis

If UserForm2.ex_btnNv1BooleanP = True Then
GlobalTable = "TabelaNv1"
GlobalServiceType = "Servicos Principais"
ElseIf UserForm2.ex_btnNv2BooleanP = True Then
GlobalTable = "TabelaNv2"
GlobalServiceType = "Servicos Principais"
ElseIf UserForm2.ex_btnNv3BooleanP = True Then
GlobalTable = "TabelaNv3"
GlobalServiceType = "Servicos Principais"
ElseIf UserForm2.ex_btnNv4BooleanP = True Then
GlobalTable = "TabelaNv4"
GlobalServiceType = "Insumos"
End If


'ex_btnNv4BooleanP = False

End Sub


Sub ChargeListByLog()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim insumo As String
Dim nivel1 As String
Dim nivel2 As String
Dim nivel3 As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim id As String
'Dim rs As Recordset
Dim counte As Integer
Dim db As Object
Dim idNv01 As Long
Dim idNv02 As Long
Dim idNv03 As Long

Dim idInsumo As Long
Dim result As VbMsgBoxResult

If GlobalServiceType = "Servicos Principais" Then


    
        ConectarBanco conexao
        If GlobalTable = "TabelaNv1" Then
            sql = "SELECT idNv1 FROM t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv1 = '" & GlobalComboBoxValue & "';"
        ElseIf GlobalTable = "TabelaNv2" Then
            sql = "SELECT idNv2 FROM t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv2 = '" & GlobalComboBoxValue & "';"
        ElseIf GlobalTable = "TabelaNv3" Then
            sql = "SELECT idNv3 FROM t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' AND descricaoNv3 = '" & GlobalComboBoxValue & "';"
        End If

        If GlobalTable = "TabelaNv1" Then
        rs.Open sql, conexao
        idNv01 = rs.Fields("idNv1").Value
        id = idNv01
        conexao.Close
        ElseIf GlobalTable = "TabelaNv2" Then
         rs.Open sql, conexao
        idNv02 = rs.Fields("idNv2").Value
        id = idNv02
        conexao.Close
        ElseIf GlobalTable = "TabelaNv3" Then
        rs.Open sql, conexao
        idNv03 = rs.Fields("idNv3").Value
        id = idNv03
        conexao.Close
        End If
    
    ConectarBanco conexao
    
        If GlobalTable = "TabelaNv1" Then
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
        ElseIf GlobalTable = "TabelaNv2" Then
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv2 = " & id & ";"
        ElseIf GlobalTable = "TabelaNv3" Then
        sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id & ";"
        End If
    
    '    sql = "SELECT COUNT(*) AS counte FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
    rs.Open sql, conexao
    counte = rs.Fields("counte").Value
    counte = rs("counte")
    rs.Close
    conexao.Close
    
    If counte = 0 Then
    Dim nivel As Integer
    Dim strNivel As String
        If GlobalTable = "Tabela1" Then
        nivel = 1
            MsgBox "O nível " & nivel & " não possui mais estruturas vinculadas." & vbCrLf & "Nível liberado para exclusão.", vbInformation
        ElseIf GlobalTable = "Tabela2" Then
        nivel = 1
            MsgBox "O nível " & nivel & " não possui mais estruturas vinculadas." & vbCrLf & "Nível liberado para exclusão.", vbInformation
        ElseIf GlobalTable = "Tabela3" Then
        strNivel = "Serviço"
            MsgBox "O nível " & strNivel & " não possui mais estruturas vinculadas." & vbCrLf & "Nível liberado para exclusão.", vbInformation
        ElseIf GlobalTable = "Tabela4" Then
       strNivel = "Insumo"
           MsgBox "O nível " & strNivel & " não possui mais estruturas vinculadas." & vbCrLf & "Nível liberado para exclusão.", vbInformation
        End If
         


    UserForm2.ex_ListBox.Clear
    Exit Sub
    End If

        
     ConectarBanco conexao
    ' Monta o código SQL para buscar as linhas com o idNv1 selecionado
        If GlobalTable = "TabelaNv1" Then
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
        ElseIf GlobalTable = "TabelaNv2" Then
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv2 = " & id & ";"
        ElseIf GlobalTable = "TabelaNv3" Then
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv3 = " & id & ";"
        End If
'    sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idNv1 = " & id & ";"
    rs.Open sql, conexao
    rs.MoveFirst

    UserForm2.ex_ListBox.Clear
'    UserForm2.ex_ListBox.AddItem "         Níve 1                    Nível 2                           Serviço                                                   Insumo"
        
    
    If counte > 0 Then
            UserForm2.ex_ListBox.ColumnCount = 4
            UserForm2.ex_ListBox.AddItem ""
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"


    ' Ajusta o tamanho das colunas
    UserForm2.ex_ListBox.columnWidths = "100;100;100;100" ' Ajuste o tamanho conforme necessário
            
        Do Until rs.EOF
'            UserForm2.ex_ListBox.AddItem rs("idNv1") & " | " & rs("idNv2") & " | " & rs("idNv3") & " | " & rs("idInsumo")
'            rs.MoveNext
            
            
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
            ' Obtém o valor de Insumo encontrado na tabela t_Insumos
            'Dim insumo As String
            'insumo = rsPInsumo("Insumo")
            nivel3 = rsDescricaoNv3P("DescricaoNv3")
            nivel2 = rsDescricaoNv2("DescricaoNv2")
            nivel1 = rsDescricaoNv1("DescricaoNv1")
' Adiciona uma nova linha ao ListBox
    UserForm2.ex_ListBox.AddItem ""
    ' Define os valores das colunas para a nova linha
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = nivel1
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = nivel2
    UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = nivel3



    ' Ajusta o tamanho das colunas
    UserForm2.ex_ListBox.columnWidths = "300;300;800" ' Ajuste o tamanho conforme necessário"
    
    

        End If
        rsPInsumo.Close
        rs.MoveNext

        Loop
    End If
    rs.Close
    
    
    
    '====================================================================
    
    Dim Num_Linhas, i, X As Integer
    
    Num_Linhas = UserForm2.ex_ListBox.ListCount - 1
    
    For i = 0 To Num_Linhas
    For X = Num_Linhas To (i + 1) Step -1
        If UserForm2.ex_ListBox.List(i, 0) = UserForm2.ex_ListBox.List(X, 0) _
            And UserForm2.ex_ListBox.List(i, 1) = UserForm2.ex_ListBox.List(X, 1) _
            And UserForm2.ex_ListBox.List(i, 2) = UserForm2.ex_ListBox.List(X, 2) Then
            UserForm2.ex_ListBox.RemoveItem X
            Num_Linhas = Num_Linhas - 1
        End If
    Next X
    Next i
    
    
    '====================================================================



    MsgBox "Existe(m) um total de " & Num_Linhas & " serviço(s) restantes.", vbExclamation
    End If


    


End Sub


Sub ChargeListByLogInsumo()
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



    
    
    
    
    ConectarBanco conexao
    sql = "SELECT idInsumo FROM t_Insumos WHERE Insumo = '" & GlobalComboBoxValue & "';"


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
    
    Dim counte1 As Long
    
    ConectarBanco conexao
    sql = "SELECT COUNT(*) AS counte1 FROM t_Servicos_Diversos_Rendimento WHERE idInsumo = " & id & ";"
    rs.Open sql, conexao
    counte1 = rs.Fields("counte1").Value
    counte1 = rs("counte1")
    rs.Close
    conexao.Close
    Dim counter As Integer
    counter = counte + counte1
    
    
    
        If counte = 0 And counte1 = 0 Then
    Dim nivel As Integer
    Dim strNivel As String

        If GlobalTable = "TabelaNv4" Then
        strNivel = "Insumo"
        MsgBox "O nível " & strNivel & " não possui mais estruturas vinculadas." & vbCrLf & "Nível liberado para exclusão.", vbInformation
        End If
         


    UserForm2.ex_ListBox.Clear
    Exit Sub
    End If
    
    
    
    
  
        
'---------------------- LISTAR VALORES -------------------------------
        UserForm2.ex_ListBox.Clear
        UserForm2.ex_ListBox.ColumnCount = 4
        UserForm2.ex_ListBox.AddItem ""
        'UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Serviços Principais"

'        UserForm2.ex_ListBox.ColumnCount = 3
        'UserForm2.ex_ListBox.AddItem ""
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = "Nível 1"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "Nível 2"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = "Nível 3"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = "Insumo"
        
If counte <> 0 Then

        ConectarBanco conexao
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Principais_Rendimento WHERE idInsumo = " & id & ";"

        rs.Open sql, conexao
        rs.MoveFirst
       ' UserForm2.ex_ListBox.Clear
            
    
        If counte > 0 Then

                


    ' Ajusta o tamanho das colunas
    UserForm2.ex_ListBox.columnWidths = "100;100;400" ' Ajuste o tamanho conforme necessário
            
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
    Dim Num_Linhas As Integer
  '  Num_Linhas = UserForm2.ex_ListBox.ListCount - 1
  '  MsgBox "Existe(m) um total de " & Num_Linhas & " serviço(s) restantes.", vbExclamation
    
    
    
    
    
    
    
    If counte1 <> 0 Then
           ConectarBanco conexao
        sql = "SELECT idNv1, idNv2, idNv3, idInsumo, rendimento FROM t_Servicos_Diversos_Rendimento WHERE idInsumo = " & id & ";"

        rs.Open sql, conexao
        rs.MoveFirst
'        UserForm2.ex_ListBox.Clear
            
    
        If counte1 > 0 Then
        UserForm2.ex_ListBox.ColumnCount = 4
        UserForm2.ex_ListBox.AddItem ""

                


    ' Ajusta o tamanho das colunas
'    UserForm2.ex_ListBox.columnWidths = "100;150;400;200" ' Ajuste o tamanho conforme necessário
            
        Do Until rs.EOF
            
        Dim strNv1I As Integer
       strNv1I = rs("idNv1")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nive1ISQL As String
        nive1ISQL = "SELECT descricaoNv1 FROM t_Nivel1 WHERE idNv1 = " & strNv1I & ";"
        Dim rsDescricaoNv1I As ADODB.Recordset
        Set rsDescricaoNv1I = New ADODB.Recordset
        rsDescricaoNv1I.Open nive1ISQL, conexao
        
    
       Dim strNv2I As Integer
       strNv2I = rs("idNv2")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nivel2ISQL As String
        nivel2ISQL = "SELECT descricaoNv2 FROM t_Nivel2 WHERE idNv2 = " & strNv2I & ";"
        Dim rsDescricaoNv2I As ADODB.Recordset
        Set rsDescricaoNv2I = New ADODB.Recordset
        rsDescricaoNv2I.Open nivel2ISQL, conexao
        
      Dim strNv3IP As Integer
       strNv3IP = rs("idNv3")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim nivel3SQLIP As String
        nivel3SQLIP = "SELECT descricaoNv3 FROM t_Nivel3 WHERE idNv3 = " & strNv3IP & ";"
        Dim rsDescricaoNv3IP As ADODB.Recordset
        Set rsDescricaoNv3IP = New ADODB.Recordset
        rsDescricaoNv3IP.Open nivel3SQLIP, conexao
               
    
       Dim strInsumoIP As Integer
       strInsumoIP = rs("idInsumo")
        ' Consulta a tabela t_Insumos para obter o valor de Insumo correspondente a idNv3
        Dim insumoSQLIP As String
        insumoSQLIP = "SELECT Insumo FROM t_Insumos WHERE idInsumo = " & strInsumoIP & ";"
        Dim rsPInsumoI As ADODB.Recordset
       Set rsPInsumoI = New ADODB.Recordset
        rsPInsumoI.Open insumoSQLIP, conexao
        
                  
        UserForm2.ex_ListBox.ColumnCount = 4
        ' Verifica se há um valor correspondente na tabela t_Insumos
        If Not rsDescricaoNv1I.EOF Then
             'Obtém o valor de Insumo encontrado na tabela t_Insumos

            insumo = rsPInsumoI("Insumo")
            nivel3 = rsDescricaoNv3IP("DescricaoNv3")
            nivel1 = rsDescricaoNv1I("DescricaoNv1")

            
    ' Adiciona uma nova linha ao ListBox
        UserForm2.ex_ListBox.AddItem ""
    
        ' Define os valores das colunas para a nova linha
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 0) = nivel1
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 1) = "-"
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 2) = nivel3
        UserForm2.ex_ListBox.List(UserForm2.ex_ListBox.ListCount - 1, 3) = insumo
    
        ' Ajusta o tamanho das colunas
        UserForm2.ex_ListBox.columnWidths = "100;150;300" ' Ajuste o tamanho conforme necessário
    
    
        End If

            rsPInsumoI.Close
            rs.MoveNext

            


            
        Loop
 
    End If
    rs.Close
    End If
    
End Sub


