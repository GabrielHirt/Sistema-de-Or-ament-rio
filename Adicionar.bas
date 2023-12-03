Attribute VB_Name = "Adicionar"
Option Explicit

Sub Inserir()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "Tabela1"

rs.Open sql, conexao, adOpenKeyset, adLockOptimistic

rs.AddNew
    rs!Nome = UserForm1.txtNome.Value
    
    If UserForm1.txtDataNasc.Value <> "" Then
    rs![Data de Nascimento] = UserForm1.txtDataNasc.Value
    End If
    
    rs![Peso] = UserForm1.txtPeso.Value
    rs![Obs] = UserForm1.txtObs.Value
      
rs.Update

rs.Close
conexao.Close
 
MsgBox "Inserido com sucesso"
End Sub



Sub CarregarNivelUm()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String


Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel1"

rs.Open "select descricaoNv1 from t_Nivel1 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv1", conexao, 3, 3



Do Until rs.EOF
UserForm2.txtComboBox_nv1_1.AddItem rs!descricaoNv1

rs.MoveNext
Loop

conexao.Close


End Sub


Sub CarregarNivelDois()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Nivel2"

rs.Open "select descricaoNv2 from t_Nivel2 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv2", conexao, 3, 3

Do Until rs.EOF
UserForm2.txtComboBox_nv2_2.AddItem rs!descricaoNv2

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

rs.Open "select descricaoNv3 from t_Nivel3 WHERE grupo = 'SERVICOS PRINCIPAIS' order BY idNv3", conexao, 3, 3

Do Until rs.EOF
UserForm2.txtComboBox_nv3_3.AddItem rs!descricaoNv3

rs.MoveNext
Loop

conexao.Close


End Sub



Sub CarregarInsumos()

Dim conexao As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim db As New ADODB.Command

Set conexao = New ADODB.Connection
Set rs = New ADODB.Recordset

ConectarBanco conexao

sql = "t_Insumos"

UserForm2.txtComboBox_nv4_4.Clear
If UserForm2.sPrincipais = True Then
rs.Open "SELECT Insumo FROM t_Insumos WHERE (Tipo = 'SERVICOS PRINCIPAIS' OR Tipo = 'GENERICO') ORDER BY idInsumo", conexao, 3, 3
ElseIf UserForm2.sDiversos = True Then
rs.Open "SELECT Insumo FROM t_Insumos WHERE (Tipo = 'SERVICOS DIVERSOS' OR Tipo = 'GENERICO') ORDER BY idInsumo", conexao, 3, 3
ElseIf UserForm2.sDiversos = False And UserForm2.sPrincipais = False Then
rs.Open "SELECT Insumo FROM t_Insumos WHERE (Tipo = 'SERVICOS PRINCIPAIS' OR Tipo = 'GENERICO') ORDER BY idInsumo", conexao, 3, 3
End If



Do Until rs.EOF
UserForm2.txtComboBox_nv4_4.AddItem rs!insumo

rs.MoveNext
Loop

conexao.Close



End Sub


Private Sub ListBox1_AfterUpdate()
    Dim selectedRow As Integer
    Dim selectedValue As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    ' Obter a posição da linha selecionada na lista
    selectedRow = UserForm2.txtComboBox_nv3_3.ListIndex
    
    ' Obter o valor da coluna 2 (por exemplo) da mesma linha na tabela
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM t_Nivel1")
    '"select descricaoNv1 from t_Nivel1"
    rs.MoveFirst
    rs.Move selectedRow
    
    selectedValue = rs.Fields(3).Value
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Fazer algo com o valor obtido
    MsgBox selectedValue
    UserForm2.ComboBox10.Value = selectedValue
End Sub
