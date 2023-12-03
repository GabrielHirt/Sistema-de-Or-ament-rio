Attribute VB_Name = "Conectar"
Option Explicit

Function ConectarBanco(conexao As ADODB.Connection)


Dim Provider As String, dataSource As String, caminho As String
Dim connectionString As String

'caminho = ThisWorkbook.Path & "\Banco de Dados.Accdb;"
caminho = ThisWorkbook.Path & "\DB_Plan_Orc.Accdb;"

Provider = "Provider=Microsoft.ACE.OLEDB.12.0;"
dataSource = "Data Source=" & caminho

connectionString = Provider & dataSource

conexao.Open connectionString


End Function
