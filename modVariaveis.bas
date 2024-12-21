Attribute VB_Name = "modVariaveis"
'Public CN As New ADODB.Connection
'Public rsTemp As New ADODB.Recordset

Public CN As Object
Public rsTemp As Object

Public clsTransacao As New clsCartao_Transacao

Public strSQL As String
Public intAcaoMomento As Integer

Public Const Limpo = 0
Public Const Incluir = 1
Public Const Editar = 2
Public strRet As Variant


