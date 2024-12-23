Attribute VB_Name = "modVariaveis"
Public CN As Object
Public rsTemp As Object

Public clsTransacao As clsCartao_Transacao
Public clsCliente As clsClientes

Public strSQL As String
Public intAcaoMomento As Integer

Public Const Limpo = 0
Public Const Incluir = 1
Public Const Editar = 2
Public strRet As Variant


