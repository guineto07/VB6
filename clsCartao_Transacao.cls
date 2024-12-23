VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartao_Transacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarstrNumero_Cartao As Variant
Private mvarcurIdTransacao As Currency
Private mvarcurIdCliente As Long
Private mvarcurNome As String
Private mvarcurValor_Transacao As Currency
Private mvardtData_transacao As Variant
Private mvarstrDescricao As String
Private mvarId_Cliente As Long
Private mvarNome As String

Public Property Let Nome(ByVal vData As String)
    mvarNome = vData
End Property

Public Property Get Nome() As String
    Nome = mvarNome
End Property

Public Property Let Id_Cliente(ByVal vData As Long)
    mvarId_Cliente = vData
End Property

Public Property Get Id_Cliente() As Long
    Id_Cliente = mvarId_Cliente
End Property

Public Property Let strDescricao(ByVal vData As String)
    mvarstrDescricao = vData
End Property

Public Property Get strDescricao() As String
    strDescricao = mvarstrDescricao
End Property

Public Property Let dtData_transacao(ByVal vData As Variant)
    mvardtData_transacao = vData
End Property

Public Property Set dtData_transacao(ByVal vData As Variant)
    Set mvardtData_transacao = vData
End Property

Public Property Get dtData_transacao() As Variant
    If IsObject(mvardtData_transacao) Then
        Set dtData_transacao = mvardtData_transacao
    Else
        dtData_transacao = mvardtData_transacao
    End If
End Property

Public Property Let curValor_Transacao(ByVal vData As Currency)
    mvarcurValor_Transacao = vData
End Property

Public Property Get curValor_Transacao() As Currency
    curValor_Transacao = mvarcurValor_Transacao
End Property

Public Property Let curIdTransacao(ByVal vData As Currency)
    mvarcurIdTransacao = vData
End Property

Public Property Get curIdTransacao() As Currency
    curIdTransacao = mvarcurIdTransacao
End Property

Public Property Let strNumero_Cartao(ByVal vData As Variant)
    mvarstrNumero_Cartao = vData
End Property

Public Property Set strNumero_Cartao(ByVal vData As Variant)
    Set mvarstrNumero_Cartao = vData
End Property

Public Property Get strNumero_Cartao() As Variant
    If IsObject(mvarstrNumero_Cartao) Then
        Set strNumero_Cartao = mvarstrNumero_Cartao
    Else
        strNumero_Cartao = mvarstrNumero_Cartao
    End If
End Property

Public Function Gravar(intAcaoMomento As Integer)
    Set rsTemp = CreateObject("ADODB.Recordset")
    On Error GoTo err_Gravar
    If intAcaoMomento = Incluir Then
        strSQL = "INSERT INTO Cartao_Transacoes (Id_Cliente,Numero_Cartao,Valor_Transacao,Data_Transacao,Descricao) "
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & clsTransacao.Id_Cliente & ","
        strSQL = strSQL & "'" & clsTransacao.strNumero_Cartao & "',"
        strSQL = strSQL & Trim(Str(clsTransacao.curValor_Transacao)) & ","
        strSQL = strSQL & "'" & clsTransacao.dtData_transacao & "',"
        strSQL = strSQL & "'" & clsTransacao.strDescricao & "')"
    ElseIf intAcaoMomento = Editar Then
        strSQL = "UPDATE Cartao_Transacoes SET "
        strSQL = strSQL & "Id_Cliente=" & clsTransacao.Id_Cliente & ","
        strSQL = strSQL & "Numero_Cartao='" & clsTransacao.strNumero_Cartao & "',"
        strSQL = strSQL & "Valor_Transacao=" & Trim(Str(clsTransacao.curValor_Transacao)) & ","
        strSQL = strSQL & "Data_Transacao='" & clsTransacao.dtData_transacao & "',"
        strSQL = strSQL & "Descricao='" & clsTransacao.strDescricao & "' "
        strSQL = strSQL & "WHERE Id_Transacao=" & clsTransacao.curIdTransacao
    End If
    
    CN.BeginTrans
    CN.Execute strSQL

err_Gravar:
    If Err.Number = 0 Then
        CN.CommitTrans
    Else
        CN.RollbackTrans
        MsgBox "Erro ao gravar registro : " & Err.Description, vbCritical
    End If
End Function

Public Function Excluir(lngId_Transacao As Long)
    On Error GoTo err_excluir
    CN.BeginTrans
    strSQL = "DELETE FROM Cartao_Transacoes"
    strSQL = strSQL & " WHERE Id_Transacao=" & lngId_Transacao
    CN.Execute strSQL
err_excluir:
    If Err.Number <> 0 Then
        MsgBox "Erro ao excluir: " & Err.Description
    ElseIf Err.Number = 0 Then
        CN.CommitTrans
    End If
End Function

