VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarId_Cliente As Long
Private mvarstrNumero_Cartao As Variant
Private mvarcurNome As String
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
        strSQL = "INSERT INTO Clientes (Nome,Numero_Cartao) "
        strSQL = strSQL & " VALUES ('"
        strSQL = strSQL & clsCliente.Nome & "',"
        strSQL = strSQL & "'" & clsCliente.strNumero_Cartao & "')"
        
    ElseIf intAcaoMomento = Editar Then
        strSQL = "UPDATE Clientes SET "
        strSQL = strSQL & "Nome='" & clsCliente.Nome & "', "
        strSQL = strSQL & "Numero_Cartao='" & clsCliente.strNumero_Cartao & "'"
        strSQL = strSQL & "WHERE Id_Cliente=" & clsCliente.Id_Cliente
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

Public Function Excluir(lngId_Cliente As Long) As Boolean
    On Error GoTo err_excluir
    
    CN.BeginTrans
    strSQL = "DELETE FROM Clientes"
    strSQL = strSQL & " WHERE Id_Cliente=" & lngId_Cliente
    CN.Execute strSQL
err_excluir:
    If Err.Number = -2147217873 Then
        MsgBox "Erro ao excluir: Cliente possui lançamento de transações.", vbCritical
    ElseIf Err.Number = 0 Then
        CN.CommitTrans
        Excluir = True
    End If
End Function

