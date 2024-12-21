Attribute VB_Name = "modFuncoes"
Option Explicit
Public Function SoNumeros(KeyAscii, Optional Nao_Aceita_Virgula_ou_ponto As Integer)
   
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Then
        Exit Function
    Else
        If KeyAscii = Asc(".") Then
            KeyAscii = Asc(",")
        Else
            KeyAscii = 0
        End If
    End If
    
End Function
Public Function Enter(KeyAscii As Integer) As String
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}", True: KeyAscii = 0
End Function

Public Function AbreConexaoDB() As Boolean
    
On Error GoTo Err_AbreConexaoDB

    Dim strConexao As String
    Dim strEndereco_db As String
    Dim strNm_db As String
    Dim struser As String
    Dim strSenha As String
    
    Set CN = CreateObject("ADODB.Connection")
    
    If CN.State = 1 Then Exit Function 'Aborta caso ja esteja conectado
    
    strEndereco_db = "177.55.110.44"
    strNm_db = "Desenvolvimento"
    struser = "sa"
    strSenha = "x7kZ7U9PP0@#$%*"
    
    strConexao = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & struser & ";Password=" & strSenha & ";Initial Catalog=" & strNm_db & ";Data Source=" & strEndereco_db
    CN.Open strConexao

Err_AbreConexaoDB:
    If Err.Number <> 0 Then
        MsgBox "Nao foi possivel conectar ao banco de dados: " & Err.Description, vbCritical
        End
    Else
        AbreConexaoDB = True
    End If
    
End Function

Public Function msgPergunta(Mensagem As String) As VbVarType
    
    msgPergunta = MsgBox(Mensagem, vbQuestion + vbYesNo, "Cartão Transações")
    
End Function

Public Function FormataCartao(strNumero_Cartao As String, blnEspacos As Boolean) As String

    If blnEspacos = True Then
        FormataCartao = Format(strNumero_Cartao, "0000 0000 0000 0000")
    Else
        FormataCartao = Replace(strNumero_Cartao, " ", "")
    End If
    
End Function
Public Function Center(frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Function


