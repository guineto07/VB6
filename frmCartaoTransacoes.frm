VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCadCartaoTransacoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de transações de cartão de crédito"
   ClientHeight    =   4545
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9090
   Icon            =   "frmCartaoTransacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialogExcel 
      Left            =   8310
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   3900
      Width           =   9525
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar Transações "
         Height          =   435
         Left            =   7290
         TabIndex        =   10
         Top             =   120
         Width           =   1755
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "&Incluir"
         Height          =   435
         Left            =   90
         TabIndex        =   6
         Top             =   120
         Width           =   1755
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   435
         Left            =   1890
         TabIndex        =   7
         Top             =   120
         Width           =   1755
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   435
         Left            =   5490
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1755
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "E&xcluir"
         Height          =   435
         Left            =   3690
         TabIndex        =   8
         Top             =   120
         Width           =   1755
      End
   End
   Begin VB.Frame FrameDados 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      TabIndex        =   12
      Top             =   150
      Width           =   9555
      Begin VB.CommandButton cmdConsultarCartao 
         Caption         =   "&Consultar Cartão"
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   55
         TabIndex        =   11
         Text            =   "txtNome"
         Top             =   2130
         Width           =   6345
      End
      Begin VB.TextBox txtId_Transacao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "txtId_Transacao"
         Top             =   450
         Width           =   2025
      End
      Begin VB.TextBox txtValor_Transacao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   6660
         MaxLength       =   18
         TabIndex        =   3
         Text            =   "txtValor_Transacao"
         Top             =   2130
         Width           =   2025
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1770
         MaxLength       =   55
         TabIndex        =   5
         Text            =   "txtDescricao"
         Top             =   2970
         Width           =   6915
      End
      Begin MSMask.MaskEdBox mskData_Transacao 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   2970
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskNumero_Cartao 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   1290
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   19
         Mask            =   "#### #### #### ####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nome Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id Transação"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   150
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Numero Cartão"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor Transacão"
         Height          =   195
         Left            =   6660
         TabIndex        =   15
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data da Transacão"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2670
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1770
         TabIndex        =   13
         Top             =   2670
         Width           =   720
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "Consultas"
      Begin VB.Menu mnuConsCateg 
         Caption         =   "Consulta transações por categoria"
      End
   End
   Begin VB.Menu mnuExportar 
      Caption         =   "Exportar"
      Begin VB.Menu smnuExportaExcel 
         Caption         =   "Exportar transações do último mês para o Excel"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "frmCadCartaoTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Limpar()

    txtNome.Text = ""
    intAcaoMomento = Limpo
    txtId_Transacao.Text = ""
    mskNumero_Cartao.Text = String(19, " ")
    txtValor_Transacao.Text = ""
    mskData_Transacao.Text = "  /  /    "
    txtDescricao.Text = ""
    cmdConsultarCartao.Enabled = False
    clsTransacao.curIdTransacao = 0
    
    Call PreparaBotoesAcao
    Call HabilitaEdicao(False)
      
End Sub
Private Function HabilitaEdicao(blnHabilita As Boolean)
   
    txtId_Transacao.Enabled = blnHabilita
    mskNumero_Cartao.Enabled = blnHabilita
    txtNome.Enabled = blnHabilita
    txtValor_Transacao.Enabled = blnHabilita
    mskData_Transacao.Enabled = blnHabilita
    txtDescricao.Enabled = blnHabilita
    cmdConsultarCartao.Enabled = blnHabilita
    cmdConsultarCartao.Enabled = blnHabilita
    
End Function
Private Function Valida() As Boolean
    
    If clsTransacao.Id_Cliente = 0 Then
        MsgBox "Cliente não encontrado. Após preencher o número do cartão, clique em 'Consultar Cartão'.", vbInformation
        txtNome.Text = ""
        mskNumero_Cartao.SetFocus
        Exit Function
    End If
    
    If Len(Trim(mskNumero_Cartao.Text)) <> 19 Then
        MsgBox "Número de cartão inválido.", vbInformation
        mskNumero_Cartao.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtValor_Transacao.Text) Or Val(txtValor_Transacao.Text) = 0 Then
        MsgBox "Valor de transação inválido.", vbInformation
        txtValor_Transacao.SetFocus
        Exit Function
    End If
    
    If Not IsDate(mskData_Transacao) Then
        MsgBox "Data de transação inválida.", vbInformation
        mskData_Transacao.SetFocus
        Exit Function
    End If
    
    If CDate(mskData_Transacao.Text) > Date Then
        MsgBox "Data da transação não pode ser posterior a data atual.", vbInformation
        mskData_Transacao.SetFocus
        Exit Function
    End If
    
    If txtDescricao.Text = "" Then
        MsgBox "Descrição inválida.", vbInformation
        txtDescricao.SetFocus
        Exit Function
    End If
    
    Valida = True
End Function

Private Sub cmdConsultar_Click()
    clsTransacao.curIdTransacao = 0
    frmConsultaTransacao.Show 1
    
    With clsTransacao
        If .curIdTransacao <> 0 Then
            txtId_Transacao.Text = .curIdTransacao
            txtId_Transacao.Enabled = False
            
            txtNome.Text = .Nome

            mskNumero_Cartao = Format(.strNumero_Cartao, "0000 0000 0000 0000")
            txtValor_Transacao.Text = Format(.curValor_Transacao, "standard")
            mskData_Transacao.Text = Format(.dtData_transacao, "dd/mm/yyyy")
            txtDescricao.Text = .strDescricao
            
            intAcaoMomento = Editar
            PreparaBotoesAcao
            Call HabilitaEdicao(True)
            mskNumero_Cartao.SetFocus
        Else
            Limpar
        End If
    End With
End Sub


Private Sub cmdConsultarCartao_Click()
    Dim strNumero_Cartao As String
    strNumero_Cartao = FormataCartao(mskNumero_Cartao, False)
    Call ConsultaCliente(strNumero_Cartao)
End Sub

Private Sub cmdExcluir_Click()

    If Not msgPergunta("Confirma exclusão desta transação?") = vbYes Then Exit Sub
     
    Call clsTransacao.Excluir(clsTransacao.curIdTransacao)

    intAcaoMomento = Limpo
    Limpar
    PreparaBotoesAcao
    cmdIncluir.SetFocus
End Sub

Private Sub AtualizaClasse()
    
    With clsTransacao
        If txtId_Transacao.Text <> "" Then .curIdTransacao = txtId_Transacao
        .Id_Cliente = clsTransacao.Id_Cliente
        .strNumero_Cartao = FormataCartao(mskNumero_Cartao.Text, False)
        .curValor_Transacao = txtValor_Transacao.Text
        .dtData_transacao = Format(mskData_Transacao.Text, "yyyy-mm-dd")
        .strDescricao = txtDescricao.Text
    End With
    
End Sub
Private Sub cmdGravar_Click()
    
    Dim strTipo As String
    Dim strMsg As String
        
    If Not Valida Then Exit Sub
    
    strTipo = IIf(intAcaoMomento = Incluir, "inclusão", IIf(intAcaoMomento = Editar, "alteração", strTipo))
    strMsg = "Confirma " & strTipo & " desta transação?"
    
    If Not msgPergunta(strMsg) = vbYes Then Exit Sub
    
    Call AtualizaClasse
    Call clsTransacao.Gravar(intAcaoMomento)
    
    Limpar
    PreparaBotoesAcao
    cmdIncluir.SetFocus
    
End Sub

Private Sub cmdIncluir_Click()
    
    intAcaoMomento = Incluir
    
    Call PreparaBotoesAcao
    Call HabilitaEdicao(True)

    mskNumero_Cartao.SetFocus
    
End Sub

Private Sub cmdLimpar_Click()
    Limpar
End Sub

Private Sub Form_Initialize()
    Me.Width = 9180
    Me.Height = 5280
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call Center(Me)
    Set rsTemp = CreateObject("ADODB.Recordset")

    Limpar
End Sub

Private Sub mnuConsCateg_Click()
    frmConsultaTransacaoCategoria.Show 1
End Sub

Private Sub mnuSair_Click()
    Set CN = Nothing
    Set rsTemp = Nothing
    End
End Sub

Private Sub mskData_Transacao_GotFocus()
    
    If Not IsDate(mskData_Transacao.Text) Then
        mskData_Transacao.Text = Format(Date, "dd/mm/yyyy")
    End If
    
    mskData_Transacao.SelStart = 0
    mskData_Transacao.SelLength = Len(mskData_Transacao.Text)
    
End Sub

Sub MostraDados()
    
    With rsTemp
        txtId_Transacao.Text = !Id_Transacao
        mskNumero_Cartao = Format(!Numero_Cartao, "0000 0000 0000 0000")
        txtValor_Transacao.Text = Format(!Valor_Transacao, "standard")
        mskData_Transacao.Text = Format(!Data_Transacao, "dd/mm/yyyy")
        txtDescricao.Text = !Descricao
    End With
    
End Sub

Private Sub mskNumero_Cartao_GotFocus()

    mskNumero_Cartao.SelStart = 0
    mskNumero_Cartao.SelLength = Len(mskNumero_Cartao.Text)
    
End Sub

Private Sub mskNumero_Cartao_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        If clsTransacao.Id_Cliente <> 0 Then
            clsTransacao.Id_Cliente = 0
        End If
    End If
End Sub

Private Sub smnuSair_Click()
    End
End Sub

Private Sub smnuExportaExcel_Click()
    ExportarTransacoes
End Sub

Private Sub txtDescricao_GotFocus()
    txtDescricao.SelStart = 0
    txtDescricao.SelLength = Len(txtDescricao.Text)
End Sub

Private Sub txtId_Transacao_KeyPress(KeyAscii As Integer)
    Call SoNumeros(KeyAscii, 1)
End Sub

Private Sub txtValor_Transacao_GotFocus()

    txtValor_Transacao.SelStart = 0
    txtValor_Transacao.SelLength = Len(txtValor_Transacao.Text)
    
End Sub

Private Sub txtValor_Transacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub txtValor_Transacao_KeyPress(KeyAscii As Integer)
    SoNumeros KeyAscii
End Sub

Private Sub txtValor_Transacao_LostFocus()
    If IsNumeric(txtValor_Transacao) Then
       txtValor_Transacao = Format(txtValor_Transacao.Text, "standard")
    End If
End Sub
Private Sub PreparaBotoesAcao()
    
    If intAcaoMomento = Incluir Then
    
        cmdIncluir.Enabled = False
        cmdConsultar.Enabled = False
        cmdGravar.Enabled = True
        cmdExcluir.Enabled = False
        cmdLimpar.Enabled = True
          
    ElseIf intAcaoMomento = Editar Then
    
        cmdIncluir.Enabled = False
        cmdConsultar.Enabled = True
        cmdGravar.Enabled = True
        cmdExcluir.Enabled = True
        cmdLimpar.Enabled = True
        
    ElseIf intAcaoMomento = Limpo Then
    
        cmdIncluir.Enabled = True
        cmdConsultar.Enabled = True
        cmdGravar.Enabled = False
        cmdExcluir.Enabled = False
        cmdLimpar.Enabled = False
    
    End If
    
End Sub

Private Sub ConsultaCliente(Numero_Cartao As String)
On Error GoTo Err_Consulta
    Numero_Cartao = FormataCartao(Numero_Cartao, False)
    strSQL = "SELECT Id_Cliente, Nome FROM Clientes WHERE Numero_Cartao = '" & Numero_Cartao & "'"
    
    Set rsTemp = CN.Execute(strSQL)
    
    If Not rsTemp.EOF Then
       clsTransacao.Id_Cliente = rsTemp!Id_Cliente
       txtNome.Text = rsTemp!Nome
       txtValor_Transacao.SetFocus
       cmdConsultarCartao.Enabled = False
    Else
       MsgBox "Cartão não encontrado.", vbCritical, "Cadastro de transações"
       mskNumero_Cartao.SetFocus
    End If
    
Err_Consulta:
    If Err.Number <> 0 Then
        MsgBox "Erro ao consultar: " & Err.Description, vbCritical
    End If
End Sub
Private Sub ExportarTransacoes()
    Dim strSQL As String
    Dim rs As Object
    Dim filePath As String
    Dim ExcelApp As Object
    Dim ExcelWorkbook As Object
    Dim ExcelSheet As Object
    Dim row As Integer
    Dim CommonDialogExp As Object
    
    Me.MousePointer = 11
    
    On Error GoTo Err_Exp
    
    strSQL = "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, dbo.fn_CategoriaTransacao(Valor_Transacao) AS Categoria "
    strSQL = strSQL & "FROM Cartao_Transacoes "
    strSQL = strSQL & "WHERE Data_Transacao >= DATEADD(MONTH, -1, GETDATE())"
    
    DoEvents
    Set rs = CN.Execute(strSQL)
    
    Set CommonDialogExp = CreateObject("MSComDlg.CommonDialog")
    CommonDialogExp.CancelError = True
    CommonDialogExp.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"
    CommonDialogExp.ShowSave
    
    If Len(CommonDialogExp.FileName) > 0 Then
        filePath = CommonDialogExp.FileName
    Else
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = False
    
    Set ExcelWorkbook = ExcelApp.Workbooks.Add
    Set ExcelSheet = ExcelWorkbook.Sheets(1)
    
    ' Cabeçalhos
    ExcelSheet.Cells(1, 1).Value = "Numero_Cartao"
    ExcelSheet.Cells(1, 2).Value = "Valor_Transacao"
    ExcelSheet.Cells(1, 3).Value = "Data_Transacao"
    ExcelSheet.Cells(1, 4).Value = "Descricao"
    ExcelSheet.Cells(1, 5).Value = "Categoria"
    
    ' Ajustar a largura das colunas manualmente para garantir boa leitura
    ExcelSheet.Columns(1).ColumnWidth = 20  ' Número do Cartão
    ExcelSheet.Columns(2).ColumnWidth = 20  ' Valor da Transação
    ExcelSheet.Columns(3).ColumnWidth = 18  ' Data da Transação
    ExcelSheet.Columns(4).ColumnWidth = 30  ' Descrição
    ExcelSheet.Columns(5).ColumnWidth = 20  ' Categoria
    
    row = 2
    Do While Not rs.EOF
        ' Definir o formato de número como texto para o número do cartão
        ExcelSheet.Cells(row, 1).NumberFormat = "@"
        ExcelSheet.Cells(row, 1).Value = rs.Fields("Numero_Cartao").Value
        
        ' Aplicar formato monetário para o Valor_Transacao
        ExcelSheet.Cells(row, 2).NumberFormat = "#,##0.00"
        ExcelSheet.Cells(row, 2).Value = rs.Fields("Valor_Transacao").Value
    
        ' Definir formato de data para Data_Transacao
        ExcelSheet.Cells(row, 3).NumberFormat = "mm/dd/yyyy"
        ExcelSheet.Cells(row, 3).Value = rs.Fields("Data_Transacao").Value
        
        ExcelSheet.Cells(row, 4).Value = rs.Fields("Descricao").Value
        ExcelSheet.Cells(row, 5).Value = rs.Fields("Categoria").Value
        
        rs.MoveNext
        row = row + 1
    Loop
    
    If LCase(Right(filePath, 4)) = ".xls" Then
        ExcelWorkbook.SaveAs filePath, 56
    ElseIf LCase(Right(filePath, 5)) = ".xlsx" Then
        ExcelWorkbook.SaveAs filePath, 51
    End If
    
    ExcelWorkbook.Close
    ExcelApp.Quit
    
Err_Exp:
    Set ExcelSheet = Nothing
    Set ExcelWorkbook = Nothing
    Set ExcelApp = Nothing
    Set rs = Nothing
    Set CommonDialogExp = Nothing
    
    MousePointer = 0
    
    If Err.Number <> 0 Then
        If Err.Number = 32755 Then Exit Sub
        MsgBox "Erro na exportação, feche o arquivo se estiver aberto no Excel: " & Err.Description
    Else
        MsgBox "Transações exportadas com sucesso.", vbInformation
    End If
    
    


End Sub

