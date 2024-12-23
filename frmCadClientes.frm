VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCadClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   Icon            =   "frmCadClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameDados 
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   0
      TabIndex        =   8
      Top             =   30
      Width           =   9555
      Begin VB.TextBox txtId_Cliente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "txtId_Cliente"
         Top             =   450
         Width           =   1965
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   150
         MaxLength       =   55
         TabIndex        =   0
         Text            =   "txtNome"
         Top             =   1440
         Width           =   7485
      End
      Begin MSMask.MaskEdBox mskNumero_Cartao 
         Height          =   345
         Left            =   150
         TabIndex        =   1
         Top             =   2460
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   19
         Mask            =   "#### #### #### ####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Numero Cartão"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   2190
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   150
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nome Cliente"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   1140
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   -60
      TabIndex        =   7
      Top             =   4110
      Width           =   9615
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "E&xcluir"
         Height          =   435
         Left            =   3780
         TabIndex        =   4
         Top             =   180
         Width           =   1755
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   435
         Left            =   5580
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1755
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   435
         Left            =   1980
         TabIndex        =   3
         Top             =   180
         Width           =   1755
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "&Incluir"
         Height          =   435
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   1755
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar Clientes"
         Height          =   435
         Left            =   7380
         TabIndex        =   6
         Top             =   180
         Width           =   1755
      End
   End
   Begin MSComDlg.CommonDialog CommonDialogExcel 
      Left            =   8310
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnEditouCartao As Boolean

Private Sub Limpar()

    txtNome.Text = ""
    intAcaoMomento = Limpo
    txtId_Cliente.Text = ""
    mskNumero_Cartao.Text = String(19, " ")
    
    txtId_Cliente.Text = ""
    
    If Not clsCliente Is Nothing Then
        clsCliente.Id_Cliente = 0
    End If
        
    PreparaBotoesAcao
    HabilitaEdicao (False)
    
    blnEditouCartao = True
    
    Set clsCliente = Nothing
    
End Sub
Private Function HabilitaEdicao(blnHabilita As Boolean)
   
    txtId_Cliente.Enabled = blnHabilita
    mskNumero_Cartao.Enabled = blnHabilita
    txtNome.Enabled = blnHabilita
    
End Function
Private Function Valida() As Boolean
  
    If Len(Trim(mskNumero_Cartao.Text)) <> 19 Then
        MsgBox "Número de cartão inválido.", vbInformation
        mskNumero_Cartao.SetFocus
        Exit Function
    End If
    
    If txtNome.Text = "" Then
        MsgBox "Nome inválido.", vbInformation
        txtNome.SetFocus
        Exit Function
    End If
    
    Valida = True
End Function

Private Sub cmdConsultar_Click()
    
    frmConsultaClientes.Show 1
    If clsCliente Is Nothing Then Exit Sub
    
    'Seta dados retornados da consulta
    With clsCliente
        If .Id_Cliente <> 0 Then
            txtId_Cliente.Text = .Id_Cliente
            txtId_Cliente.Enabled = False
            
            txtId_Cliente.Text = .Id_Cliente
            txtNome.Text = .Nome

            mskNumero_Cartao = Format(.strNumero_Cartao, "0000 0000 0000 0000")
            
            intAcaoMomento = Editar
            PreparaBotoesAcao
            HabilitaEdicao (True)
            mskNumero_Cartao.SetFocus
        Else
            Limpar
        End If
    End With
End Sub


Private Sub cmdExcluir_Click()

    If Not msgPergunta("Confirma exclusão deste cliente?") = vbYes Then Exit Sub
     
    clsCliente.Excluir (clsCliente.Id_Cliente)

    intAcaoMomento = Limpo
    Limpar
    PreparaBotoesAcao
    cmdIncluir.SetFocus
End Sub

Private Sub AtualizaClasse()
    
    With clsCliente
        If txtId_Cliente.Text <> "" Then .Id_Cliente = txtId_Cliente
        .Nome = txtNome.Text
        .strNumero_Cartao = FormataCartao(mskNumero_Cartao.Text, False)
    End With
    
End Sub
Private Sub cmdGravar_Click()
    
    Dim strTipo As String
    Dim strMsg As String
        
    If Not Valida Then Exit Sub
    
    strTipo = IIf(intAcaoMomento = Incluir, "inclusão", IIf(intAcaoMomento = Editar, "alteração", strTipo))
    strMsg = "Confirma " & strTipo & " desta transação?"
    
    If msgPergunta(strMsg) = vbYes Then
    
         Set clsCliente = New clsClientes
         AtualizaClasse
    
         clsCliente.Gravar (intAcaoMomento)
    Else
        Exit Sub
    End If
    
    Limpar
    PreparaBotoesAcao
    cmdIncluir.SetFocus
    
End Sub

Private Sub cmdIncluir_Click()
    
    intAcaoMomento = Incluir
    
    PreparaBotoesAcao
    HabilitaEdicao (True)

    txtNome.SetFocus
    
End Sub

Private Sub cmdLimpar_Click()
    Limpar
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
    Set rsTemp = CreateObject("ADODB.Recordset")
    Limpar
    
    Me.Width = 9180
    Me.Height = 5280
    Center Me
    
    
End Sub

Private Sub mnuCadClientes_Click()
    frmCadClientes.Show 1
End Sub

Private Sub mnuConsCateg_Click()
    frmConsultaTransacaoCategoria.Show 1
End Sub

Sub MostraDados()
    
    With rsTemp
        txtId_Cliente.Text = !Id_Transacao
        mskNumero_Cartao = Format(!Numero_Cartao, "0000 0000 0000 0000")
        txtNome.Text = !Nome
    End With
    
End Sub

Private Sub mskNumero_Cartao_GotFocus()

    mskNumero_Cartao.SelStart = 0
    mskNumero_Cartao.SelLength = Len(mskNumero_Cartao.Text)
    
End Sub

Private Sub mskNumero_Cartao_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 48 And KeyAscii <= 57 Then 'Considerando apenas caracteres
    
        If blnEditouCartao = True Then Exit Sub
        
        If clsCliente Is Nothing Then Exit Sub
        
        If clsCliente.Id_Cliente <> 0 Then
            blnEditouCartao = True
            clsCliente.Id_Cliente = 0
        End If
    End If
End Sub


Private Sub txtNome_GotFocus()
    txtNome.SelStart = 0
    txtNome.SelLength = Len(txtNome.Text)
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


