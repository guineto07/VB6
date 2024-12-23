VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsultaTransacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta transação"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14610
   ControlBox      =   0   'False
   Icon            =   "frmConsultaTransacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "&Consultar"
      Height          =   465
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   2475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Height          =   465
      Left            =   240
      TabIndex        =   10
      Top             =   4650
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Filtrar consulta"
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   30
      TabIndex        =   8
      Top             =   180
      Width           =   2865
      Begin MSMask.MaskEdBox mskDataTransacaoFim 
         Height          =   315
         Left            =   1530
         TabIndex        =   5
         Top             =   1890
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtValorTransacao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   3000
         Width           =   2475
      End
      Begin VB.OptionButton optNumero_Cartao 
         Appearance      =   0  'Flat
         Caption         =   "Número Cartão"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton optDataTransacao 
         Appearance      =   0  'Flat
         Caption         =   "Data Transação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   1560
         Width           =   1515
      End
      Begin VB.OptionButton optValorTransacao 
         Appearance      =   0  'Flat
         Caption         =   "Valor Transação"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   2700
         Width           =   1725
      End
      Begin MSMask.MaskEdBox mskDataTransacao 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   1890
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskNumero_Cartao 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   810
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   19
         Mask            =   "#### #### #### ####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "a"
         Height          =   225
         Left            =   1320
         TabIndex        =   9
         Top             =   1920
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   2940
      TabIndex        =   1
      Top             =   270
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Id_Transacao"
         Caption         =   "Id Transação"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Numero_Cartao"
         Caption         =   "Número do cartão"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Id_Cliente"
         Caption         =   "Id Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Nome"
         Caption         =   "Nome Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Valor_Transacao"
         Caption         =   "Valor Transação"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   " #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Data_Transacao"
         Caption         =   "Data Transação"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Descricao"
         Caption         =   "Descrição"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   1
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1679,811
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3000,189
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   -30
      ScaleHeight     =   1815
      ScaleWidth      =   2925
      TabIndex        =   12
      Top             =   3660
      Width           =   2925
   End
   Begin VB.Label Label2 
      Caption         =   "Para abrir a transação dê um duplo clique "
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2970
      TabIndex        =   13
      Top             =   30
      Width           =   3165
   End
End
Attribute VB_Name = "frmConsultaTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsulta_Click()
    Consulta
End Sub
Private Sub Consulta()
    Set rsTemp = CreateObject("ADODB.Recordset")
    
    strSQL = "SELECT "
    strSQL = strSQL & "Cartao_Transacoes.Id_Transacao,Clientes.Id_Cliente,Clientes.Nome,"
    strSQL = strSQL & "Cartao_Transacoes.Numero_Cartao,"
    strSQL = strSQL & "Cartao_Transacoes.Valor_Transacao,"
    strSQL = strSQL & "Cartao_Transacoes.Data_Transacao,Cartao_Transacoes.Descricao "
    
    strSQL = strSQL & " FROM Cartao_Transacoes "
    strSQL = strSQL & " INNER JOIN Clientes on Clientes.id_Cliente=Cartao_Transacoes.Id_Cliente "
    
    If optNumero_Cartao.Value = True Then
        
        If Len(Trim(mskNumero_Cartao.Text)) <> 19 Then
            MsgBox "Número de cartão inválido.", vbInformation
            mskNumero_Cartao.SetFocus
            Exit Sub
        End If
        strSQL = strSQL & " WHERE Cartao_Transacoes.Numero_Cartao='" & Replace(mskNumero_Cartao.Text, " ", "") & "'"
        strSQL = strSQL & " ORDER BY Cartao_Transacoes.Numero_Cartao"
    ElseIf optDataTransacao.Value = True Then
        
        If Not IsDate(mskDataTransacao.Text) Then
            MsgBox "Data de transação inválida.", vbInformation
            mskDataTransacao.SetFocus
            Exit Sub
        End If
        
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " Cartao_Transacoes.Data_Transacao BETWEEN '" & Format(mskDataTransacao.Text, "yyyy-mm-dd") & "' AND "
        strSQL = strSQL & "'" & Format(mskDataTransacaoFim.Text, "yyyy-mm-dd") & "'"
        strSQL = strSQL & " ORDER BY Cartao_Transacoes.Data_Transacao"
    ElseIf optValorTransacao.Value = True Then
        
        If Not IsNumeric(txtValorTransacao.Text) Or Val(txtValorTransacao.Text) = 0 Then
            MsgBox "Valor de transação inválido.", vbInformation
            txtValorTransacao.SetFocus
            Exit Sub
        End If
        
        strSQL = strSQL & " WHERE Cartao_Transacoes.Valor_Transacao=" & Trim(Str(txtValorTransacao.Text))
        strSQL = strSQL & " ORDER BY Cartao_Transacoes.Valor_Transacao"
    End If
    MousePointer = 11
    DoEvents
    If rsTemp.State = 0 Then
        rsTemp.Open strSQL, CN, adOpenStatic, adLockReadOnly
    End If
    MousePointer = 0
    If rsTemp.EOF Then
        MsgBox "Nenhum registro encontrado.", vbInformation
        Set DataGrid1.DataSource = Nothing
        DataGrid1.Refresh
        Set rsTemp = Nothing
        Exit Sub
    End If

    Set DataGrid1.DataSource = rsTemp
      
    Set rsTemp = Nothing
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If DataGrid1.Text <> "" Then
              
        Set clsTransacao = New clsCartao_Transacao
        With clsTransacao
            .curIdTransacao = DataGrid1.Columns(0).Value
            .strNumero_Cartao = DataGrid1.Columns(1).Value
            .Id_Cliente = DataGrid1.Columns(2).Value
            .Nome = DataGrid1.Columns(3).Value
            .curValor_Transacao = DataGrid1.Columns(4).Value
            .dtData_transacao = DataGrid1.Columns(5).Value
            .strDescricao = DataGrid1.Columns(6).Value
        End With
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Form_Load()
     Center Me
    Limpar
    DataGrid1.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub mskDataTransacao_GotFocus()
    mskDataTransacao.SelStart = 0
    mskDataTransacao.SelLength = Len(mskDataTransacao.Text)
End Sub

Private Sub mskDataTransacao_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub mskDataTransacaoFim_GotFocus()
    mskDataTransacaoFim.SelStart = 0
    mskDataTransacaoFim.SelLength = Len(mskDataTransacaoFim.Text)
End Sub

Private Sub mskDataTransacaoFim_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdConsulta.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub mskNumero_Cartao_GotFocus()
    mskNumero_Cartao.SelStart = 0
    mskNumero_Cartao.SelLength = Len(mskNumero_Cartao.Text)
End Sub

Private Sub mskNumero_Cartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdConsulta.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub optDataTransacao_Click()
    Limpar
    
    mskDataTransacao.Text = Format(Date, "dd/mm/yyyy")
    mskDataTransacaoFim.Text = Format(Date, "dd/mm/yyyy")
    mskDataTransacao.Enabled = True
    mskDataTransacaoFim.Enabled = True
    
    mskDataTransacao.SetFocus
End Sub

Private Sub optNumero_Cartao_Click()
    Limpar
    mskNumero_Cartao.Enabled = True
    mskNumero_Cartao.SetFocus
End Sub

Private Sub optValorTransacao_Click()
    Limpar
    txtValorTransacao.Enabled = True
    txtValorTransacao.SetFocus
End Sub

Private Sub txtValorTransacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdConsulta.SetFocus
        KeyCode = 0
    End If
End Sub

Private Sub txtValorTransacao_KeyPress(KeyAscii As Integer)
    SoNumeros KeyAscii
End Sub

Private Sub Limpar()
    mskNumero_Cartao.Text = String(19, " ")
    mskNumero_Cartao.Enabled = True
    
    mskDataTransacao.Text = "  /  /    "
    mskDataTransacao.Enabled = False
    
    mskDataTransacaoFim.Text = "  /  /    "
    mskDataTransacaoFim.Enabled = False

    txtValorTransacao.Text = ""
    txtValorTransacao.Enabled = False
    
    Set DataGrid1.DataSource = Nothing
    DataGrid1.Refresh
    Set rsTemp = Nothing
End Sub

Private Sub txtValorTransacao_LostFocus()
    If IsNumeric(txtValorTransacao) Then
       txtValorTransacao.Text = Format(txtValorTransacao.Text, "standard")
    End If
End Sub
