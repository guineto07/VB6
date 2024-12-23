VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsultaClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Clientes"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11115
   Icon            =   "frmConsultaCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Filtrar consulta"
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   30
      TabIndex        =   5
      Top             =   150
      Width           =   2865
      Begin VB.OptionButton optNome 
         Appearance      =   0  'Flat
         Caption         =   "Nome "
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   1590
         Width           =   1725
      End
      Begin VB.OptionButton optNumero_Cartao 
         Appearance      =   0  'Flat
         Caption         =   "Número Cartão"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.TextBox txtNome 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   150
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1890
         Width           =   2475
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
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   2475
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "&Consultar"
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   3870
      Width           =   2475
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5085
      Left            =   2940
      TabIndex        =   7
      Top             =   240
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   8969
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
      ColumnCount     =   3
      BeginProperty Column00 
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
            ColumnWidth     =   5414,74
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
      TabIndex        =   8
      Top             =   3450
      Width           =   2925
   End
   Begin VB.Label Label2 
      Caption         =   "Para abrir a transação dê um duplo clique "
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2970
      TabIndex        =   9
      Top             =   30
      Width           =   3165
   End
End
Attribute VB_Name = "frmConsultaClientes"
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
    
    strSQL = "SELECT TOP 100 "
    strSQL = strSQL & "Clientes.Id_Cliente,Clientes.Nome,"
    strSQL = strSQL & "Clientes.Numero_Cartao"
    
    strSQL = strSQL & " FROM Clientes "
    
    If optNumero_Cartao.Value = True Then
        
        If Len(Trim(mskNumero_Cartao.Text)) <> 19 Then
            MsgBox "Número de cartão inválido.", vbInformation
            mskNumero_Cartao.SetFocus
            Exit Sub
        End If
        strSQL = strSQL & " WHERE Clientes.Numero_Cartao='" & Replace(mskNumero_Cartao.Text, " ", "") & "'"
    
    ElseIf optNome.Value = True Then
        
        strSQL = strSQL & " WHERE Clientes.Nome LIKE '" & txtNome & "%'"
       
    End If
    
    strSQL = strSQL & " ORDER BY Clientes.Nome"
     
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
              
        Set clsCliente = New clsClientes
        With clsCliente
            .Id_Cliente = DataGrid1.Columns(0).Value
            .strNumero_Cartao = DataGrid1.Columns(1).Value
            .Nome = DataGrid1.Columns(2).Value
        End With
        Unload Me
    End If
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
     Center Me
    Limpar
    DataGrid1.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub mskNumero_Cartao_GotFocus()
    mskNumero_Cartao.SelStart = 0
    mskNumero_Cartao.SelLength = Len(mskNumero_Cartao.Text)

End Sub

Private Sub optNumero_Cartao_Click()
    Limpar
    mskNumero_Cartao.Enabled = True
    mskNumero_Cartao.SetFocus
End Sub

Private Sub optValorTransacao_Click()
    Limpar
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
    
    Set DataGrid1.DataSource = Nothing
    DataGrid1.Refresh
    Set rsTemp = Nothing
End Sub

