VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmConsultaTransacaoCategoria 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta por categoria em todas as transações existentes"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10560
   Icon            =   "frmConsultaTransacaoCategoria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   465
      Left            =   8640
      TabIndex        =   1
      Top             =   5160
      Width           =   1845
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Numero_Cartao"
         Caption         =   "Numero Cartao"
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
         DataField       =   "Valor_Transacao"
         Caption         =   "Valor Transacao"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "standard"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Data_Transacao"
         Caption         =   "Data Transacao"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Descricao"
         Caption         =   "Descricao"
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
         DataField       =   "Categoria"
         Caption         =   "Categoria"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   1
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1590,236
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1679,811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1725,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3915,213
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1665,071
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaTransacaoCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CarregarTransacoesCategoria()

    Dim strSQL As String
    Set rsTemp = CreateObject("ADODB.Recordset")
    MousePointer = 11
    
    strSQL = "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, dbo.fn_CategoriaTransacao(Valor_Transacao) AS Categoria "
    strSQL = strSQL & "FROM Cartao_Transacoes ORDER BY Categoria,Data_Transacao "
    
    If rsTemp.State = 0 Then
        rsTemp.Open strSQL, CN, adOpenStatic, adLockReadOnly
    End If
    DoEvents
    Set DataGrid1.DataSource = rsTemp
    
    If rsTemp.EOF Then
        MsgBox "Nenhum registro encontrado.", vbInformation
        Set DataGrid1.DataSource = Nothing
        DataGrid1.Refresh
        Set rsTemp = Nothing
        Exit Sub
    End If
    
    MousePointer = 0
End Sub


Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Center Me
    CarregarTransacoesCategoria
End Sub
