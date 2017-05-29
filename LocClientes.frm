VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form LocClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Clientes"
   ClientHeight    =   3705
   ClientLeft      =   1650
   ClientTop       =   1665
   ClientWidth     =   6015
   Icon            =   "LocClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2670
      Top             =   1470
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   300
      Left            =   4950
      TabIndex        =   6
      ToolTipText     =   "Confirma."
      Top             =   3330
      Width           =   930
   End
   Begin VB.CommandButton cmdNome 
      Caption         =   "&Nome"
      Height          =   300
      Left            =   3180
      TabIndex        =   5
      ToolTipText     =   "Indexar por nome."
      Top             =   3330
      Width           =   1200
   End
   Begin VB.CommandButton cmdCota 
      Caption         =   "Co&ta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1830
      TabIndex        =   4
      ToolTipText     =   "Indexar por cota."
      Top             =   3330
      Width           =   1200
   End
   Begin VB.CommandButton cmdCodigo 
      Caption         =   "&Código"
      Height          =   300
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Indexar por código."
      Top             =   3330
      Width           =   1200
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LocClientes.frx":000C
      Height          =   2745
      Left            =   15
      OleObjectBlob   =   "LocClientes.frx":0040
      TabIndex        =   2
      Top             =   495
      Width           =   5955
   End
   Begin VB.TextBox cpChave 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1575
      MaxLength       =   45
      TabIndex        =   0
      Top             =   105
      Width           =   3600
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\progvb60\bahia\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Cadastro"
      Top             =   2535
      Width           =   1275
   End
   Begin VB.Label tChave 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Cliente"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   195
      Width           =   1170
   End
End
Attribute VB_Name = "LocClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub cmdCodigo_Click()
  Data1.Recordset.Index = "PerCodigo"
  tChave = "Código do Cliente"
  tChave.Refresh
  cmdCodigo.FontBold = True
  cmdCota.FontBold = False
  cmdNome.FontBold = False
  cpChave.SetFocus
End Sub

Private Sub cmdCota_Click()
  Data1.Recordset.Index = "PerCota"
  tChave = "Número da Cota"
  tChave.Refresh
  cmdCodigo.FontBold = False
  cmdCota.FontBold = True
  cmdNome.FontBold = False
  cpChave.SetFocus
End Sub

Private Sub cmdNome_Click()
  Data1.Recordset.Index = "PerNome"
  tChave = "Nome do Cliente"
  tChave.Refresh
  cmdCodigo.FontBold = False
  cmdCota.FontBold = False
  cmdNome.FontBold = True
  cpChave.SetFocus
End Sub

Private Sub cmdOk_Click()
  RetCodigo = DBGrid1.Columns(0).Text
  Unload Me
End Sub

Private Sub cpChave_Change()
  With Data1.Recordset
    Select Case .Index
      Case "PerNome"
        .Seek ">=", cpChave
      Case "PerCodigo"
        .Seek "<=", Val(cpChave)
      Case "PerCota"
        .Seek ">=", cpChave
    End Select
  End With
End Sub

Private Sub cpChave_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    Select Case KeyCode
      Case 38
        Data1.Recordset.MovePrevious
        DBGrid1.SetFocus
      Case 40
        Data1.Recordset.MoveNext
        DBGrid1.SetFocus
    End Select
  End If
End Sub

Private Sub cpChave_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    DBGrid1.SetFocus
  End If
End Sub

Private Sub DBGrid1_DblClick()
  cmdOk = True
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    RetCodigo = DBGrid1.Columns(0).Text
    Unload Me
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    Select Case KeyCode
      Case 27
        Unload Me
    End Select
  End If
End Sub

Private Sub Form_Load()
  Me.Refresh
  DoEvents
  Data1.DatabaseName = Trim(Parametros.Dados) & "\dados.mdb"
  Timer1.Enabled = True
  Me.KeyPreview = True
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Data1.Recordset.Index = "PerCota"
End Sub
