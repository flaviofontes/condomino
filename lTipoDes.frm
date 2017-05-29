VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form lTipoDesp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de tipos de despesa"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "lTipoDes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\PROGVB60\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   4950
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "TPDESP"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSDBGrid.DBGrid DB_lista 
      Bindings        =   "lTipoDes.frx":000C
      Height          =   3180
      Left            =   45
      OleObjectBlob   =   "lTipoDes.frx":0020
      TabIndex        =   0
      Top             =   30
      Width           =   7365
   End
End
Attribute VB_Name = "lTipoDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
  RetCodigo = 0
  Unload Me
End Sub

Private Sub cmdOk_Click()
  If Not (DB_lista.Columns(0).Text = "") Then
    RetCodigo = CLng(DB_lista.Columns(0).Text)
    Unload Me
  End If
End Sub

Private Sub DB_lista_DblClick()
  RetCodigo = CLng(DB_lista.Columns(0).Text)
  Unload Me
End Sub

Private Sub Form_Activate()
  Data1.Recordset.Index = "descricaoid"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Refresh
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
End Sub
