VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form lBlocos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de blocos"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "lBlocos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   4950
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "BLOCOS"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSDBGrid.DBGrid DB_lista 
      Bindings        =   "lBlocos.frx":000C
      Height          =   3180
      Left            =   45
      OleObjectBlob   =   "lBlocos.frx":0020
      TabIndex        =   0
      Top             =   30
      Width           =   7365
   End
End
Attribute VB_Name = "lBlocos"
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
  Data1.Recordset.Index = "ID_BLOCO"
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
