VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Tipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3450
   Icon            =   "Tipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DB_Tipos 
      Bindings        =   "Tipos.frx":000C
      Height          =   3105
      Left            =   15
      OleObjectBlob   =   "Tipos.frx":0020
      TabIndex        =   0
      Top             =   30
      Width           =   3405
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\PROGVB60\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Tipos"
      Top             =   1980
      Visible         =   0   'False
      Width           =   1845
   End
End
Attribute VB_Name = "Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DB_Tipos_AfterUpdate()
  DB_Tipos.Refresh
End Sub

Private Sub DB_Tipos_BeforeDelete(Cancel As Integer)
  Resp = MsgBox("Desletar este tipo?", vbQuestion + vbYesNo + vbDefaultButton2, Caption)
  If Resp = vbNo Then
    Cancel = True
  End If
End Sub

Private Sub DB_Tipos_Error(ByVal DataError As Integer, Response As Integer)
  MsgBox Error$(DataError), vbCritical + vbOKOnly, Caption
  Response = 0
End Sub

Private Sub DB_Tipos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 13
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

'Private Sub Form_Initialize()
'  Data1.Recordset.Index = "descricaoid"
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Refresh
  Data1.DatabaseName = Parametros.Dados
  KeyPreview = True
End Sub
