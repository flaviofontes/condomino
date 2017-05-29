VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form l_Condominio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de condomínios"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   ControlBox      =   0   'False
   Icon            =   "lCond.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7980
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4500
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "&Ok"
      Height          =   315
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4500
      Width           =   975
   End
   Begin rdActiveText.ActiveText cpChave 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   4560
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   50
      TextCase        =   1
      RawText         =   0
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin MSDBGrid.DBGrid db_Lista 
      Bindings        =   "lCond.frx":000C
      Height          =   4215
      Left            =   45
      OleObjectBlob   =   "lCond.frx":0020
      TabIndex        =   1
      Top             =   45
      Width           =   8955
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\PROGVB60\Predio\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5205
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM CONDOMINIO ORDER BY NOME"
      Top             =   3750
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F3 - Próxima ocorrência"
      Height          =   195
      Left            =   2700
      TabIndex        =   4
      Top             =   4320
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Procurar por Mínimo: 4 caracteres"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "l_Condominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Compare Text

Dim traduz As New Soundex

Private Sub cmdCancelar_Click()
  Form_KeyPress 27
End Sub

Private Sub cmdOk_Click()
  Form_KeyPress (13)
End Sub

Private Sub cpChave_Change()
  '"busca Like '*" & traduz.Soundex(cpChave.Text) & "*'"
  Data1.Recordset.FindFirst "nome Like '*" & cpChave.Text & "*'"
End Sub

Private Sub cpChave_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    db_Lista.SetFocus
  End If
End Sub

Private Sub DB_lista_DblClick()
  RetCodigo = Val(db_Lista.Columns(0).Text)
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  ElseIf KeyAscii = 13 Then
    KeyAscii = 0
    RetCodigo = Val(db_Lista.Columns(0).Text)
    Unload Me
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 114 Then
    Data1.Recordset.FindNext "busca Like '*" & traduz.Soundex(cpChave.Text) & "*'"
    If Data1.Recordset.NoMatch Then
      MsgBox "Não existe mais ocorrência para '" & cpChave.Text & "'.", vbInformation + vbOKOnly, "Aviso"
    End If
  End If
End Sub

Private Sub Form_Load()
  Refresh
  Data1.DatabaseName = Parametros.dados
  KeyPreview = True
End Sub
