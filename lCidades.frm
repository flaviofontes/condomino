VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.3#0"; "ActiveText.ocx"
Begin VB.Form lCidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de cidades"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "lCidades.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "lCidades.frx":000C
      Height          =   3045
      Left            =   45
      OleObjectBlob   =   "lCidades.frx":0020
      TabIndex        =   2
      Top             =   525
      Width           =   6180
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\PROGVB60\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3645
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Cidades"
      Top             =   2475
      Visible         =   0   'False
      Width           =   2160
   End
   Begin rdActiveText.ActiveText cpChave 
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      Top             =   90
      Width           =   4800
      _ExtentX        =   8467
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Procurar..."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "lCidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cpChave_Change()
  Data1.Recordset.Seek ">=", cpChave.Text
End Sub

Private Sub DBGrid1_DblClick()
  RetCidade(0) = DBGrid1.Columns(0).Text
  RetCidade(1) = DBGrid1.Columns(1).Text
  RetCidade(2) = DBGrid1.Columns(2).Text
  Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Select Case ColIndex
    Case 0
      Data1.Recordset.Index = "nome"
    Case 1
      Data1.Recordset.Index = "estado"
    Case 2
      Data1.Recordset.Index = "cep"
  End Select
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Not (DBGrid1.Text = "") Then
      RetCidade(0) = DBGrid1.Columns(0).Text
      RetCidade(1) = DBGrid1.Columns(1).Text
      RetCidade(2) = DBGrid1.Columns(2).Text
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Activate()
  Data1.Recordset.Index = "Nome"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Refresh
  Data1.DatabaseName = Parametros.Dados
  KeyPreview = True
End Sub
