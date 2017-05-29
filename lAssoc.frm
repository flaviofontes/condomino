VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form lAssociado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de inquilinos"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   Icon            =   "lAssoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4380
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocCond 
      Caption         =   "..."
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   4500
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   2940
      Top             =   2160
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   315
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4380
      Width           =   1095
   End
   Begin rdActiveText.ActiveText cpChave 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   4080
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   50
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDBGrid.DBGrid DB_lista 
      Bindings        =   "lAssoc.frx":000C
      Height          =   3915
      Left            =   45
      OleObjectBlob   =   "lAssoc.frx":0020
      TabIndex        =   1
      Top             =   75
      Width           =   9795
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5865
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3225
      Visible         =   0   'False
      Width           =   1710
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   4500
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   50
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   4500
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filtrar"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Procurar por"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   4140
      Width           =   870
   End
End
Attribute VB_Name = "lAssociado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nCond As Long
Dim nCod As Integer
Dim tbCondominio As Recordset

Private Sub cmdCancelar_Click()
  RetCodigo = 0
  Unload Me
End Sub

Private Sub cmdLocCond_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpNome.Text = !Nome
        cpCodigo.Text = !Codigo
        Seleciona
      End If
    End With
  End If
End Sub

Private Sub cmdOk_Click()
  DB_lista_DblClick
End Sub

Private Sub cpChave_Change()
  Select Case nCod
    Case 0
      Data1.Recordset.FindFirst "codigo >= " & Val(SoNumeros(cpChave.Text))
    Case 1
      Data1.Recordset.FindFirst "mnome like '*" & cpChave.Text & "*'"
  End Select
End Sub

Private Sub cpChave_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    DB_lista.SetFocus
  End If
End Sub

Private Sub Seleciona()
  nCond = cpCodigo.Text
  If nCond <= 0 Then
    Select Case nCod
      Case 0
        Data1.RecordSource = "SELECT Associados.CODIGO, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS MNOME, BLOCOS.NOME_BLOCO, CONDOMINIO.NOME " _
            & "FROM (BLOCOS RIGHT JOIN Associados ON BLOCOS.ID_BLOCO = Associados.ID_BLOCO) LEFT JOIN CONDOMINIO ON Associados.CONDOMINIO = CONDOMINIO.CODIGO " _
            & "ORDER BY Associados.CODIGO;"
        Data1.Refresh
      Case 2
        Data1.RecordSource = "SELECT Associados.CODIGO, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS MNOME, BLOCOS.NOME_BLOCO, CONDOMINIO.NOME " _
            & "FROM (BLOCOS RIGHT JOIN Associados ON BLOCOS.ID_BLOCO = Associados.ID_BLOCO) LEFT JOIN CONDOMINIO ON Associados.CONDOMINIO = CONDOMINIO.CODIGO " _
            & "ORDER BY ASSOCIADOS.CONDOMINIO, ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO;"
        Data1.Refresh
    End Select
    DB_lista.ReBind
  Else
    Select Case nCod
      Case 0
        Data1.RecordSource = "SELECT Associados.CODIGO, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS MNOME, BLOCOS.NOME_BLOCO, CONDOMINIO.NOME " _
            & "FROM (BLOCOS RIGHT JOIN Associados ON BLOCOS.ID_BLOCO = Associados.ID_BLOCO) LEFT JOIN CONDOMINIO ON Associados.CONDOMINIO = CONDOMINIO.CODIGO " _
            & "Where (((Associados.Condominio) = " & nCond & ")) ORDER BY Associados.CODIGO;"
        Data1.Refresh
      Case 2
        Data1.RecordSource = "SELECT Associados.CODIGO, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS MNOME, BLOCOS.NOME_BLOCO, CONDOMINIO.NOME " _
            & "FROM (BLOCOS RIGHT JOIN Associados ON BLOCOS.ID_BLOCO = Associados.ID_BLOCO) LEFT JOIN CONDOMINIO ON Associados.CONDOMINIO = CONDOMINIO.CODIGO " _
            & "Where (((Associados.Condominio) = " & nCond & ")) ORDER BY ASSOCIADOS.CONDOMINIO, ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO;"
        Data1.Refresh
    End Select
    DB_lista.ReBind
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodigo.Text
        If Not .NoMatch Then
          cpNome.Text = !Nome
          cpChave.SetFocus
        Else
          MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
          cpCodigo.SetFocus
        End If
      End With
    Else
      RetCodigo = 0
      l_Condominio.Show 1
      If RetCodigo > 0 Then
        With tbCondominio
          .Index = "codigoid"
          .Seek "=", RetCodigo
          If Not .NoMatch Then
            cpNome.Text = !Nome
            cpCodigo.Text = !Codigo
          End If
        End With
        cpChave.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpCodigo_LostFocus()
  Seleciona
End Sub

Private Sub DB_lista_DblClick()
  RetCodigo = Val(DB_lista.Columns(0).Text)
  Unload Me
End Sub

Private Sub DB_lista_HeadClick(ByVal ColIndex As Integer)
  nCod = ColIndex
  Seleciona
  DB_lista.ClearSelCols
  cpChave.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Data1.RecordSource = "SELECT Associados.CODIGO, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS MNOME, BLOCOS.NOME_BLOCO, CONDOMINIO.NOME " _
      & "FROM (BLOCOS RIGHT JOIN Associados ON BLOCOS.ID_BLOCO = Associados.ID_BLOCO) LEFT JOIN CONDOMINIO ON Associados.CONDOMINIO = CONDOMINIO.CODIGO " _
      & "ORDER BY Associados.CODIGO;"
  Data1.Refresh
  nCod = 2
End Sub
