VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form lDespInquilino 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de despesas do inquilino"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   ControlBox      =   0   'False
   Icon            =   "lDespInquilino.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OpCondominio 
      Caption         =   "Selecionar por Condomínio"
      Height          =   195
      Left            =   2385
      TabIndex        =   9
      Top             =   120
      Width           =   2220
   End
   Begin VB.OptionButton OpInquilino 
      Caption         =   "Selecionar por Inquilino"
      Height          =   195
      Left            =   315
      TabIndex        =   8
      Top             =   105
      Value           =   -1  'True
      Width           =   1950
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5025
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   315
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5025
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   465
      Width           =   435
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   465
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DB_desp 
      Bindings        =   "lDespInquilino.frx":000C
      Height          =   4020
      Left            =   60
      OleObjectBlob   =   "lDespInquilino.frx":0020
      TabIndex        =   3
      Top             =   885
      Width           =   8955
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\PROGVB60\Predio\Dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6345
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4275
      Visible         =   0   'False
      Width           =   1845
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Top             =   465
      Width           =   5355
      _ExtentX        =   9446
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
      Left            =   1080
      TabIndex        =   0
      Top             =   465
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
   Begin VB.Label lblSelecionar 
      AutoSize        =   -1  'True
      Caption         =   "Inquilino"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   525
      Width           =   585
   End
End
Attribute VB_Name = "lDespInquilino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tbAssociados As Recordset
Private tbCondominio As Recordset

Private Sub cmdCancelar_Click()
  RetCodigo = 0
  Unload Me
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  If OpInquilino.Value Then
    lAssociado.Show 1
    If RetCodigo > 0 Then
      With tbAssociados
        .Index = "codigoid"
        .Seek "=", RetCodigo
        If Not .NoMatch Then
          cpNome.Text = NomeCompleto(!Codigo)
          cpCodigo.Text = !Codigo
        End If
      End With
      RetCodigo = 0
    End If
  Else
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
      RetCodigo = 0
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  If DB_desp.ApproxCount > 0 Then
    Data1.Recordset.Bookmark = DB_desp.Bookmark
    RetCodigo = Data1.Recordset!id
  Else
    RetCodigo = 0
  End If
  Unload Me
End Sub

Private Sub cmdSelecionar_Click()
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um inquilino/condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  With Data1
    If OpInquilino.Value Then
      .RecordSource = "Select * From porfora Where associado = " _
        & cpCodigo.Text & " Order By mes;"
      .Refresh
    Else
      .RecordSource = "Select porfora.* From porfora Where associado in " _
        & " (select codigo from associados where condominio = " _
        & cpCodigo.Text & ") Order By mes;"
      .Refresh
    End If
    DB_desp.ReBind
    If .Recordset.RecordCount > 0 Then
      .Recordset.MoveLast
    End If
  End With
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      If OpInquilino.Value Then
        With tbAssociados
          .Index = "codigoid"
          .Seek "=", cpCodigo.Text
          If Not .NoMatch Then
            cpNome.Text = NomeCompleto(!Codigo)
            cmdSelecionar.SetFocus
          Else
            MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
            cpCodigo.SetFocus
          End If
        End With
      Else
        With tbCondominio
          .Index = "codigoid"
          .Seek "=", cpCodigo.Text
          If Not .NoMatch Then
            cpNome.Text = !Nome
            cmdSelecionar.SetFocus
          Else
            MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
            cpCodigo.SetFocus
          End If
        End With
      End If
    Else
      RetCodigo = 0
      If OpInquilino.Value Then
        lAssociado.Show 1
        If RetCodigo > 0 Then
          With tbAssociados
            .Index = "codigoid"
            .Seek "=", RetCodigo
            If Not .NoMatch Then
              cpNome.Text = NomeCompleto(!Codigo)
              cpCodigo.Text = !Codigo
            End If
          End With
          cmdSelecionar.SetFocus
        End If
      Else
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
          cmdSelecionar.SetFocus
        End If
      End If
    End If
  End If
End Sub

Private Sub DB_desp_DblClick()
  Data1.Recordset.Bookmark = DB_desp.Bookmark
  RetCodigo = Data1.Recordset!id
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Data1.DatabaseName = Parametros.dados
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  KeyPreview = True
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbAssociados = Nothing
  Set tbCondominio = Nothing
End Sub

Private Sub OpCondominio_Click()
  If OpCondominio.Value Then
    lblSelecionar.Caption = "Condomínio"
  Else
    lblSelecionar.Caption = "Inquilino"
  End If
End Sub

Private Sub OpInquilino_Click()
  If OpInquilino.Value Then
    lblSelecionar.Caption = "Inquilino"
  Else
    lblSelecionar.Caption = "Condomínio"
  End If
End Sub
