VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form LocalizaBoleto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar boleto"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   ControlBox      =   0   'False
   Icon            =   "LocalizaBoleto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   315
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   375
      Left            =   7500
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1035
      Top             =   3630
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LocalizaBoleto.frx":000C
      Height          =   3945
      Left            =   45
      OleObjectBlob   =   "LocalizaBoleto.frx":0020
      TabIndex        =   3
      Top             =   600
      Width           =   8745
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programas desenvolvidos\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1410
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   4995
      _ExtentX        =   8811
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
      Left            =   1140
      TabIndex        =   0
      Top             =   120
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "LocalizaBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Situacao As String
Dim tbCondominio As Recordset

Private Sub cmdCancelar_Click()
  RetNome = ""
  RetCodigo = 0
  Unload Me
End Sub

Private Sub cmdLocalizar_Click()
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
  End If
End Sub

Private Sub cmdPrint_Click()
  If DBGrid1.Text <> "" Then
    RetNome = DBGrid1.Columns(1).Text
    RetCodigo = cpCodigo.Text
  End If
  Unload Me
End Sub

Private Sub cmdSelecionar_Click()
  If Val(cpCodigo.Text) > 0 Then
    Selecionar cpCodigo.Text
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
          cmdSelecionar.SetFocus
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
        cmdSelecionar.SetFocus
      End If
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
End Sub

Private Sub Selecionar(ByVal nCod As Long)
  If Situacao = "" Then
    Data1.RecordSource = "SELECT Boletos.*, Blocos.nome_bloco+' '+[associados.tipo]+' '+[associados.apartamento]+' '+[boletos.nome] AS nomecompleto " _
        & "FROM BLOCOS RIGHT JOIN (ASSOCIADOS LEFT JOIN Boletos ON ASSOCIADOS.CODIGO = Boletos.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO " _
        & "WHERE (((Boletos.COND)=" & nCod & ")) and BOLETOS.CANCELADO = 'N' ORDER BY Blocos.nome_bloco, ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO, Boletos.VCTO;"
  Else
    Data1.RecordSource = "SELECT Boletos.*, Blocos.nome_bloco+' '+[associados.tipo]+' '+[associados.apartamento]+' '+[boletos.nome] AS nomecompleto " _
      & "FROM BLOCOS RIGHT JOIN (ASSOCIADOS LEFT JOIN Boletos ON ASSOCIADOS.CODIGO = Boletos.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO " _
      & "WHERE (((Boletos.COND)=" & nCod & ") AND ((Boletos.PAGO)='" & Situacao & "')) and BOLETOS.CANCELADO = 'N' ORDER BY Blocos.nome_bloco, ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO, Boletos.VCTO;"
  End If
  Data1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
