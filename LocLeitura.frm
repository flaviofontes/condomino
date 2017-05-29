VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form LocLeitura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de leituras"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   ControlBox      =   0   'False
   Icon            =   "LocLeitura.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   480
      Width           =   1275
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   480
      Width           =   1275
      _ExtentX        =   2249
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
      MaxLength       =   7
      TextMask        =   9
      RawText         =   9
      Mask            =   "##/####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   315
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4500
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LocLeitura.frx":000C
      Height          =   3435
      Left            =   60
      OleObjectBlob   =   "LocLeitura.frx":0020
      TabIndex        =   4
      Top             =   960
      Width           =   8415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4500
      Visible         =   0   'False
      Width           =   1695
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
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
      Left            =   1020
      TabIndex        =   0
      Top             =   60
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
      Caption         =   "Mês/ano"
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   540
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "LocLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCod As Long
Dim tbCondominio As Recordset

Public Property Get Retorno() As Long
  Retorno = lCod
End Property

Private Sub cmdCancelar_Click()
  lCod = -1
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

Private Sub cmdOk_Click()
  If Data1.Recordset.RecordCount > 0 Then
    lCod = Data1.Recordset.Fields("ID_LEITURA").Value
  End If
  Unload Me
End Sub

Private Sub cmdSelecionar_Click()
  
  If Not IsDate("25/" & cpMes.Text) Then
    MsgBox "O mês/ano informado não é válido.", vbInformation + vbOKOnly, "Aviso"
    cpMes.SetFocus
    Exit Sub
  End If
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  Data1.RecordSource = "SELECT CONTAS_LEITURA.DATA_LEITURA, CONTAS_LEITURA.ANT_LEITURA, CONTAS_LEITURA.ATUAL_LEITURA, CONTAS_LEITURA.VALOR_LEITURA, TIPO_LEITURA.DESCRICAO, CONTAS_LEITURA.MES_LEITURA, CONTAS_LEITURA.ID_LEITURA " _
    & "FROM TIPO_LEITURA RIGHT JOIN CONTAS_LEITURA ON TIPO_LEITURA.TIPO_LEITURA = CONTAS_LEITURA.TIPO_LEITURA " _
    & "WHERE (((CONTAS_LEITURA.CODIGO_CONDOMINIO)=" & cpCodigo.Text & ") AND ((CONTAS_LEITURA.MES_LEITURA) = '" & cpMes.Text & "')) ORDER BY CONTAS_LEITURA.DATA_LEITURA DESC;"
  Data1.Refresh
  DBGrid1.ReBind

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
          cpMes.SetFocus
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
        cpMes.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSelecionar.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Data1.DatabaseName = Parametros.dados
  KeyPreview = True
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
