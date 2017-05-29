VERSION 5.00
Begin VB.Form ConfBoleto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração dos Boletos"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   Icon            =   "CongBoleto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cpDigAgencia 
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
      Left            =   2985
      MaxLength       =   1
      TabIndex        =   1
      Top             =   165
      Width           =   255
   End
   Begin VB.TextBox cpOperacao 
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
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1845
      Width           =   630
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   795
      Left            =   4215
      Picture         =   "CongBoleto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   945
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirma"
      Height          =   795
      Left            =   4215
      Picture         =   "CongBoleto.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   150
      Width           =   1095
   End
   Begin VB.TextBox cpConta 
      Alignment       =   1  'Right Justify
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
      Left            =   2490
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1500
      Width           =   1380
   End
   Begin VB.TextBox cpCarteira 
      Alignment       =   1  'Right Justify
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
      Left            =   2490
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1170
      Width           =   360
   End
   Begin VB.TextBox cpAgencia 
      Alignment       =   1  'Right Justify
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
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   3
      Top             =   825
      Width           =   960
   End
   Begin VB.TextBox cpMoeda 
      Alignment       =   1  'Right Justify
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
      Left            =   2490
      MaxLength       =   1
      TabIndex        =   2
      Top             =   495
      Width           =   270
   End
   Begin VB.TextBox cpIdent 
      Alignment       =   1  'Right Justify
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
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   0
      Top             =   165
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Código da Operação"
      Height          =   195
      Left            =   840
      TabIndex        =   14
      Top             =   1920
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Identificação do Banco"
      Height          =   195
      Left            =   660
      TabIndex        =   11
      Top             =   210
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código da Moeda"
      Height          =   195
      Left            =   1050
      TabIndex        =   10
      Top             =   555
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código da Agência (Sem Digito)"
      Height          =   195
      Left            =   45
      TabIndex        =   9
      Top             =   870
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código da Carteira"
      Height          =   195
      Left            =   1005
      TabIndex        =   8
      Top             =   1224
      Width           =   1305
   End
   Begin VB.Label Conta 
      AutoSize        =   -1  'True
      Caption         =   "Conta do Cedente (Sem Digito)"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1575
      Width           =   2190
   End
End
Attribute VB_Name = "ConfBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirma_Click()
  If MsgBox("Confirma as Informa Para Impressão do Boleto?", vbQuestion + vbYesNo, Titulo) = vbYes Then
    WritePrivateProfileString "BOLETO", "INDET", CStr(cpIdent.Text), App.Path & "\PAAC.CFG"
    WritePrivateProfileString "BOLETO", "MOEDA", CStr(cpMoeda.Text), App.Path & "\PAAC.CFG"
    WritePrivateProfileString "BOLETO", "AGENCIA", CStr(StrZero(cpAgencia.Text, 4)), App.Path & "\PAAC.CFG"
    WritePrivateProfileString "BOLETO", "CARTEIRA", CStr(cpCarteira.Text), App.Path & "\PAAC.CFG"
    WritePrivateProfileString "BOLETO", "CONTA", CStr(StrZero(cpConta.Text, 8)), App.Path & "\PAAC.CFG"
    WritePrivateProfileString "BOLETO", "DIGIDENT", CStr(cpDigAgencia.Text), App.Path & "\PAAC.CFG"
    WritePrivateProfileString "BOLETO", "OPERACAO", CStr(StrZero(cpOperacao.Text, 3)), App.Path & "\PAAC.CFG"
    vbIdent = cpIdent.Text
    vbMoeda = cpMoeda.Text
    vbAgCedente = cpAgencia.Text
    vbCarteira = cpCarteira.Text
    vbConta = cpConta.Text
    vbOperacao = cpOperacao.Text
    DigAgencia = cpDigAgencia.Text
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cpAgencia_GotFocus()
  cpAgencia.SelStart = 0
  cpAgencia.SelLength = Len(cpAgencia.Text)
End Sub

Private Sub cpAgencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCarteira.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpCarteira_GotFocus()
  cpCarteira.SelStart = 0
  cpCarteira.SelLength = Len(cpCarteira.Text)
End Sub

Private Sub cpCarteira_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpConta.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpConta_GotFocus()
  cpConta.SelStart = 0
  cpConta.SelLength = Len(cpConta.Text)
End Sub

Private Sub cpConta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpOperacao.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpDigAgencia_GotFocus()
  cpDigAgencia.SelStart = 0
  cpDigAgencia.SelLength = Len(cpDigAgencia.Text)
End Sub

Private Sub cpDigAgencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMoeda.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpDigAgencia_LostFocus()
  If Not IsNumeric(cpDigAgencia.Text) Then cpDigAgencia.Text = "0"
End Sub

Private Sub cpIdent_GotFocus()
  cpIdent.SelStart = 0
  cpIdent.SelLength = Len(cpIdent.Text)
End Sub

Private Sub cpIdent_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDigAgencia.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpMoeda_GotFocus()
  cpMoeda.SelStart = 0
  cpMoeda.SelLength = Len(cpMoeda.Text)
End Sub

Private Sub cpMoeda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpAgencia.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpOperacao_GotFocus()
  cpOperacao.SelStart = 0
  cpOperacao.SelLength = Len(cpOperacao.Text)
End Sub

Private Sub cpOperacao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdConfirma.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpOperacao_LostFocus()
  cpOperacao.Text = StrZero(cpOperacao.Text, 3)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then KeyAscii = 0: Unload Me
End Sub

Private Sub Form_Load()
  Me.KeyPreview = True
  Me.Refresh
  DoEvents
  cpIdent.Text = vbIdent
  cpMoeda.Text = vbMoeda
  cpAgencia.Text = vbAgCedente
  cpCarteira.Text = vbCarteira
  cpConta.Text = vbConta
  cpDigAgencia.Text = DigAgencia
  cpOperacao.Text = vbOperacao
End Sub
