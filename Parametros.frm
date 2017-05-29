VERSION 5.00
Begin VB.Form Parame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parâmetros"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "Parametros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cpCorrecao 
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
      Left            =   3405
      MaxLength       =   5
      TabIndex        =   12
      Top             =   1140
      Width           =   765
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   5250
      Picture         =   "Parametros.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   795
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   645
      Left            =   5250
      Picture         =   "Parametros.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   150
      Width           =   1050
   End
   Begin VB.TextBox cpDespesa 
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
      Left            =   1650
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1125
      Width           =   885
   End
   Begin VB.TextBox cpCarencia 
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
      Left            =   3405
      MaxLength       =   5
      TabIndex        =   3
      Top             =   795
      Width           =   765
   End
   Begin VB.TextBox cpJuros 
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
      Left            =   1650
      MaxLength       =   5
      TabIndex        =   2
      Top             =   795
      Width           =   900
   End
   Begin VB.TextBox cpSenha 
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
      IMEMode         =   3  'DISABLE
      Left            =   1650
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   465
      Width           =   1740
   End
   Begin VB.TextBox cpTelefone 
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
      Left            =   1665
      MaxLength       =   30
      TabIndex        =   0
      Top             =   135
      Width           =   3375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Correção"
      Height          =   195
      Left            =   2670
      TabIndex        =   13
      Top             =   1215
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Multa"
      Height          =   195
      Left            =   1170
      TabIndex        =   9
      Top             =   1185
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Carência"
      Height          =   195
      Left            =   2670
      TabIndex        =   8
      Top             =   870
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Juros por mês"
      Height          =   195
      Left            =   570
      TabIndex        =   7
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Senha do supervisor"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   525
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      Height          =   195
      Left            =   945
      TabIndex        =   5
      Top             =   225
      Width           =   630
   End
End
Attribute VB_Name = "Parame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tbParametros As Recordset

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdSalvar_Click()
  With tbParametros
    .MoveFirst
    .Edit
    !Telefone = cpTelefone.Text
    !Supervisor = cpSenha.Text
    !DespesaAd = cpDespesa.Text
    !juros = cpJuros.Text
    !Carencia = cpCarencia.Text
    !Cofins = cpCorrecao.Text
'    !OutrasDesp = cpOutras.Text
    .Update
  End With
  With tbParametros
    .MoveFirst
    Parametros.Carencia = !Carencia
    Parametros.juros = !juros
    Parametros.Multa = !DespesaAd
    Parametros.Periodo = !OutrasDesp
    SenhaSupervisor = !Supervisor
  End With
End Sub

Private Sub cpCarencia_GotFocus()
  cpCarencia.SelStart = 0
  cpCarencia.SelLength = Len(cpCarencia.Text)
End Sub

Private Sub cpCarencia_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpDespesa.SetFocus
    Case 48 To 57
      KeyAscii = KeyAscii
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub cpCarencia_LostFocus()
  If Not IsNumeric(cpCarencia.Text) Then cpCarencia.Text = "0"
  cpCarencia.Text = Format(cpCarencia.Text, "#0")
End Sub

Private Sub cpCorrecao_GotFocus()
  cpCorrecao.SelStart = 0
  cpCorrecao.SelLength = Len(cpCorrecao.Text)
End Sub

Private Sub cpCorrecao_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      KeyAscii = 0
      cpTelefone.SetFocus
    Case Else
      KeyAscii = vMoeda(KeyAscii, cpCorrecao)
  End Select
End Sub

Private Sub cpCorrecao_LostFocus()
  If Not IsNumeric(cpCorrecao.Text) Then cpCorrecao.Text = "0"
  cpCorrecao.Text = Format(cpCorrecao.Text, "#0.00")
End Sub

Private Sub cpDespesa_GotFocus()
  cpDespesa.SelStart = 0
  cpDespesa.SelLength = Len(cpDespesa.Text)
End Sub

Private Sub cpDespesa_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      KeyAscii = 0
      cpCorrecao.SetFocus
    Case Else
      KeyAscii = vMoeda(KeyAscii, cpDespesa)
  End Select
End Sub

Private Sub cpDespesa_LostFocus()
  If Not IsNumeric(cpDespesa.Text) Then cpDespesa.Text = "0"
  cpDespesa.Text = Format(cpDespesa.Text, "#0.00")
End Sub

Private Sub cpJuros_GotFocus()
  cpJuros.SelStart = 0
  cpJuros.SelLength = Len(cpJuros.Text)
End Sub

Private Sub cpJuros_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      KeyAscii = 0
      cpCarencia.SetFocus
    Case Else
      KeyAscii = vMoeda(KeyAscii, cpJuros)
  End Select
End Sub

Private Sub cpJuros_LostFocus()
  If Not IsNumeric(cpJuros.Text) Then cpJuros.Text = "0"
  cpJuros.Text = Format(cpJuros.Text, "#0.00")
End Sub

Private Sub cpSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpJuros.SetFocus
  End If
End Sub

Private Sub cpTelefone_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpSenha.SetFocus
    Case 48 To 57
      KeyAscii = KeyAscii
    Case 32, 40, 41, 45
      KeyAscii = KeyAscii
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbParametros = db.OpenRecordset("Parametros", dbOpenTable)
  Refresh
  KeyPreview = True
  With tbParametros
    .MoveFirst
    cpTelefone.Text = IIf(IsNull(!Telefone), "", !Telefone)
    cpSenha.Text = IIf(IsNull(!Supervisor), "", !Supervisor)
    cpDespesa.Text = IIf(IsNull(!DespesaAd), "", Format(!DespesaAd, "#0"))
    cpJuros.Text = IIf(IsNull(!juros), "", Format(!juros, "#0.00"))
    cpCarencia.Text = IIf(IsNull(!Carencia), "", Format(!Carencia, "#0"))
'    cpOutras.Text = IIf(IsNull(!OutrasDesp), "", Format(!OutrasDesp, "#0"))
    cpCorrecao.Text = IIf(IsNull(!Cofins), "", !Cofins)
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbParametros = Nothing
End Sub
