VERSION 5.00
Begin VB.Form ServidorEmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro do servidor de e-mail"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "ServidorEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7080
      Top             =   1440
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   120
      Width           =   1035
   End
   Begin VB.CheckBox chSsl 
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   2340
      Width           =   195
   End
   Begin VB.CheckBox chAut 
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   1980
      Width           =   195
   End
   Begin VB.TextBox cpSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox cpUsuario 
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox cpPorta 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
   Begin VB.TextBox cpServidor 
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox cpEmail 
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Usar segurança SSL"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   2340
      Width           =   1470
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Usar altenticação segura"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   1980
      Width           =   1770
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   1620
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nome de usuário"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1260
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Porta"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Servidor de e-mail"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   540
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E-mail"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "ServidorEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As Recordset

Private Sub cmdSalvar_Click()
  
  If cpEmail.Text = "" Then
    MsgBox "Todos os campos são obrigatórios.", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  If cpServidor.Text = "" Then
    MsgBox "Todos os campos são obrigatórios.", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  If cpPorta.Text = "" Then
    MsgBox "Todos os campos são obrigatórios.", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  If cpUsuario.Text = "" Then
    MsgBox "Todos os campos são obrigatórios.", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  If cpSenha.Text = "" Then
    MsgBox "Todos os campos são obrigatórios.", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      .Edit
    Else
      .AddNew
    End If
    !email = cpEmail.Text
    !servidor = cpServidor.Text
    !porta = cpPorta.Text
    !usuario = cpUsuario.Text
    !Senha = cpSenha.Text
    !usaraut = chAut.Value
    !usarssl = chSsl.Value
    .Update
  End With
  Unload Me
End Sub

Private Sub cpEmail_GotFocus()
  cpEmail.SelStart = 0
  cpEmail.SelLength = Len(cpEmail.Text)
End Sub

Private Sub cpEmail_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpServidor.SetFocus
  End If
End Sub

Private Sub cpPorta_GotFocus()
  cpPorta.SelStart = 0
  cpPorta.SelLength = Len(cpPorta.Text)
End Sub

Private Sub cpPorta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpUsuario.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpSenha_GotFocus()
  cpSenha.SelStart = 0
  cpSenha.SelLength = Len(cpSenha.Text)
End Sub

Private Sub cpSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    chAut.SetFocus
  End If
End Sub

Private Sub cpServidor_GotFocus()
  cpServidor.SelStart = 0
  cpServidor.SelLength = Len(cpServidor.Text)
End Sub

Private Sub cpServidor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpPorta.SetFocus
  End If
End Sub

Private Sub cpUsuario_GotFocus()
  cpUsuario.SelStart = 0
  cpUsuario.SelLength = Len(cpUsuario.Text)
End Sub

Private Sub cpUsuario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpSenha.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Refresh
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Set rs = db.OpenRecordset("emailserver", dbOpenTable)
  With rs
    If .RecordCount > 0 Then
      cpEmail.Text = !email
      cpServidor.Text = !servidor
      cpPorta.Text = !porta
      cpUsuario.Text = !usuario
      cpSenha.Text = !Senha
      chAut.Value = !usaraut
      chSsl.Value = !usarssl
    End If
  End With
End Sub
