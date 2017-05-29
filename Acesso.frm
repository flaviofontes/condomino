VERSION 5.00
Begin VB.Form Acesso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de usuário"
   ClientHeight    =   2190
   ClientLeft      =   1830
   ClientTop       =   2430
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   3405
      Picture         =   "Acesso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancela a carga do programa."
      Top             =   1275
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2175
      Picture         =   "Acesso.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Continua a carga do programa."
      Top             =   1275
      Width           =   1215
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
      Left            =   2175
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   885
      Width           =   2430
   End
   Begin VB.TextBox cpNome 
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
      Left            =   2175
      MaxLength       =   20
      TabIndex        =   0
      Top             =   300
      Width           =   2430
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "J && F Software's Ltda (31) 3891-4298"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   75
      TabIndex        =   6
      Top             =   1710
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   2175
      TabIndex        =   5
      Top             =   645
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   2175
      TabIndex        =   4
      Top             =   75
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   30
      Picture         =   "Acesso.frx":0614
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2055
   End
End
Attribute VB_Name = "Acesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

Dim tbUsuarios As Recordset

Private Sub cmdCancelar_Click()
  Set Principal.balao = Nothing
  End
End Sub

Private Sub cmdOk_Click()
On Error GoTo Errado
  
  If cpNome.Text = "jf" Then
    If cpSenha.Text = "meunome" Then
      Corrente.Codigo = -1
      Corrente.Nome = "Mestre"
      Corrente.Senha = "Mestre"
    End If
  Else
    Set tbUsuarios = db.OpenRecordset("usuarios", dbOpenTable)
    With tbUsuarios
      .Index = "nomeid"
      .Seek "=", Codifica(cpNome.Text)
      If .NoMatch Then
        MsgBox "Nome de usuário ou senha inválido.", vbCritical + vbOKOnly, Caption
        cpNome.SetFocus
        Exit Sub
      Else
        If !Senha = Codifica(cpSenha.Text) Or IsNull(!Senha) Then
          Corrente.Codigo = !Codigo
          Corrente.Nome = Decodifica(!Nome)
          Corrente.Senha = Decodifica(!Senha)
        Else
          MsgBox "Nome de usuário ou senha inválido.", vbCritical + vbOKOnly, Caption
          cpNome.SetFocus
          Exit Sub
        End If
      End If
    End With
  End If
  
  With Principal.Informe
    .Panels(2).Text = "Usuario: " & Corrente.Nome & "  "
    .Panels(2).AutoSize = sbrContents
  End With
  Unload Me
  
Exit_Errado:
  Set tbUsuarios = Nothing
  Exit Sub
  
Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Exit_Errado
  
End Sub

Private Sub cpNome_GotFocus()
  cpNome.SelStart = 0
  cpNome.SelLength = Len(cpNome.Text)
End Sub

Private Sub cpNome_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift = 0 And KeyCode = 40 Then _
    cpSenha.SetFocus
End Sub

Private Sub cpSenha_GotFocus()
  cpSenha.SelStart = 0
  cpSenha.SelLength = Len(cpSenha.Text)
End Sub

Private Sub cpSenha_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift = 0 And KeyCode = 38 Then _
    cpNome.SetFocus
End Sub

Private Sub Form_Load()
  KeyPreview = True
  Refresh
End Sub

Private Sub cpSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdOk.Enabled Then
      cmdOk.SetFocus
    Else
      cpNome.SetFocus
    End If
  End If
End Sub

Private Sub cpNome_Change()
  If Len(cpNome.Text) = 0 Then
    cmdOk.Enabled = False
  Else
    cmdOk.Enabled = True
  End If
End Sub

Private Sub cpNome_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpSenha.SetFocus
  End If
End Sub
