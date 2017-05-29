VERSION 5.00
Begin VB.Form Usuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de usuários"
   ClientHeight    =   2430
   ClientLeft      =   1650
   ClientTop       =   2340
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4845
      TabIndex        =   17
      ToolTipText     =   "Fecha o cadastro."
      Top             =   2025
      Width           =   1560
   End
   Begin VB.CommandButton cmdPermissao 
      Caption         =   "&Permissões"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4845
      TabIndex        =   12
      ToolTipText     =   "Configura as permissões para este usuário."
      Top             =   1710
      Width           =   1560
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1830
      TabIndex        =   10
      ToolTipText     =   "Salva as informacões da tela no banco de dados."
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox cpCodigo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2775
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   645
   End
   Begin VB.CommandButton cmdUltimo 
      Height          =   315
      Left            =   5925
      Picture         =   "Usuarios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Último registro."
      Top             =   1140
      Width           =   495
   End
   Begin VB.CommandButton cmdProximo 
      Height          =   315
      Left            =   5925
      Picture         =   "Usuarios.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Próximo registro."
      Top             =   825
      Width           =   495
   End
   Begin VB.CommandButton cmdAnterior 
      Height          =   315
      Left            =   5925
      Picture         =   "Usuarios.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Registro anterior."
      Top             =   510
      Width           =   495
   End
   Begin VB.CommandButton cmdPrimeiro 
      Height          =   315
      Left            =   5925
      Picture         =   "Usuarios.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Primeiro registro."
      Top             =   195
      Width           =   495
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2685
      TabIndex        =   11
      ToolTipText     =   "Elimina o usuáio do cadastro."
      Top             =   1740
      Width           =   855
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   975
      TabIndex        =   9
      ToolTipText     =   "Permite a alteração dos dados do usuário atual."
      Top             =   1740
      Width           =   855
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "&Incluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Permite a inclusão de um novo usuário."
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox cpConfSenha 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2775
      Locked          =   -1  'True
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1035
      Width           =   2310
   End
   Begin VB.TextBox cpSenha 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2775
      Locked          =   -1  'True
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   720
      Width           =   2310
   End
   Begin VB.TextBox cpNome 
      Height          =   300
      Left            =   2775
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   3
      Top             =   405
      Width           =   2310
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Có&digo"
      Height          =   195
      Left            =   2220
      TabIndex        =   0
      Top             =   195
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Co&nfirmação da senha"
      Height          =   195
      Left            =   1110
      TabIndex        =   6
      Top             =   1125
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sen&ha"
      Height          =   195
      Left            =   2235
      TabIndex        =   4
      Top             =   810
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Id&entificação"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   495
      Width           =   915
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim EscSim  As Boolean
Dim Alterar As Boolean
Dim OldBook As Variant
Dim tbUsuarios As Recordset

Private Sub cmdAdicionar_Click()
On Error GoTo Errado
'  If Corrente.Op2001 = 0 Then
'    MsgBox S_Permissao, vbCritical + vbOKOnly, caption
'  Else
    With tbUsuarios
      OldBook = .Bookmark
      .Index = "CODIGOid"
      If .RecordCount > 0 Then
        .MoveLast
        cpCodigo = !Codigo + 1
      Else
        cpCodigo = "1"
      End If
    End With
    Limpar
    AcertaBotao (True)
    Travar (False)
    Alterar = False
'  End If
Fim:
  Exit Sub
Errado:
  MsgBox Err.Description
  Resume Fim
End Sub

Private Sub cmdAlterar_Click()
'  If !Op2002 = 0 Then
'    MsgBox S_Permissao, vbCritical + vbOKOnly, caption
'  Else
    OldBook = tbUsuarios.Bookmark
    AcertaBotao (True)
    Travar (False)
    Alterar = True
'  End If
End Sub

Private Sub cmdAnterior_Click()
  With tbUsuarios
    If .RecordCount > 0 Then
      .MovePrevious
      If .BOF Then
        .MoveNext
        If .EOF Then
          Limpar
        Else
          Mostra
        End If
      Else
        Mostra
      End If
    End If
  End With
End Sub

Private Sub cmdExcluir_Click()
'  If Corrente = 0 Then
'    MsgBox S_Permissao, vbCritical + vbOKOnly, caption
'  Else
    Dim Resp As VbMsgBoxResult
    Resp = MsgBox("Você quer mesmo deletar " & cpNome.Text & "?", vbQuestion + vbYesNo, Caption)
    If Resp = vbYes Then
      With tbUsuarios
        If .RecordCount = 1 Then
          MsgBox "É o único usuário do sistema. Não pode ser deletado!", vbInformation + vbOKOnly, Caption
          .Edit
          !Op2001 = 1
          !Op2002 = 1
          !Op2003 = 1
          .Update
        Else
          .Delete
          .MovePrevious
          If .BOF Then
            .MoveNext
            If .EOF Then
              Limpar
            Else
              Mostra
            End If
          Else
            Mostra
          End If
        End If
      End With
    End If
'  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdGravar_Click()
  Resp = MsgBox("Você confirma os dados?", vbQuestion + vbYesNo, Caption)
  If Resp = vbYes Then
    If cpSenha = cpConfSenha Then
      If Alterar = True Then
        Regravar
      Else
        Gravar
      End If
      AcertaBotao (False)
      Travar (True)
      tbUsuarios.Bookmark = tbUsuarios.LastModified
      Mostra
    Else
      MsgBox "A senha e a confirmação não estão iguais", vbCritical + vbOKOnly, Caption
      cpSenha.SetFocus
    End If
  End If
End Sub

Private Sub cmdPermissao_Click()
Dinovo:
  Supervisor.Show vbModal
  If sSuper = SenhaSupervisor Then
    NivelDoUsuario.Label1.Caption = cpCodigo.Text & " - " & cpNome.Text
    NivelDoUsuario.Show
  ElseIf sSuper <> "" Then
    MsgBox "A senha que você informou não está correta! Verifique por favor e tente novamente.", vbCritical + vbOKOnly, Caption
    GoTo Dinovo
  End If
End Sub

Private Sub cmdPrimeiro_Click()
  With tbUsuarios
    If .RecordCount > 0 Then
      .MoveFirst
      If Not .BOF Then
        Mostra
      Else
        Limpar
      End If
    End If
  End With
End Sub

Private Sub cmdProximo_Click()
  With tbUsuarios
    If .RecordCount > 0 Then
      .MoveNext
      If .EOF Then
        .MovePrevious
        If .BOF Then
          Limpar
        Else
          Mostra
        End If
      Else
        Mostra
      End If
    End If
  End With
End Sub

Private Sub cmdUltimo_Click()
  With tbUsuarios
    If .RecordCount > 0 Then
      .MoveLast
      If .EOF Then
        Limpar
      Else
        Mostra
      End If
    End If
  End With
End Sub

Private Sub cpConfSenha_GotFocus()
  cpConfSenha.SelStart = 0
  cpConfSenha.SelLength = Len(cpConfSenha.Text)
End Sub

Private Sub cpNome_GotFocus()
  cpNome.SelStart = 0
  cpNome.SelLength = Len(cpNome.Text)
End Sub

Private Sub cpSenha_GotFocus()
  cpSenha.SelStart = 0
  cpSenha.SelLength = Len(cpSenha.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    If KeyCode = 27 Then
      KeyCode = 0
      If EscSim = True Then
        AcertaBotao (False)
        Travar (True)
        With tbUsuarios
          If .RecordCount > 0 Then
            .Bookmark = OldBook
            Mostra
          End If
        End With
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  KeyPreview = True
  Set tbUsuarios = db.OpenRecordset("Usuarios", dbOpenTable)
  Refresh
  With tbUsuarios
    .Index = "codigoid"
    If .RecordCount > 0 Then
      .MoveFirst
      Mostra
    End If
  End With
End Sub

Private Function Mostra()
  With tbUsuarios
    cpCodigo.Text = IIf(IsNull(!Codigo), "nada", !Codigo)
    cpNome.Text = IIf(IsNull(!Nome), "", Decodifica(!Nome))
    cpSenha.Text = IIf(IsNull(!Senha), "", Decodifica(!Senha))
    cpConfSenha.Text = IIf(IsNull(!Senha), "", Decodifica(!Senha))
  End With
End Function

Private Function Limpar()
  cpNome.Text = ""
  cpSenha.Text = ""
  cpConfSenha.Text = ""
  cpNome.SetFocus
End Function

Private Function Travar(ByVal Acerta As Boolean)
  cpNome.Locked = Acerta
  cpSenha.Locked = Acerta
  cpConfSenha.Locked = Acerta
End Function

Private Function Gravar()
  With tbUsuarios
    .AddNew
    !Codigo = cpCodigo.Text
    !Nome = Codifica(cpNome.Text)
    !Senha = Codifica(cpSenha.Text)
    .Update
  End With
End Function

Private Function Regravar()
  With tbUsuarios
    .Edit
    !Codigo = cpCodigo.Text
    !Nome = Codifica(cpNome.Text)
    !Senha = Codifica(cpSenha.Text)
    .Update
  End With
End Function

Private Sub cpConfSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0: cmdGravar.SetFocus
End Sub

Private Sub cpSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0: cpConfSenha.SetFocus
End Sub

Private Sub cpNome_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0: cpSenha.SetFocus
End Sub

Private Sub AcertaBotao(ByVal Tipo As Boolean)
  cmdAlterar.Enabled = Not Tipo
  cmdAdicionar.Enabled = Not Tipo
  cmdExcluir.Enabled = Not Tipo
  cmdPrimeiro.Enabled = Not Tipo
  cmdAnterior.Enabled = Not Tipo
  cmdProximo.Enabled = Not Tipo
  cmdUltimo.Enabled = Not Tipo
  cmdGravar.Enabled = Tipo
  EscSim = Tipo
  cmdPermissao.Enabled = Not Tipo
  cmdFechar.Enabled = Not Tipo
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbUsuarios = Nothing
End Sub
