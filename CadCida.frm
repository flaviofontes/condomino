VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form CadCidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de cidades"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ControlBox      =   0   'False
   Icon            =   "CadCida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6345
      Picture         =   "CadCida.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   4995
      Picture         =   "CadCida.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4005
      Picture         =   "CadCida.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3015
      Picture         =   "CadCida.frx":092A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2025
      Picture         =   "CadCida.frx":0C34
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1035
      Picture         =   "CadCida.frx":0F3E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   45
      Picture         =   "CadCida.frx":1248
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   30
      Width           =   990
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   60
      ScaleHeight     =   0
      ScaleWidth      =   7215
      TabIndex        =   6
      Top             =   960
      Width           =   7275
   End
   Begin rdActiveText.ActiveText cpCep 
      Height          =   315
      Left            =   2430
      TabIndex        =   2
      Top             =   1515
      Width           =   1335
      _ExtentX        =   2355
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
      MaxLength       =   9
      TextMask        =   6
      RawText         =   6
      Mask            =   "#####-###"
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpEstado 
      Height          =   315
      Left            =   1125
      TabIndex        =   1
      Top             =   1515
      Width           =   660
      _ExtentX        =   1164
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
      MaxLength       =   2
      TextCase        =   1
      RawText         =   0
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpCidade 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   1170
      Width           =   5100
      _ExtentX        =   8996
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cep"
      Height          =   195
      Left            =   2055
      TabIndex        =   5
      Top             =   1590
      Width           =   285
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   570
      TabIndex        =   4
      Top             =   1590
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   1245
      Width           =   495
   End
End
Attribute VB_Name = "CadCidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vbGrava As Integer
Private vbBook  As Variant
Private vbIndex As String
Private vbRetg  As Boolean
Private tbCidades As Recordset

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdLocalizar_Click()
  RetCidade(0) = ""
  RetCidade(1) = ""
  lCidades.Show 1
  If RetCidade(0) <> "" Then
    With tbCidades
      .Index = "nomeid"
      .Seek "=", RetCidade(0), RetCidade(1)
      If Not .NoMatch Then
        LerDados
      End If
    End With
  End If
End Sub

Private Sub cpCep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    With tbCidades
      .AddNew
      !Nome = cpCidade.Text
      !estado = cpEstado.Text
      !cep = cpCep.Text
      .Update
    End With
    Unload Me
  End If
End Sub

Private Sub cpCidade_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEstado.SetFocus
  End If
End Sub

Private Sub cpEstado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCep.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    Select Case KeyCode
      Case 33
        With tbCidades
          If Not .BOF Then
            .MovePrevious
            If .BOF Then
              .MoveNext
              If .EOF Then
                Limpar
              Else
                LerDados
              End If
            Else
              LerDados
            End If
          End If
        End With
      Case 34
        With tbCidades
          If Not .EOF Then
            .MoveNext
            If .EOF Then
              .MovePrevious
              If .BOF Then
                Limpar
              Else
                LerDados
              End If
            Else
              LerDados
            End If
          End If
        End With
    End Select
  ElseIf Shift = 2 Then
    Select Case KeyCode
      Case 33
        With tbCidades
          If .RecordCount > 0 Then
            .MoveFirst
            LerDados
          End If
        End With
      Case 34
        With tbCidades
          If .RecordCount > 0 Then
            .MoveLast
            LerDados
          End If
        End With
    End Select
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCidades = db.OpenRecordset("cidades", dbOpenTable)
  Refresh
  KeyPreview = True
  With tbCidades
    .Index = "nome"
    If .RecordCount > 0 Then
      .MoveFirst
      LerDados
    End If
  End With
  Travar (True)
  Botoes (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RetCidade(0) = cpCidade.Text
  RetCidade(1) = cpEstado.Text
  RetCidade(2) = cpCep.Text
End Sub

Private Sub Botoes(Tipo As Boolean)
  cmdDesfazer.Enabled = Not Tipo
  cmdSalvar.Enabled = Not Tipo
  cmdAlterar.Enabled = Tipo
  cmdExcluir.Enabled = Tipo
  cmdFechar.Enabled = Tipo
  cmdIncluir.Enabled = Tipo
  cmdLocalizar.Enabled = Tipo
End Sub

Private Sub cmdIncluir_Click()
  Limpar
  Travar (False)
  Botoes (False)
  vbGrava = 1
  cpCidade.SetFocus
End Sub

Private Sub cmdExcluir_Click()
  Resp = MsgBox("Confirma a exclusão de '" + cpCidade.Text + "' do cadastro?", vbQuestion + vbYesNo + vbDefaultButton2, Titulo)
  If Resp = vbYes Then
    With tbCidades
      .Delete
      DoEvents
      .MovePrevious
      If .BOF Then
        .MoveNext
        If .EOF Then
          Limpar
        Else
          LerDados
        End If
      Else
        LerDados
      End If
    End With
  End If
End Sub

Private Sub cmdAlterar_Click()
  Travar (False)
  Botoes (False)
  With tbCidades
    If cpCidade.Text <> "" Then
      vbBook = .Bookmark
    End If
  End With
  vbGrava = 2
  cpCidade.SetFocus
End Sub

Private Sub cmdDesfazer_Click()
  Resp = MsgBox("Desfazer as alterações?", vbQuestion + vbYesNo + vbDefaultButton2, Titulo)
  If Resp = vbYes Then
    With tbCidades
      If vbGrava = 1 Then
        If .RecordCount > 0 Then
          .MoveFirst
          LerDados
        Else
          Limpar
        End If
      ElseIf vbGrava = 2 Then
        If .EOF And .BOF Then
          Limpar
        Else
          LerDados
        End If
      End If
    End With
    Botoes (True)
    Travar (True)
  End If
End Sub

Private Sub cmdSalvar_Click()
  If vbGrava = 1 Then
    If GravaDados(1) Then
      Travar (True)
      Botoes (True)
    End If
  ElseIf vbGrava = 2 Then
    If GravaDados(2, vbBook) Then
      Travar (True)
      Botoes (True)
    End If
  End If
End Sub

Private Sub Travar(ByVal Tipo As Boolean)
  cpCidade.Locked = Tipo
  cpEstado.Locked = Tipo
  cpCep.Locked = Tipo
End Sub

Private Sub Limpar()
  cpCidade.Text = ""
  cpEstado.Text = ""
  cpCep.Text = ""
End Sub

Private Function GravaDados(nTipo As Integer, Optional nBook As Variant) As Boolean
  With tbCidades
    If nTipo = 1 Then
      .AddNew
    ElseIf nTipo = 2 Then
      If IsNull(vbBook) Then
        .AddNew
      Else
        .Bookmark = nBook
      End If
      .Edit
    End If
    !estado = cpEstado.Text
    !cep = cpCep.Text
    !Nome = cpCidade.Text
    .Update
  End With
End Function

Private Sub LerDados()
  With tbCidades
    If .RecordCount > 0 Then
      cpEstado.Text = IIf(IsNull(!estado), "", !estado)
      cpCep.Text = IIf(IsNull(!cep), "", !cep)
      cpCidade.Text = IIf(IsNull(!Nome), "", !Nome)
    End If
  End With
End Sub
