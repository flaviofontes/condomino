VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form Cadhistorico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de históricos"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ControlBox      =   0   'False
   Icon            =   "CadHistorico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6345
      Picture         =   "CadHistorico.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   4995
      Picture         =   "CadHistorico.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4005
      Picture         =   "CadHistorico.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3015
      Picture         =   "CadHistorico.frx":092A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2025
      Picture         =   "CadHistorico.frx":0C34
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1035
      Picture         =   "CadHistorico.frx":0F3E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   45
      Picture         =   "CadHistorico.frx":1248
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   30
      Width           =   990
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   60
      ScaleHeight     =   0
      ScaleWidth      =   7215
      TabIndex        =   2
      Top             =   960
      Width           =   7275
   End
   Begin rdActiveText.ActiveText cpCidade 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   1170
      Width           =   6120
      _ExtentX        =   10795
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   1245
      Width           =   615
   End
End
Attribute VB_Name = "Cadhistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vbGrava As Integer
Private vbBook  As Variant
Dim rs As Recordset

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdLocalizar_Click()
  LocHistoricos.Show 1
  If Trim(RetNome) <> "" Then
    With rs
      .FindFirst "historico = '" & Trim(RetNome) & "'"
      If Not .NoMatch Then
        LerDados
      End If
    End With
  End If
End Sub

Private Sub cpCidade_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    Select Case KeyCode
      Case 33
        With rs
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
        With rs
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
        With rs
          If .RecordCount > 0 Then
            .MoveFirst
            LerDados
          End If
        End With
      Case 34
        With rs
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
  Refresh
  Set rs = db.OpenRecordset("select * from historicos order by historico;", dbOpenDynaset)
  
  KeyPreview = True
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      LerDados
    End If
  End With
  Travar (True)
  Botoes (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
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
  Resp = MsgBox("Confirma a exclusão de '" + cpCidade.Text + "' do cadastro?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir")
  If Resp = vbYes Then
    With rs
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
  With rs
    If cpCidade.Text <> "" Then
      vbBook = .Bookmark
    End If
  End With
  vbGrava = 2
  cpCidade.SetFocus
End Sub

Private Sub cmdDesfazer_Click()
  Resp = MsgBox("Desfazer as alterações?", vbQuestion + vbYesNo + vbDefaultButton2, "Desfazer")
  If Resp = vbYes Then
    With rs
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
End Sub

Private Sub Limpar()
  cpCidade.Text = ""
End Sub

Private Function GravaDados(nTipo As Integer, Optional nBook As Variant) As Boolean
On Error GoTo 10
  GravaDados = True
  With rs
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
    !Historico = cpCidade.Text
    .Update
  End With
5 Exit Function
10
  GravaDados = False
  MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume 5
End Function

Private Sub LerDados()
  With rs
    If .RecordCount > 0 Then
      cpCidade.Text = !Historico
    End If
  End With
End Sub
