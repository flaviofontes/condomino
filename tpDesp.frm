VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form tpDesp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de despesa"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   ControlBox      =   0   'False
   Icon            =   "tpDesp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpDescricao 
      Height          =   315
      Left            =   1095
      TabIndex        =   9
      Top             =   1665
      Width           =   5625
      _ExtentX        =   9922
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
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   1095
      TabIndex        =   8
      Top             =   1290
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextMask        =   3
      RawText         =   3
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6360
      Picture         =   "tpDesp.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   5010
      Picture         =   "tpDesp.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4020
      Picture         =   "tpDesp.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3030
      Picture         =   "tpDesp.frx":092A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2040
      Picture         =   "tpDesp.frx":0C34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1050
      Picture         =   "tpDesp.frx":0F3E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   60
      Picture         =   "tpDesp.frx":1248
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   990
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   75
      ScaleHeight     =   0
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   990
      Width           =   7275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1785
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   465
      TabIndex        =   10
      Top             =   1380
      Width           =   495
   End
End
Attribute VB_Name = "tpDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private grTipo  As Integer
Private vbBook  As Variant
Private vbIndex As String
Private tbTipoDesp As Recordset
Private tbDespesas As Recordset

Private Sub cmdAlterar_Click()
  Travar (False)
  Botoes (False)
  grTipo = 2
  With tbTipoDesp
    If Not (cpCodigo.Text = "") Then
      vbBook = .Bookmark
    Else
      vbBook = Empty
    End If
  End With
  cpCodigo.SetFocus
End Sub

Private Sub cmdDesfazer_Click()
  With tbTipoDesp
    If .RecordCount > 0 Then
      If Not (.BOF And .EOF) Then
        LerDados
      Else
        Limpar
      End If
    Else
      Limpar
    End If
  End With
End Sub

Private Sub cmdExcluir_Click()
  With tbDespesas
    .Index = "Tipoid"
    .Seek "=", Val(cpCodigo.Text)
    If .NoMatch Then
      Resp = MsgBox("Confirma a exclusão deste tipo de despesa?", vbQuestion + vbYesNo + vbDefaultButton2, Caption)
      If Resp = vbYes Then
        With tbTipoDesp
          .Delete
          .MovePrevious
          If Not .BOF Then
            LerDados
          Else
            .MoveNext
            If Not .EOF Then
              LerDados
            Else
              Limpar
            End If
          End If
        End With
      End If
    Else
      MsgBox "Este tipo de despesa possui movimento, não pode ser excluido.", vbInformation + vbOKOnly, Caption
    End If
  End With
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdIncluir_Click()
  Botoes (False)
  Travar (False)
  Limpar
  grTipo = 1
  With tbTipoDesp
    .Index = "codigoid"
    If .RecordCount > 0 Then
      .MoveLast
      cpCodigo.Text = !Codigo + 1
    Else
      cpCodigo.Text = "1"
    End If
  End With
  cpCodigo.SetFocus
End Sub

Private Sub cmdLocalizar_Click()
  Dim retBook  As Variant
  RetCodigo = 0
  lTipoDesp.Show 1
  If RetCodigo > 0 Then
    With tbTipoDesp
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        retBook = .Bookmark
        .Index = "descricaoid"
        .Bookmark = retBook
        LerDados
      End If
    End With
  End If
End Sub

Private Sub cmdSalvar_Click()
  If grTipo = 1 Then
    If Gravar(1) Then
      Travar (True)
      Botoes (True)
    End If
  ElseIf grTipo = 2 Then
    If Gravar(2, vbBook) Then
      Travar (True)
      Botoes (True)
    End If
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDescricao.SetFocus
  End If
End Sub

Private Sub cpDescricao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSalvar.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    Select Case KeyCode
      Case 33
        With tbTipoDesp
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
        With tbTipoDesp
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
        With tbTipoDesp
          If .RecordCount > 0 Then
            .MoveFirst
            LerDados
          End If
        End With
      Case 34
        With tbTipoDesp
          If .RecordCount > 0 Then
            .MoveLast
            LerDados
          End If
        End With
    End Select
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    If cmdFechar.Enabled Then
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbDespesas = db.OpenRecordset("Despesas", dbOpenTable)
  Set tbTipoDesp = db.OpenRecordset("tpDesp", dbOpenTable)
  Refresh
  Botoes (True)
  Travar (True)
  KeyPreview = True
  With tbTipoDesp
    .Index = "descricaoid"
    If .RecordCount > 0 Then
      .MoveFirst
      LerDados
    End If
  End With
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

Private Sub LerDados()
  With tbTipoDesp
    If .RecordCount > 0 Then
      cpCodigo.Text = IIf(IsNull(!Codigo), "", !Codigo)
      cpDescricao.Text = IIf(IsNull(!Descricao), "", !Descricao)
    End If
  End With
End Sub

Private Function Gravar(ByVal gTipo As Integer, Optional vBook As Variant) As Boolean
On Error GoTo Errado
  With tbTipoDesp
    If gTipo = 1 Then
      .AddNew
      !Codigo = cpCodigo.Text
    Else
      .Edit
    End If
    !Descricao = cpDescricao.Text
    .Update
  End With
  Gravar = True
  
Fim:
  Exit Function

Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim

End Function

Private Sub Travar(ByVal bTipo As Boolean)
  cpDescricao.Locked = bTipo
End Sub

Private Sub Limpar()
  cpCodigo.Text = ""
  cpDescricao.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbTipoDesp = Nothing
  Set tbDespesas = Nothing
End Sub
