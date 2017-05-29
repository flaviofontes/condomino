VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form CadTipoLeitura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de leituras"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "CadTipoLeitura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpDescricao 
      Height          =   315
      Left            =   1020
      TabIndex        =   11
      Top             =   1620
      Width           =   3735
      _ExtentX        =   6588
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
      MaxLength       =   20
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   1020
      TabIndex        =   10
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
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
      MaxLength       =   6
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   60
      Picture         =   "CadTipoLeitura.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1050
      Picture         =   "CadTipoLeitura.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2040
      Picture         =   "CadTipoLeitura.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3030
      Picture         =   "CadTipoLeitura.frx":092A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4020
      Picture         =   "CadTipoLeitura.frx":0C34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   5010
      Picture         =   "CadTipoLeitura.frx":0F3E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6360
      Picture         =   "CadTipoLeitura.frx":1248
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   75
      ScaleHeight     =   0
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   990
      Width           =   7260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1260
      Width           =   495
   End
End
Attribute VB_Name = "CadTipoLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As Recordset
Dim grTipo As Integer

Private Sub Botoes(Tipo As Boolean)
  cmdDesfazer.Enabled = Not Tipo
  cmdSalvar.Enabled = Not Tipo
  cmdAlterar.Enabled = Tipo
  cmdExcluir.Enabled = Tipo
  cmdFechar.Enabled = Tipo
  cmdIncluir.Enabled = Tipo
  cmdLocalizar.Enabled = Tipo
  cpDescricao.Locked = Tipo
End Sub

Private Sub cmdAlterar_Click()
  If cpDescricao.Text = "" Then
    MsgBox "Nenhum dado para alteração!", vbExclamation + vbOKOnly, "Aviso"
  Else
    grTipo = 2
    Botoes (False)
    cpDescricao.SetFocus
  End If
End Sub

Private Sub cmdDesfazer_Click()
  Resp = MsgBox("Cancelar a operação?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelar")
  If Resp = vbYes Then
    If rs.RecordCount > 0 Then
      LerDados
    Else
      Limpar
    End If
    Botoes (True)
  End If
End Sub

Private Sub cmdExcluir_Click()
  If cpDescricao.Text = "" Then
    MsgBox "Nenhum item para exclusão!", vbExclamation + vbOKOnly, "Aviso"
  Else
    If (PadeExcluir) Then
      Resp = MsgBox("Confirma a exclusão de '" & cpDescricao.Text & "'?", vbQuestion + vbYesNo, "Excluir")
      If Resp = vbYes Then
        rs.Delete
        Form_KeyDown 33, 0
      End If
    Else
      MsgBox "Este item possui movimentação e não pode ser excluido!", vbExclamation + vbOKOnly, "Aviso"
    End If
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdIncluir_Click()
  grTipo = 1
  Botoes (False)
  Limpar
  cpCodigo.Text = ProximoCodigo("tipo_leitura", "tipo_leitura")
  cpDescricao.SetFocus
End Sub

Private Sub cmdSalvar_Click()
  If cpDescricao.Text = "" Then
    MsgBox "Preencha o campo descrição.", vbInformation + vbOKOnly, "Aviso"
    cpDescricao.SetFocus
    Exit Sub
  End If
  If (Gravar()) Then
    grTipo = 0
    Botoes (True)
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
  KeyPreview = True
  Call Centraliza(Me)
  Refresh
  Botoes (True)
  Set rs = db.OpenRecordset("tipo_leitura", dbOpenTable)
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    cpCodigo.Text = rs!tipo_leitura
    cpDescricao.Text = rs!descricao
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub

Private Sub LerDados()
  If rs.RecordCount > 0 Then
    cpCodigo.Text = rs!tipo_leitura
    cpDescricao.Text = rs!descricao
  End If
End Sub

Private Sub Limpar()
  cpCodigo.Text = 0
  cpDescricao.Text = ""
End Sub

Private Function Gravar() As Boolean
  Dim ret As Boolean
  ret = False
  With rs
    If grTipo = 1 Then
      .AddNew
      rs!tipo_leitura = cpCodigo.Text
    ElseIf grTipo = 2 Then
      .Edit
    Else
      Exit Function
    End If
    rs!descricao = cpDescricao.Text
    rs.Update
    ret = True
  End With
  Gravar = ret
End Function

Private Function PadeExcluir() As Boolean
  Dim ret As Boolean
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("select * from contas_leitura where tipo_leitura = " & cpCodigo.Text & ";", dbOpenDynaset)
  If rs.RecordCount > 0 Then
    ret = False
  Else
    ret = True
  End If
  Set rs = Nothing
  PadeExcluir = ret
End Function
