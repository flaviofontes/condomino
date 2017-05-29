VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form PorFora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inclusão de despesa por inquilino"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   6660
      TabIndex        =   16
      Top             =   1620
      Width           =   1275
      _ExtentX        =   2249
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
      MaxLength       =   13
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   60
      Picture         =   "porfora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1050
      Picture         =   "porfora.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2040
      Picture         =   "porfora.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3030
      Picture         =   "porfora.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4020
      Picture         =   "porfora.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizaDespesa 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   5010
      Picture         =   "porfora.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   7020
      Picture         =   "porfora.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin rdActiveText.ActiveText cpVenc 
      Height          =   315
      Left            =   6660
      TabIndex        =   1
      Top             =   1080
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
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   1050
      Width           =   795
      _ExtentX        =   1402
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
   Begin VB.TextBox cpHistorico 
      Height          =   315
      Left            =   885
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1590
      Width           =   5205
   End
   Begin VB.TextBox cpNome 
      Height          =   315
      Left            =   2085
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1050
      Width           =   4005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   6240
      TabIndex        =   7
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1635
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mês"
      Height          =   195
      Left            =   6300
      TabIndex        =   5
      Top             =   1140
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inquilino"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1140
      Width           =   585
   End
End
Attribute VB_Name = "PorFora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As Recordset
Dim tbAssociados As Recordset
Dim tipoGR As Integer

Private Sub Botoes(Tipo As Boolean)
  cmdDesfazer.Enabled = Not Tipo
  cmdSalvar.Enabled = Not Tipo
  cmdAlterar.Enabled = Tipo
  cmdExcluir.Enabled = Tipo
  cmdFechar.Enabled = Tipo
  cmdIncluir.Enabled = Tipo
  cmdLocalizaDespesa.Enabled = Tipo
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdAlterar_Click()
  If rs Is Nothing Then
    MsgBox "Por favor localize o registro a alterar.", vbInformation + vbOKOnly, "Aviso"
  Else
    tipoGR = 2
    Botoes False
    Travar False
    cpCodigo.SetFocus
  End If
End Sub

Private Sub cmdDesfazer_Click()
  Botoes True
  Travar True
End Sub

Private Sub cmdExcluir_Click()
  If rs Is Nothing Then
    MsgBox "Por favor localize o registro a excluir.", vbInformation + vbOKOnly, "Aviso"
  Else
    Resp = MsgBox("Tem certeza que deseja excluir este débito?", vbQuestion + vbYesNo, "Ecluir")
    If Resp = vbYes Then
      rs.Delete
      DoEvents
      Limpar
    End If
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdIncluir_Click()
  Travar False
  Botoes False
  tipoGR = 1
  Limpar
  Set rs = db.OpenRecordset("porfora", dbOpenTable)
  cpCodigo.SetFocus
End Sub

Private Sub cmdLocalizaDespesa_Click()
  RetCodigo = 0
  Limpar
  lDespInquilino.Show 1
  If RetCodigo > 0 Then
    Set rs = db.OpenRecordset("select * from porfora where id = " & RetCodigo & ";", dbOpenDynaset)
    With rs
      If .RecordCount > 0 Then
        .MoveFirst
        cpVenc.Text = !mes
        cpHistorico.Text = !Historico
        cpValor.Text = !valor
        cpCodigo.Text = !Associado
        cpNome.Text = !Nome
      End If
    End With
  End If
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  lAssociado.Show 1
  If RetCodigo > 0 Then
    With tbAssociados
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpCodigo.Text = RetCodigo
        cpNome.Text = NomeCompleto(!Codigo)
        cpVenc.SetFocus
      End If
    End With
  End If
End Sub

Private Sub Limpar()
  Set rs = Nothing
  cpVenc.Text = ""
  cpHistorico.Text = ""
  cpValor.Text = "0"
  cpCodigo.Text = ""
  cpNome.Text = ""
End Sub

Private Sub cmdSalvar_Click()
  If Len(cpCodigo.Text) = 0 Then
    MsgBox "Informe o associado.", vbInformation + vbOKOnly, "Informações"
    Exit Sub
  End If
  
  If Not IsDate("25/" & cpVenc.Text) Then
    MsgBox "Informe o mês para inclusão no boleto.", vbInformation + vbOKOnly, "Informações"
    Exit Sub
  End If
  
  If Not IsNumeric(cpValor.Text) Then
    MsgBox "Informe o valor da despesa.", vbInformation + vbOKOnly, "Informações"
    Exit Sub
  End If
  
  Resp = MsgBox("Incluir/Alterar o valor de '" & Format$(cpValor.Text, "#,##0.00") & "' para '" & cpNome.Text & "'?", vbQuestion + vbOKCancel)
  If (Resp = vbOK) Then
    With rs
      If tipoGR = 1 Then
        .AddNew
      Else
        .Edit
      End If
      !mes = cpVenc.Text
      !Historico = cpHistorico.Text
      !valor = cpValor.Text
      !Associado = cpCodigo.Text
      !Nome = cpNome.Text
      .Update
    End With
    Travar True
    Botoes True
    Limpar
  End If

End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbAssociados
        .Index = "codigoid"
        .Seek "=", Val(cpCodigo.Text)
        If Not .NoMatch Then
          cpNome.Text = NomeCompleto(!Codigo)
          cpVenc.SetFocus
        Else
          MsgBox "Código não encontrado!", vbInformation + vbOKOnly, Caption
        End If
      End With
    Else
      RetCodigo = 0
      lAssociado.Show 1
      If RetCodigo > 0 Then
        With tbAssociados
          .Index = "codigoid"
          .Seek "=", RetCodigo
          If Not .NoMatch Then
            cpCodigo.Text = RetCodigo
            cpNome.Text = NomeCompleto(!Codigo)
            cpVenc.SetFocus
          End If
        End With
      End If
    End If
  End If
End Sub

Private Sub cpHistorico_GotFocus()
  cpHistorico.SelStart = 0
  cpHistorico.SelLength = Len(cpHistorico.Text)
End Sub

Private Sub cpHistorico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpNome_GotFocus()
  cpNome.SelStart = 0
  cpNome.SelLength = Len(cpNome.Text)
End Sub

Private Sub cpNome_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpVenc.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpValor_GotFocus()
  cpValor.SelStart = 0
  cpValor.SelLength = Len(cpValor.Text)
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdSalvar.Enabled Then
      cmdSalvar.SetFocus
    Else
      cpCodigo.SetFocus
    End If
  End If
End Sub

Private Sub cpVenc_GotFocus()
  cpVenc.SelStart = 0
  cpVenc.SelLength = Len(cpVenc.Text)
End Sub

Private Sub cpVenc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpHistorico.SetFocus
  End If
End Sub

Private Sub cpVenc_LostFocus()
  If cpVenc.Text <> "" Then
    cpVenc.Text = FormataMesAno(cpVenc.Text)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  Botoes True
  Travar True
  Refresh
  KeyPreview = True
End Sub

Private Sub Travar(ByVal Tipo As Boolean)
  cpVenc.Locked = Tipo
  cpHistorico.Locked = Tipo
  cpValor.Locked = Tipo
  cpCodigo.Locked = Tipo
  cmdLocalizar.Enabled = Not Tipo
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbAssociados = Nothing
  Set rs = Nothing
End Sub
