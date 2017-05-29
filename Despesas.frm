VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Despesas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de despesas"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ControlBox      =   0   'False
   Icon            =   "Despesas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cpHistorico 
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   2280
      Width           =   5490
   End
   Begin VB.CheckBox ChPercentual 
      Caption         =   "Não contar percentuais (reserva, manutenção, outros...)"
      Height          =   195
      Left            =   2580
      TabIndex        =   10
      Top             =   3120
      Width           =   4275
   End
   Begin VB.CheckBox chDesconto 
      Caption         =   "É um desconto"
      Height          =   195
      Left            =   5340
      TabIndex        =   12
      Top             =   3600
      Width           =   1395
   End
   Begin VB.CheckBox ChSindico 
      Caption         =   "Participação do síndico"
      Height          =   195
      Left            =   4800
      TabIndex        =   8
      Top             =   2760
      Width           =   1995
   End
   Begin VB.CheckBox chProprietario 
      Caption         =   "Despesa do proprietário"
      Height          =   195
      Left            =   2580
      TabIndex        =   7
      Top             =   2760
      Width           =   1995
   End
   Begin VB.ComboBox cpFracao 
      Height          =   315
      ItemData        =   "Despesas.frx":000C
      Left            =   2580
      List            =   "Despesas.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
   End
   Begin rdActiveText.ActiveText cpSubTotal 
      Height          =   315
      Left            =   5340
      TabIndex        =   27
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   15
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.ComboBox cpTipoDesp 
      Height          =   315
      Left            =   1290
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1545
      Width           =   5490
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   1920
      Width           =   1170
      _ExtentX        =   2064
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
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1290
      TabIndex        =   9
      Top             =   3045
      Width           =   1170
      _ExtentX        =   2064
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
      MaxLength       =   12
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpData 
      Height          =   315
      Left            =   1290
      TabIndex        =   6
      Top             =   2685
      Width           =   1170
      _ExtentX        =   2064
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
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2190
      TabIndex        =   1
      Top             =   1185
      Width           =   4590
      _ExtentX        =   8096
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
      MaxLength       =   50
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   1185
      Width           =   885
      _ExtentX        =   1561
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
      TextMask        =   3
      RawText         =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   60
      ScaleHeight     =   0
      ScaleWidth      =   7200
      TabIndex        =   20
      Top             =   975
      Width           =   7260
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6345
      Picture         =   "Despesas.frx":005A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   4995
      Picture         =   "Despesas.frx":0364
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4005
      Picture         =   "Despesas.frx":066E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3015
      Picture         =   "Despesas.frx":0978
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2025
      Picture         =   "Despesas.frx":0C82
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1035
      Picture         =   "Despesas.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   45
      Picture         =   "Despesas.frx":1296
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   45
      Width           =   990
   End
   Begin MSForms.ListBox cpTipoFracao 
      Height          =   1335
      Left            =   1260
      TabIndex        =   11
      Top             =   3780
      Width           =   3255
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "5741;1906"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Informe a qual fração esta despesa pertence:"
      Height          =   195
      Left            =   1260
      TabIndex        =   29
      Top             =   3540
      Width           =   3210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total do mês para este tipo"
      Height          =   195
      Left            =   4860
      TabIndex        =   28
      Top             =   4560
      Width           =   1920
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tipo despesa"
      Height          =   195
      Left            =   255
      TabIndex        =   26
      Top             =   1620
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Do mês"
      Height          =   195
      Left            =   645
      TabIndex        =   25
      Top             =   1980
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   825
      TabIndex        =   24
      Top             =   3120
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   840
      TabIndex        =   23
      Top             =   2760
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   570
      TabIndex        =   22
      Top             =   2385
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condominio"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   1275
      Width           =   825
   End
End
Attribute VB_Name = "Despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private grTipo   As Integer
Private OldBook  As Variant
Private OldIndex As String
Private SelTipo  As Recordset
Private tbCondominio As Recordset
Private tbDespesas As Recordset
Dim pre As Preenchimento

Private Sub ChSindico_Click()
  Select Case TipoDivisao(cpCodigo.Text)
    Case 0, 2, 4
      ChSindico.Value = 0
      Principal.balao.ShowBalloonTip "Neste condomínio o síndico já participa das despesas.", beInformation, "Participação do síndico", 5000
  End Select
End Sub

Private Sub cmdAlterar_Click()
  If cpCodigo.Text = "" Then
    MsgBox "Selecione um lançamento!", vbInformation + vbOKOnly, "Aviso"
  Else
    Botoes (False)
    Travar (False)
    grTipo = 2
    cpCodigo.SetFocus
  End If
End Sub

Private Sub cmdDesfazer_Click()
  If grTipo = 1 Then
    Limpar
    Botoes (True)
    Travar (True)
  ElseIf grTipo = 2 Then
    With tbDespesas
      LerDados
    End With
    Botoes (True)
    Travar (True)
  End If
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Errado
  If cpHistorico.Text = "" Then
    MsgBox "Selecione um lançamento!", vbInformation + vbOKOnly, "Aviso"
  Else
    Resp = MsgBox("Confirma a exclusão deste lançamento?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir")
    If Resp = vbYes Then
      With tbDespesas
        .Delete
        Limpar
      End With
    End If
  End If

Fim:
  Exit Sub

Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdIncluir_Click()
  Limpar
  Travar (False)
  Botoes (False)
  grTipo = 1
  cpCodigo.SetFocus
End Sub

Private Sub cmdLocalizar_Click()
On Error GoTo Errado
  Retorno.Historico = ""
  lDespesa.Show 1
  If Retorno.Condominio > 0 Then
    With tbDespesas
      .Index = "primarykey"
      .Seek "=", Retorno.id_despesa
      If Not .NoMatch Then
        LerDados
        cpMes_LostFocus
      End If
    End With
  End If

Fim:
  Exit Sub

Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim

End Sub

Private Sub cmdSalvar_Click()
  If cpCodigo.Text = "" Then
    MsgBox "Escolha um condomínio!", vbExclamation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  If cpHistorico.Text = "" Then
    MsgBox "É necessário um histórico!", vbExclamation + vbOKOnly, "Aviso"
    cpHistorico.SetFocus
    Exit Sub
  End If
  If cpData.Text = "" Then
    MsgBox "Informe a data de pagamento da despesa!", vbExclamation + vbOKOnly, "Aviso"
    cpData.SetFocus
    Exit Sub
  End If
  If CDbl(cpValor.Text) <= 0 Then
    MsgBox "Informe o valor desta despesa!", vbExclamation + vbOKOnly, "Aviso"
    cpValor.SetFocus
    Exit Sub
  End If
  If cpMes.Text = "" Then
    MsgBox "Informe em que mês deve ser cobrada do associado!", vbExclamation + vbOKOnly, "Aviso"
    cpMes.SetFocus
    Exit Sub
  End If
  If cpFracao.ListIndex < 0 Then
    MsgBox "Informe o tipo de distribuição da despesa!", vbExclamation + vbOKOnly, "Aviso"
    cpFracao.SetFocus
    Exit Sub
  End If
  If cpTipoDesp.ListIndex < 0 Then
    MsgBox "Informe o tipo da despesa!", vbExclamation + vbOKOnly, "Aviso"
    cpTipoDesp.SetFocus
    Exit Sub
  End If
  
  Dim nTem As Integer, i As Integer
  nTem = 0
  For i = 0 To cpTipoFracao.ListCount - 1
    If cpTipoFracao.Selected(i) = True Then
      nTem = nTem + 1
    End If
  Next i
  
  If nTem = 0 Then
    MsgBox "Informe pelo menos uma fração a qual esta despesa pertence!", vbExclamation + vbOKOnly, "Aviso"
    cpTipoFracao.SetFocus
    Exit Sub
  End If
    
  
  If grTipo = 1 Then
    If Gravar(1) Then
      Botoes (True)
      Travar (True)
      Limpar
    End If
  ElseIf grTipo = 2 Then
    If Gravar(2, OldBook) Then
      Botoes (True)
      Travar (True)
      Limpar
    End If
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Not cpCodigo.Locked Then
      If Val(cpCodigo.Text) > 0 Then
        With tbCondominio
          .Index = "codigoid"
          .Seek "=", Val(cpCodigo.Text)
          If Not .NoMatch Then
            cpNome.Text = IIf(IsNull(!Nome), "", !Nome)
            ComboFracoes !Codigo
            cpTipoDesp.SetFocus
          Else
            MsgBox "Código não encontrado!", vbInformation + vbOKOnly, "Aviso"
            cpCodigo.SetFocus
          End If
        End With
      Else
        RetCodigo = 0
        l_Condominio.Show 1
        If RetCodigo > 0 Then
          With tbCondominio
            .Index = "codigoid"
            .Seek "=", RetCodigo
            If Not .NoMatch Then
              cpCodigo.Text = Str(RetCodigo)
              cpNome.Text = IIf(IsNull(!Nome), "", !Nome)
              ComboFracoes !Codigo
              cpTipoDesp.SetFocus
            End If
          End With
        End If
      End If
    End If
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpCodigo_LostFocus()
  ComboFracoes IIf(Trim(cpCodigo.Text) <> "", cpCodigo.Text, 0)
End Sub

Private Sub cpData_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  End If
End Sub

Private Sub cpHistorico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If InStr(cpHistorico.Text, "?") > 0 Then
      cpHistorico.SelStart = InStr(cpHistorico.Text, "?") - 1
      cpHistorico.SelLength = 1
    Else
      cpData.SetFocus
    End If
  Else
    Call pre.InfCbo(cpHistorico, KeyAscii)
  End If
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpHistorico.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  Dim vbTipo As Long
  Dim nSoma  As Double
  If IsDate("25/" & cpMes.Text) Then
    cpMes.Text = FormataMesAno(cpMes.Text)
    If cpTipoDesp.ListIndex >= 0 Then
      vbTipo = cpTipoDesp.ItemData(cpTipoDesp.ListIndex)
      Set SelTipo = db.OpenRecordset("Select * From Despesas Where Condominio = " & cpCodigo.Text & " And Tipo = " & vbTipo & " And Mes = '" & cpMes.Text & "'", dbOpenDynaset)
      nSoma = 0
      With SelTipo
        If .RecordCount > 0 Then
          .MoveLast
          .MoveFirst
          While Not .EOF
            nSoma = nSoma + !valor
            .MoveNext
          Wend
        End If
      End With
      Set SelTipo = Nothing
    End If
  End If
  cpSubTotal.Text = nSoma
End Sub

Private Sub cpTipoDesp_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMes.SetFocus
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSalvar.SetFocus
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
  KeyPreview = True
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Set tbDespesas = db.OpenRecordset("Despesas", dbOpenTable)
  Set pre = New Preenchimento
  Travar (True)
  Botoes (True)
  Call PreencheCombo(cpTipoDesp, "tpdesp", "codigo", "descricao")
  Call PreencheCombo(cpHistorico, "historicos", "", "historico")
  cpFracao.ListIndex = 0
  Refresh
End Sub

Private Sub LerDados()
  Dim i As Integer, j As Integer
  Dim s() As String
  With tbDespesas
    If .RecordCount > 0 Then
      cpCodigo.Text = IIf(IsNull(!Condominio), "", !Condominio)
      cpNome.Text = IIf(IsNull(!Nome), "", !Nome)
      cpHistorico.Text = IIf(IsNull(!Historico), "", !Historico)
      cpData.Text = IIf(IsNull(!Data), "", !Data)
      chDesconto.Value = !desconto
      If !desconto = 1 Then
        cpValor.Text = !valor * -1
      Else
        cpValor.Text = !valor
      End If
      cpMes.Text = IIf(IsNull(!mes), "", !mes)
      For i = 0 To cpTipoDesp.ListCount - 1
        If cpTipoDesp.ItemData(i) = !Tipo Then
          cpTipoDesp.ListIndex = i
          Exit For
        End If
      Next i
      chProprietario.Value = !Proprietario
      cpFracao.ListIndex = !Fracao
      ChSindico.Value = !Sindico
      ChPercentual.Value = !Percentual
      ComboFracoes !Condominio, False
      s = Split(!desc_fracao, "|")
      For j = 0 To UBound(s)
        If Trim(s(j)) <> "" Then
          For i = 0 To cpTipoFracao.ListCount - 1
            If cpTipoFracao.List(i) = Trim(s(j)) Then
              cpTipoFracao.Selected(i) = True
              Exit For
            End If
          Next i
        End If
      Next j
    End If
  End With
End Sub

Private Sub Limpar()
  cpHistorico.Text = ""
  cpData.Text = ""
  cpValor.Text = ""
  cpMes.Text = ""
  cpTipoDesp.ListIndex = -1
  cpSubTotal.Text = "0"
  cpFracao.ListIndex = 0
  chProprietario.Value = 0
  ChSindico.Value = 0
  chDesconto.Value = 0
  ChPercentual.Value = 0
  cpTipoFracao.Clear
End Sub

Private Sub Travar(ByVal bTipo As Boolean)
  cpCodigo.Locked = bTipo
  cpNome.Locked = bTipo
  cpHistorico.Locked = bTipo
  cpData.Locked = bTipo
  cpValor.Locked = bTipo
  cpMes.Locked = bTipo
  cpTipoDesp.Locked = bTipo
  cpFracao.Locked = bTipo
  chProprietario.Enabled = Not bTipo
  ChSindico.Enabled = Not bTipo
  ChPercentual.Enabled = Not bTipo
  chDesconto.Enabled = Not bTipo
  cpTipoFracao.Locked = bTipo
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

Private Function Gravar(ByVal iTipo As Integer, Optional mBook As Variant) As Boolean
On Error GoTo Errado
  Dim i As Integer
  Dim s As String
  Gravar = False
  With tbDespesas
    If iTipo = 1 Then
      .AddNew
    Else
      .Edit
    End If
    !Condominio = cpCodigo.Text
    !Nome = cpNome.Text
    !Historico = cpHistorico.Text
    !Data = cpData.Text
    !desconto = chDesconto.Value
    !Percentual = ChPercentual.Value
    If chDesconto.Value = 1 Then
      !valor = cpValor.Text * -1
    Else
      !valor = cpValor.Text
    End If
    !Tipo = cpTipoDesp.ItemData(cpTipoDesp.ListIndex)
    !mes = cpMes.Text
    !Fracao = cpFracao.ListIndex
    !Proprietario = chProprietario.Value
    !Sindico = ChSindico.Value
    s = ""
    For i = 0 To cpTipoFracao.ListCount - 1
      If cpTipoFracao.Selected(i) = True Then
        s = s & cpTipoFracao.List(i) & "|"
      End If
    Next i
    !desc_fracao = s
    .Update
  End With
  Gravar = True

Fim:
  Exit Function

Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim

End Function

Private Sub TravarBotoes()
  cmdDesfazer.Enabled = False
  cmdSalvar.Enabled = False
  cmdAlterar.Enabled = False
  cmdExcluir.Enabled = False
  cmdFechar.Enabled = True
  cmdIncluir.Enabled = False
  cmdLocalizar.Enabled = False
End Sub

Private Sub ComboFracoes(ByVal nCod As Long, Optional ByVal marcar As Boolean = True)
  Dim rs As Recordset
  Dim pos As Integer
  Set rs = db.OpenRecordset("SELECT distinct  FRACAO.DESCRICAO FROM ASSOCIADOS LEFT JOIN FRACAO ON " _
    & "(ASSOCIADOS.CODIGO = FRACAO.ID_ASSOCIADO) AND (ASSOCIADOS.CODIGO = FRACAO.ID_ASSOCIADO) " _
    & "WHERE (((ASSOCIADOS.CONDOMINIO)=" & nCod & "));", dbOpenDynaset)
  
  cpTipoFracao.Clear
  pos = -1
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If !Descricao & "" <> "" Then
          cpTipoFracao.AddItem !Descricao & ""
          If marcar Then
            If UCase(!Descricao) = "PRINCIPAL" Then
              cpTipoFracao.Selected(cpTipoFracao.ListCount - 1) = True
            End If
          End If
        End If
        .MoveNext
      Loop
    End If
  End With
  Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set pre = Nothing
  Set tbCondominio = Nothing
  Set tbDespesas = Nothing
End Sub
