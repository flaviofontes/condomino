VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form DespesaFixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de despesas fixas"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ControlBox      =   0   'False
   Icon            =   "Despesafixa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Despesafixa.frx":000C
      Height          =   2715
      Left            =   120
      OleObjectBlob   =   "Despesafixa.frx":0020
      TabIndex        =   7
      Top             =   3180
      Width           =   9135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CheckBox ChSindico 
      Caption         =   "Participação do síndico"
      Height          =   195
      Left            =   4980
      TabIndex        =   5
      Top             =   2400
      Width           =   1995
   End
   Begin VB.CheckBox chProprietario 
      Caption         =   "Despesa do proprietário"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   1995
   End
   Begin VB.ComboBox cpFracao 
      Height          =   315
      ItemData        =   "Despesafixa.frx":121F
      Left            =   1470
      List            =   "Despesafixa.frx":1221
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1470
      TabIndex        =   6
      Top             =   2685
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
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpData 
      Height          =   315
      Left            =   1470
      TabIndex        =   3
      Top             =   2325
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
   Begin rdActiveText.ActiveText cpHistorico 
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   1965
      Width           =   5490
      _ExtentX        =   9684
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
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   16
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
      Left            =   1470
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
      ScaleWidth      =   10200
      TabIndex        =   15
      Top             =   975
      Width           =   10260
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   9300
      Picture         =   "Despesafixa.frx":1223
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   4995
      Picture         =   "Despesafixa.frx":152D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4005
      Picture         =   "Despesafixa.frx":1837
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3015
      Picture         =   "Despesafixa.frx":1B41
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2025
      Picture         =   "Despesafixa.frx":1E4B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1035
      Picture         =   "Despesafixa.frx":2155
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   45
      Picture         =   "Despesafixa.frx":245F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   45
      Width           =   990
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Informe a qual fração esta despesa pertence:"
      Height          =   195
      Left            =   7080
      TabIndex        =   23
      Top             =   1140
      Width           =   3210
   End
   Begin MSForms.ListBox cpTipoFracao 
      Height          =   1515
      Left            =   7080
      TabIndex        =   22
      Top             =   1380
      Width           =   3255
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "5741;2672"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Left            =   1020
      TabIndex        =   21
      Top             =   1620
      Width           =   315
   End
   Begin VB.Label lbValor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   1005
      TabIndex        =   20
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data de cadastro"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   750
      TabIndex        =   18
      Top             =   2025
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condominio"
      Height          =   195
      Left            =   540
      TabIndex        =   17
      Top             =   1275
      Width           =   825
   End
End
Attribute VB_Name = "DespesaFixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private grTipo   As Integer
Private tbDespFixa  As Recordset
Private Ident As Long
Private tbCondominio As Recordset

Dim nInd As Double

Private Sub ChSindico_Click()
  Select Case TipoDivisao(cpCodigo.Text)
    Case 0, 2, 4
      ChSindico.Value = 0
      Principal.balao.ShowBalloonTip "Neste condomínio o síndico já participa das despesas.", beInformation, "Participação do síndico", 5000
  End Select
End Sub

Private Sub cmdAlterar_Click()
  If cpCodigo.Text = "" Then
    MsgBox "Localize um lançamento!", vbInformation + vbOKOnly, "Aviso"
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
    Botoes (True)
    Travar (True)
  End If
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Errado
  If cpHistorico.Text = "" Then
    MsgBox "Localize um lançamento!", vbInformation + vbOKOnly, "Aviso"
  Else
    Resp = MsgBox("Confirma a exclusão deste lançamento?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir")
    If Resp = vbYes Then
      With tbDespFixa
        db.Execute "delete from subdesp_fixa where id_desp_fixa = " & !id_despesa & ";"
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
  Ident = -1
  cpData.Text = Date
  cpCodigo.SetFocus
End Sub

Private Sub cmdLocalizar_Click()
On Error GoTo Errado
  Retorno.Historico = ""
  lDespFixa.Show 1
  If Not (Retorno.Historico = "") Then
    With tbDespFixa
      .Index = "primarykey"
      .Seek "=", Retorno.id_despesa
      If Not .NoMatch Then
        Ident = !id_despesa
        LerDados
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
  
  If CDbl(cpValor.Text) <= 0 And cpFracao.ListIndex < 3 Then
    MsgBox "Informe o valor desta despesa!", vbExclamation + vbOKOnly, "Aviso"
    cpValor.SetFocus
    Exit Sub
  End If
  
  If cpFracao.ListIndex < 0 Then
    MsgBox "Informe o tipo de distribuição da despesa!", vbExclamation + vbOKOnly, "Aviso"
    cpFracao.SetFocus
    Exit Sub
  End If
  
  If grTipo = 1 Then
    If Gravar(1) Then
      Botoes (True)
      Travar (True)
      Limpar
    End If
  ElseIf grTipo = 2 Then
    If Gravar(2) Then
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
            cpFracao_Click
            cpFracao.SetFocus
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
              cpFracao_Click
              cpFracao.SetFocus
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
  If grTipo < 2 Then
    ComboFracoes cpCodigo.Text
  End If
End Sub

Private Sub cpData_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  End If
End Sub

Private Sub cpFracao_Click()
  If cpFracao.ListIndex < 3 Then
    Me.Height = 3600
    If cpFracao.ListIndex = 2 Then
      lbValor.Caption = "Percentual"
      cpValor.Decimals = 5
    Else
      cpValor.Decimals = 2
      lbValor.Caption = "Valor"
    End If
  Else
    Me.Height = 6465
    If Val(cpCodigo.Text) > 0 Then
      PreencheLista cpCodigo.Text
      Data1.RecordSource = "select * from subdesp_fixa where id_desp_fixa = " & Ident & ";"
      Data1.Refresh
    End If
    If cpFracao.ListIndex = 3 Then
      Principal.balao.ShowBalloonTip "Preencha o campo 'Valor' e tecle F5 para atualizar lista de inquilinos!", beInformation, cpFracao.Text, 10000
    ElseIf cpFracao.ListIndex = 4 Then
      Principal.balao.ShowBalloonTip "Preencha os valores individuais na lista de inquilinos e tecle F5 para atualizar o campo 'Valor'!", beInformation, cpFracao.Text, 10000
    End If
  End If
End Sub

Private Sub cpFracao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpHistorico.SetFocus
  End If
End Sub

Private Sub cpHistorico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpData.SetFocus
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSalvar.SetFocus
  End If
End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  If ColIndex = 3 Then
    If Val(DBGrid1.Columns(3).Text) > 0 Then
      DBGrid1.Columns(2).Text = "S"
    End If
  End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If DBGrid1.Col = 2 Then
    If KeyAscii <> 13 Then
      If KeyAscii = Asc("s") Or KeyAscii = Asc("S") Or _
          KeyAscii = Asc("n") Or KeyAscii = Asc("N") Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Else
        KeyAscii = 0
      End If
    End If
  End If
  If DBGrid1.Col = 3 Then
    If KeyAscii <> Asc(",") Then
      KeyAscii = vNumero(KeyAscii)
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 116 Then
    If cpFracao.ListIndex = 3 And Trim(cpValor.Text) <> "" Then
      With Data1.Recordset
        If .RecordCount > 0 Then
          nInd = Round(CDbl(cpValor.Text) / Data1.Recordset.RecordCount, 2)
          .MoveFirst
          Do While Not .EOF
            .Edit
            !participa = "S"
            !valor = nInd
            .Update
            .MoveNext
          Loop
        End If
      End With
    ElseIf cpFracao.ListIndex = 4 Then
      With Data1.Recordset
        If .RecordCount > 0 Then
          nInd = 0
          .MoveFirst
          Do While Not .EOF
            If !participa = "S" Then
              nInd = nInd + !valor
            End If
            .MoveNext
          Loop
          cpValor.Text = nInd
        End If
      End With
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  KeyPreview = True
  Travar (True)
  Botoes (True)
  Data1.DatabaseName = Parametros.dados
  cpFracao.AddItem "Distribuição pela fração/Cota"
  cpFracao.AddItem "Distribuição pela Qtde Unidades"
  cpFracao.AddItem "Porcentagem sobre valor das despesas"
  cpFracao.AddItem "Valor fixo dividido por participante"
  cpFracao.AddItem "Valor fixo individual"
  cpFracao.ListIndex = 0
  db.Execute "delete from subdesp_fixa where id_desp_fixa = -1;"
  Set tbDespFixa = db.OpenRecordset("despesafixa", dbOpenTable)
  Refresh
End Sub

Private Sub LerDados()
  Dim i As Integer, j As Integer
  Dim s() As String
  With tbDespFixa
    If .RecordCount > 0 Then
      If !Fracao > 2 Then
        Me.Height = 6465
        Data1.RecordSource = "select * from subdesp_fixa where id_desp_fixa = " & !id_despesa & ";"
        Data1.Refresh
      Else
        Me.Height = 3600
      End If
      cpCodigo.Text = IIf(IsNull(!Condominio), "", !Condominio)
      cpNome.Text = IIf(IsNull(!Nome), "", !Nome)
      cpHistorico.Text = IIf(IsNull(!Historico), "", !Historico)
      cpData.Text = IIf(IsNull(!Data_cadastro), "", !Data_cadastro)
      chProprietario.Value = !Proprietario
      cpFracao.ListIndex = !Fracao
      If !Fracao = 2 Then
        cpValor.Decimals = 5
        lbValor.Caption = "Percentual"
      Else
        cpValor.Decimals = 2
        lbValor.Caption = "Valor"
      End If
      cpValor.Text = IIf(IsNull(!valor), "", !valor)
      ChSindico.Value = !Sindico
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
  cpCodigo.Text = ""
  cpNome.Text = ""
  cpHistorico.Text = ""
  cpData.Text = ""
  cpValor.Text = ""
  cpFracao.ListIndex = 0
  chProprietario.Value = 0
  ChSindico.Value = 0
  db.Execute "delete from subdesp_fixa where id_desp_fixa = -1;"
  Data1.RecordSource = "select * from subdesp_fixa where id_desp_fixa = -1;"
  Data1.Refresh
  DBGrid1.ReBind
  cpTipoFracao.Clear
End Sub

Private Sub Travar(ByVal bTipo As Boolean)
  cpCodigo.Locked = bTipo
  cpNome.Locked = bTipo
  cpHistorico.Locked = bTipo
  cpData.Locked = bTipo
  cpValor.Locked = bTipo
  cpFracao.Locked = bTipo
  chProprietario.Enabled = Not bTipo
  ChSindico.Enabled = Not bTipo
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

Private Function Gravar(ByVal iTipo As Integer) As Boolean
On Error GoTo Errado
  Dim i As Long
  Dim s As String
  Gravar = False
  With tbDespFixa
    If iTipo = 1 Then
      .AddNew
    Else
      .Edit
    End If
    !Condominio = cpCodigo.Text
    !Nome = cpNome.Text
    !Historico = cpHistorico.Text
    !Data_cadastro = cpData.Text
    !valor = cpValor.Text
    !Fracao = cpFracao.ListIndex
    !Proprietario = chProprietario.Value
    !Sindico = ChSindico.Value
    If cpFracao.ListIndex < 3 Then
      !por_associado = 0
    Else
      !por_associado = 1
    End If
    s = ""
    For i = 0 To cpTipoFracao.ListCount - 1
      If cpTipoFracao.Selected(i) = True Then
        s = s & cpTipoFracao.List(i) & "|"
      End If
    Next i
    !desc_fracao = s
    .Update
    .Bookmark = .LastModified
    i = !id_despesa
    db.Execute "update subdesp_fixa set id_desp_fixa = " & i & " where id_desp_fixa = -1;"
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

Private Sub Form_Unload(Cancel As Integer)
  Set tbDespFixa = Nothing
  Set tbCondominio = Nothing
End Sub


Private Sub PreencheLista(nCod As Long)
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("select * from associados where condominio = " & nCod & " order by tipo, apartamento;", dbOpenDynaset)
  db.Execute "delete from subdesp_fixa where id_desp_fixa = -1;"
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        Data1.Recordset.FindFirst "id_associado = " & !Codigo
        If Data1.Recordset.NoMatch Then
          db.Execute "insert into subdesp_fixa (id_associado, nome_associado, participa, valor, id_desp_fixa) " _
            & "values (" & !Codigo & ", '" & RetornaBloco(!id_bloco) & " " & NomeCompleto(!Codigo) & "', 'N', 0, " & Ident & ");"
        End If
        .MoveNext
      Loop
    End If
  End With
  Set rs = Nothing
End Sub

Private Sub ComboFracoes(ByVal nCod As Long, Optional ByVal marcar As Boolean = True)
  Dim rs As Recordset
  Dim pos As Integer
  Set rs = db.OpenRecordset("SELECT distinct FRACAO.DESCRICAO FROM ASSOCIADOS LEFT JOIN FRACAO ON " _
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


