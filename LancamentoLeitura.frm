VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form LancamentoLeitura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento de leituras"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "LancamentoLeitura.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocCond 
      Caption         =   "..."
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin VB.ComboBox cpBlocos 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   780
      Width           =   3255
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   435
      Left            =   7260
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin rdActiveText.ActiveText cpReferencia 
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   420
      Width           =   1095
      _ExtentX        =   1931
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
   Begin rdActiveText.ActiveText cpSomaCusto 
      Height          =   315
      Left            =   5280
      TabIndex        =   24
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
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
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpSomaGasto 
      Height          =   315
      Left            =   3900
      TabIndex        =   23
      Top             =   6360
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
      MaxLength       =   10
      Text            =   "0,000"
      TextMask        =   4
      RawText         =   4
      Decimals        =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.CommandButton cmdLimpardados 
      Caption         =   "Limpar dados"
      Height          =   435
      Left            =   7260
      TabIndex        =   15
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdGravarDados 
      Caption         =   "Gravar dados"
      Height          =   435
      Left            =   7260
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LEITURA_AUXILIAR"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LancamentoLeitura.frx":000C
      Height          =   2955
      Left            =   60
      OleObjectBlob   =   "LancamentoLeitura.frx":0020
      TabIndex        =   12
      Top             =   3360
      Width           =   7035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da leitura"
      Height          =   1995
      Left            =   60
      TabIndex        =   18
      Top             =   1320
      Width           =   7035
      Begin VB.OptionButton OpFracao 
         Caption         =   "Fração/Cota"
         Height          =   195
         Left            =   5640
         TabIndex        =   29
         Top             =   660
         Width           =   1215
      End
      Begin rdActiveText.ActiveText cpDescricao 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   5115
         _ExtentX        =   9022
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
         MaxLength       =   30
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Divisão total/residentes"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   660
         Width           =   1995
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Divisão por total/m3"
         Height          =   195
         Left            =   1740
         TabIndex        =   7
         Top             =   660
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por valor do m3"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1455
      End
      Begin rdActiveText.ActiveText cpValor 
         Height          =   315
         Left            =   5160
         TabIndex        =   11
         Top             =   1260
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
         MaxLength       =   10
         Text            =   "0,00"
         TextMask        =   4
         RawText         =   4
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpConsumo 
         Height          =   315
         Left            =   3900
         TabIndex        =   10
         Top             =   1260
         Width           =   1215
         _ExtentX        =   2143
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
         Text            =   "0,000"
         TextMask        =   4
         RawText         =   4
         Decimals        =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpData 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1260
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbUltimaLeitura 
         Alignment       =   2  'Center
         Caption         =   "Última leitura"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label lbValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor total"
         Height          =   195
         Left            =   5880
         TabIndex        =   21
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label lbConsumo 
         AutoSize        =   -1  'True
         Caption         =   "Consumo (m3)"
         Height          =   195
         Left            =   4080
         TabIndex        =   20
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1020
         Width           =   345
      End
   End
   Begin VB.ComboBox cpTipoLeitura 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   2715
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2580
      TabIndex        =   28
      Top             =   60
      Width           =   4515
      _ExtentX        =   7964
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
      Left            =   1320
      TabIndex        =   0
      Top             =   60
      Width           =   735
      _ExtentX        =   1296
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Bloco"
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Ref. ao mês de"
      Height          =   195
      Left            =   4080
      TabIndex        =   25
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de leitura"
      Height          =   195
      Left            =   180
      TabIndex        =   17
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "LancamentoLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cGasto As Double
Dim cValor As Double
Dim DataUltima As Date
Dim Reinicio As Double
Dim sSql As String
Dim exIDleitura As Long
Dim id_cond As Long
Dim tbCondominio As Recordset

Private Sub cmdGravarDados_Click()
  
  If Not IsDate(cpData.Text) Then
    MsgBox "A data da leitura não foi informada ou não é válida!", vbCritical + vbOKOnly, "Aviso"
    cpData.SetFocus
    Exit Sub
  End If
  
  If Trim(cpDescricao.Text) = "" Then
    MsgBox "Informe uma descrição para esta leitura.", vbCritical + vbOKOnly, "Aviso"
    cpDescricao.SetFocus
    Exit Sub
  End If
  
  If cmdLimpardados.Caption <> "Excluir" Then
    If CDate(cpData.Text) <= DataUltima Then
      MsgBox "A data da leitura não pode ser anterior ou igual a data da última leitura!", vbCritical + vbOKOnly, "Aviso"
      cpData.SetFocus
      Exit Sub
    End If
  End If
  
  If Not IsDate("01/" & cpReferencia.Text) Then
    MsgBox "Informe o mês de referência ao qual esta despesa pertence!", vbCritical + vbOKOnly, "Aviso"
    cpReferencia.SetFocus
    Exit Sub
  End If
  
  Dim rsContas As Recordset
  Dim rsLeitura As Recordset
  Dim idLeitura As Long
  Dim idCondominio As Long
  
  Set rsContas = db.OpenRecordset("contas_leitura", dbOpenTable)
  Set rsLeitura = db.OpenRecordset("leitura_individual", dbOpenTable)
  
  If cmdLimpardados.Caption = "Excluir" Then
    db.Execute "delete from leitura_individual where id_leitura = " & exIDleitura & ";"
    db.Execute "delete from contas_leitura where id_leitura = " & exIDleitura & ";"
  End If
  
  idCondominio = cpCodigo.Text
  rsContas.AddNew
  rsContas!data_leitura = cpData.Text
  If Option1.Value Then
    rsContas!ant_leitura = 1
  ElseIf Option2.Value Then
    rsContas!ant_leitura = 2
  ElseIf Option3.Value Then
    rsContas!ant_leitura = 3
  ElseIf OpFracao.Value Then
    rsContas!ant_leitura = 4
  End If
  rsContas!atual_leitura = cpConsumo.Text
  rsContas!desc_leitura = cpDescricao.Text
  rsContas!total_leitura = cpValor.Text
  rsContas!valor_leitura = cpValor.Text
  rsContas!cond_leitura = idCondominio
  rsContas!mes_leitura = cpReferencia.Text
  rsContas!codigo_condominio = idCondominio
  rsContas!tipo_leitura = cpTipoLeitura.ItemData(cpTipoLeitura.ListIndex)
  rsContas!id_bloco = cpBlocos.ItemData(cpBlocos.ListIndex)
  idLeitura = rsContas!id_leitura
  rsContas.Update
  
  With Data1.Recordset
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        rsLeitura.AddNew
        rsLeitura!ant_individual = !ant_individual
        rsLeitura!atual_individual = !atual_individual
        rsLeitura!gasto_individual = !gasto_individual
        rsLeitura!valor_individual = !valor_individual
        rsLeitura!id_leitura = idLeitura
        rsLeitura!id_associado = !id_associado
        rsLeitura!id_bloco = cpBlocos.ItemData(cpBlocos.ListIndex)
        rsLeitura.Update
        .MoveNext
      Loop
    End If
  End With
  Limpar
  cpData.Locked = False
  MsgBox "Dados gravados com sucesso!", vbExclamation + vbOKOnly, "Aviso"
End Sub

Private Sub cmdLimpardados_Click()
  If cmdLimpardados.Caption = "Limpar dados" Then
    Resp = MsgBox("Limpar os dados?", vbQuestion + vbYesNo, "Limpar")
    If Resp = vbYes Then
      Limpar
    End If
  ElseIf cmdLimpardados.Caption = "Excluir" Then
    Resp = MsgBox("Todos os dados desta leitura serão escluidos. Continuar?", vbQuestion + vbYesNo, "Limpar")
    If Resp = vbYes Then
      db.Execute "delete from leitura_individual where id_leitura = " & exIDleitura & ";"
      db.Execute "delete from contas_leitura where id_leitura = " & exIDleitura & ";"
      Limpar
    End If
  End If
End Sub

Private Sub Limpar()
  dbLocal.Execute "delete from leitura_auxiliar;"
  Data1.Refresh
  DBGrid1.ReBind
  cpData.Text = ""
  cpValor.Text = 0
  cpConsumo.Text = 0
  cpReferencia.Text = ""
  cpSomaCusto.Text = 0
  cpSomaGasto.Text = 0
  cpDescricao.Text = ""
  cpCodigo.Text = ""
  cpNome.Text = ""
  cpTipoLeitura.ListIndex = -1
  exIDleitura = -1
End Sub

Private Sub cmdLocalizar_Click()
  Dim retCod As Long
  Dim rsContas As Recordset
  Dim rsLeitura As Recordset
  Dim i As Long
  Dim sg As Double
  Dim sv As Double
  
  
  LocLeitura.Show 1
  retCod = LocLeitura.Retorno
  
  If retCod > 0 Then
    'ler os dados
    Set rsContas = db.OpenRecordset("select * from contas_leitura where id_leitura = " & retCod & ";", dbOpenDynaset)
    Set rsLeitura = db.OpenRecordset("select * from LEITURA_INDIVIDUAL where id_leitura = " & retCod & ";", dbOpenDynaset)
    
    With rsContas
      If .RecordCount > 0 Then
        .MoveFirst
        cpData.Text = !data_leitura
        cpValor.Text = !valor_leitura
        cpReferencia.Text = !mes_leitura
        cpConsumo.Text = !atual_leitura
        cpDescricao.Text = !desc_leitura
        
        If rsContas!ant_leitura = 1 Then
          Option1.Value = True
          Option1_Click
        ElseIf rsContas!ant_leitura = 2 Then
          Option2.Value = True
          Option2_Click
        ElseIf rsContas!ant_leitura = 3 Then
          Option3.Value = True
          Option3_Click
        ElseIf rsContas!ant_leitura = 4 Then
          OpFracao.Value = True
          OpFracao_Click
        End If
        
        cpCodigo.Text = !codigo_condominio
        cpNome.Text = NomeCondominio(!codigo_condominio)
        For i = 0 To cpTipoLeitura.ListCount - 1
          If !tipo_leitura = cpTipoLeitura.ItemData(i) Then
            cpTipoLeitura.ListIndex = i
            Exit For
          End If
        Next i
        cpBlocos.ListIndex = MostraBloco(!id_bloco)
        ConfereSelecao !ant_leitura
      End If
    End With
    sg = 0
    sv = 0
    dbLocal.Execute "delete from leitura_auxiliar;"
    With rsLeitura
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          Data1.Recordset.AddNew
          Data1.Recordset!Fracao = RetornaFracao(!id_associado)
          Data1.Recordset!ant_individual = !ant_individual
          Data1.Recordset!atual_individual = !atual_individual
          Data1.Recordset!gasto_individual = !gasto_individual
          Data1.Recordset!valor_individual = !valor_individual
          Data1.Recordset!id_leitura = !id_leitura
          Data1.Recordset!id_associado = !id_associado
          Data1.Recordset!nome_associado = PegaNomeAssociado(!id_associado)
          Data1.Recordset!id_bloco = cpBlocos.ItemData(cpBlocos.ListIndex)
          Data1.Recordset.Update
          sg = sg + !gasto_individual
          sv = sv + !valor_individual
          .MoveNext
        Loop
      End If
    End With
    Data1.Refresh
    DBGrid1.ReBind
    cpSomaGasto.Text = sg
    cpSomaCusto.Text = sv
    exIDleitura = retCod
    lbUltimaLeitura.Caption = "Leitura carrega para Alteração/Exclusão"
    cmdLimpardados.Caption = "Excluir"
    cpData.Locked = True
  Else
    exIDleitura = -1
    cmdLimpardados.Caption = "Limpar dados"
  End If
  
End Sub

Private Sub PegaDados()
  If Val(cpCodigo.Text) > 0 And cpTipoLeitura.Text <> "" Then
    PegaUltimaLeitura
  End If
  If Val(cpCodigo.Text) > 0 Then
    cpBlocos.Clear
    sSql = "WHERE ID_BLOCO IN (SELECT DISTINCT ASSOCIADOS.ID_BLOCO FROM CONDOMINIO INNER JOIN ASSOCIADOS ON CONDOMINIO.CODIGO = ASSOCIADOS.CONDOMINIO WHERE (((CONDOMINIO.CODIGO)=" & cpCodigo.Text & ")))"
    Call PreencheCombo(cpBlocos, "blocos", "id_bloco", "nome_bloco", sSql)
    cpBlocos.ListIndex = BlocoUnico()
    cpTipoLeitura_LostFocus
  End If
End Sub

Private Sub cmdLocCond_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpNome.Text = !Nome
        cpCodigo.Text = !Codigo
        PegaDados
      End If
    End With
  End If
End Sub

Private Sub cpBlocos_LostFocus()
  If Val(cpCodigo.Text) > 0 And cpTipoLeitura.Text <> "" Then
    dbLocal.Execute "delete from leitura_auxiliar;"
    Data1.Refresh
    PegaUltimaLeitura
    'chama divisão
    cValor = 0
    MontaDivisao
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodigo.Text
        If Not .NoMatch Then
          cpNome.Text = !Nome
          cpTipoLeitura.SetFocus
        Else
          MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
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
            cpNome.Text = !Nome
            cpCodigo.Text = !Codigo
          End If
        End With
        cpTipoLeitura.SetFocus
      End If
    End If
    PegaDados
  End If
End Sub

Private Sub cpData_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    DBGrid1.SetFocus
  End If
End Sub

Private Sub cpDescricao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpData.SetFocus
  End If
End Sub

Private Sub cpReferencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDescricao.SetFocus
  End If
End Sub

Private Sub cpReferencia_LostFocus()
  cpReferencia.Text = FormataMesAno(cpReferencia.Text)
End Sub

Private Sub cpTipoLeitura_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpReferencia.SetFocus
  End If
End Sub

Private Sub cpTipoLeitura_LostFocus()
  If Val(cpCodigo.Text) > 0 And cpTipoLeitura.Text <> "" Then
    PegaUltimaLeitura
    'chama divisão
    cValor = 0
    MontaDivisao
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    DBGrid1.SetFocus
  End If
End Sub

Private Sub cpValor_LostFocus()
  If cpConsumo.Text = 0 And OpFracao.Value = False Then
    MsgBox "Quantidade de consumo deve ser positiva.", vbInformation + vbOKOnly, "Aviso"
  Else
    MontaDivisao
  End If
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo Errado
  If ColIndex = 1 Or ColIndex = 2 Then
    If CDbl(DBGrid1.Columns(2).Text) < CDbl(DBGrid1.Columns(1).Text) Then
      MsgBox "O valor para a leitura atual deve ser maior que a anterior.", vbCritical + vbOKOnly, "Aviso"
      Resp = MsgBox("É reinício de leitura?", vbQuestion + vbYesNo, "Leitura")
      If Resp = vbYes Then
        Reinicio = CDbl(PadRight("1", "0", Len(DBGrid1.Columns(1).Text) + 1))
        cGasto = Reinicio - CDbl(DBGrid1.Columns(1).Text) + CDbl(DBGrid1.Columns(2).Text)
        cpSomaGasto.Text = cpSomaGasto.Text - CDbl(DBGrid1.Columns(3).Text) + cGasto
        DBGrid1.Columns(3).Text = cGasto
        cpSomaCusto.Text = cpSomaCusto.Text - CDbl(DBGrid1.Columns(4).Text)
        DBGrid1.Columns(4).Text = Round(cGasto * cValor, 2)
        cpSomaCusto.Text = cpSomaCusto.Text + CDbl(DBGrid1.Columns(4).Text)
      Else
        Cancel = True
      End If
      DBGrid1.SetFocus
    Else
      cGasto = Round(CDbl(DBGrid1.Columns(2).Text) - CDbl(DBGrid1.Columns(1).Text), 3)
      cpSomaGasto.Text = cpSomaGasto.Text - CDbl(DBGrid1.Columns(3).Text) + cGasto
      DBGrid1.Columns(3).Text = cGasto
      cpSomaCusto.Text = cpSomaCusto.Text - CDbl(DBGrid1.Columns(4).Text)
      DBGrid1.Columns(4).Text = Round(cGasto * cValor, 2)
      cpSomaCusto.Text = cpSomaCusto.Text + CDbl(DBGrid1.Columns(4).Text)
    End If
    cpConsumo.Text = cpSomaGasto.Text
  ElseIf ColIndex = 3 Then
    cpConsumo.Text = cpConsumo.Text - OldValue
    cpConsumo.Text = cpConsumo.Text + CDbl(DBGrid1.Columns(3).Text)
    cpSomaGasto.Text = cpConsumo.Text
  End If
Sair:
  Exit Sub
Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Mensagem: " & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 46 Then KeyAscii = 44
  If KeyAscii <> 44 Then
    KeyAscii = vMoedaDbgrid(KeyAscii, DBGrid1.Columns(DBGrid1.Col).Text)
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 116 Then
    DoEvents
    MontaDivisao
    DoEvents
    DBGrid1.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Call Centraliza(Me)
  Call PreencheCombo(cpTipoLeitura, "tipo_leitura", "tipo_leitura", "descricao")
  Call PreencheCombo(cpBlocos, "blocos", "id_bloco", "nome_bloco")
  dbLocal.Execute "delete from leitura_auxiliar;"
  Data1.DatabaseName = sFormataCaminho(App.Path) & "supersoft.mdb"
  KeyPreview = True
  cpBlocos.ListIndex = BlocoUnico()
  Refresh
End Sub

Private Sub PegaUltimaLeitura()
  Dim rsLeitura As Recordset
  
  If cpBlocos.ListIndex >= 0 Then
    Set rsLeitura = db.OpenRecordset("select * from contas_leitura where codigo_condominio = " & _
        cpCodigo.Text & " and tipo_leitura = " & _
        cpTipoLeitura.ItemData(cpTipoLeitura.ListIndex) & " and id_bloco = " & _
        cpBlocos.ItemData(cpBlocos.ListIndex) & " order by data_leitura;", dbOpenDynaset)
    
    With rsLeitura
      If .RecordCount > 0 Then
        .MoveLast
        ConfereSelecao !ant_leitura
        lbUltimaLeitura.Caption = "Última leitura: " & Format$(.Fields("data_leitura").Value, "dd/MM/yyyy")
        DataUltima = .Fields("data_leitura").Value
      Else
        lbUltimaLeitura.Caption = "Última leitura: Sem leitura anterior"
        DataUltima = CDate("01/01/1900")
      End If
    End With
    Set rsLeitura = Nothing
  End If
End Sub

Private Function PegaUltimaLeituraIndividual(ByVal iCod As Long) As Long
  Dim rsLeitura As Recordset
  
  Set rsLeitura = db.OpenRecordset("SELECT LEITURA_INDIVIDUAL.ATUAL_INDIVIDUAL, CONTAS_LEITURA.DATA_LEITURA " _
      & "FROM CONTAS_LEITURA LEFT JOIN LEITURA_INDIVIDUAL ON CONTAS_LEITURA.ID_LEITURA = LEITURA_INDIVIDUAL.ID_LEITURA " _
      & "Where (((LEITURA_INDIVIDUAL.ID_ASSOCIADO) = " & iCod & ") And ((CONTAS_LEITURA.tipo_leitura) = " _
      & cpTipoLeitura.ItemData(cpTipoLeitura.ListIndex) & ")) ORDER BY CONTAS_LEITURA.DATA_LEITURA;", dbOpenDynaset)
  
  With rsLeitura
    If .RecordCount > 0 Then
      .MoveLast
      PegaUltimaLeituraIndividual = .Fields("ATUAL_INDIVIDUAL").Value
    Else
      PegaUltimaLeituraIndividual = 0
    End If
  End With
  Set rsLeitura = Nothing
End Function

Private Sub MontaDivisao()
  Dim rsAssociado As Recordset
  Dim nSoma As Double
  Dim nTotalFrac As Double
  
Dinovo:
  If cpBlocos.ListIndex >= 0 Then
    If Data1.Recordset.RecordCount = 0 Then
      Set rsAssociado = db.OpenRecordset("select tipo, apartamento, codigo from associados where condominio = " _
        & cpCodigo.Text & " and id_bloco = " & _
        cpBlocos.ItemData(cpBlocos.ListIndex) & " order by tipo, apartamento;", dbOpenDynaset)
      id_cond = cpCodigo.Text
      With rsAssociado
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            Data1.Recordset.AddNew
            Data1.Recordset!ant_individual = PegaUltimaLeituraIndividual(!Codigo)
            Data1.Recordset!atual_individual = 0
            Data1.Recordset!gasto_individual = 0
            Data1.Recordset!valor_individual = 0
            Data1.Recordset!id_leitura = 0
            Data1.Recordset!id_associado = !Codigo
            Data1.Recordset!nome_associado = !Tipo & " " & !Apartamento
            Data1.Recordset.Update
            .MoveNext
          Loop
        End If
      End With
      If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
      End If
      DBGrid1.ReBind
      Set rsAssociado = Nothing
    Else
      If id_cond > 0 And id_cond <> cpCodigo.Text Then
        dbLocal.Execute "delete from leitura_auxiliar;"
        Data1.Refresh
        DoEvents
        GoTo Dinovo
      End If
      If Option1.Value Then
        cValor = cpValor.Text
      ElseIf Option2.Value Then
        nTotalFrac = 0
        With Data1.Recordset
          If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
              nTotalFrac = nTotalFrac + !gasto_individual
              .MoveNext
            Loop
          End If
        End With
        cpConsumo.Text = nTotalFrac
        If cpConsumo.Text > 0 Then
          cValor = cpValor.Text / cpConsumo.Text
        Else
          cValor = 0
        End If
      ElseIf OpFracao.Value Then
        nTotalFrac = 0
        With Data1.Recordset
          If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
              .Edit
              !Fracao = RetornaFracao(!id_associado)
              nTotalFrac = nTotalFrac + !Fracao
              .Update
              .MoveNext
            Loop
            If nTotalFrac > 0 Then
              nSoma = cpValor.Text / nTotalFrac
            Else
              nSoma = 0
            End If
            nTotalFrac = 0
            .MoveFirst
            Do While Not .EOF
              .Edit
              !valor_individual = Round(!Fracao * nSoma, 2)
              nTotalFrac = nTotalFrac + !valor_individual
              .Update
              .MoveNext
            Loop
            cpSomaCusto.Text = Format$(nTotalFrac, "#,##0.00")
          End If
        End With
      Else
        If cpConsumo.Text > 0 Then
          cValor = cpValor.Text / cpConsumo.Text
        Else
          cValor = 0
        End If
      End If
      If OpFracao.Value = False Then
        With Data1.Recordset
          .MoveFirst
          nSoma = 0
          Do While Not .EOF
            .Edit
            !valor_individual = Round(!gasto_individual * cValor, 2)
            nSoma = nSoma + !valor_individual
            .Update
            .MoveNext
          Loop
          cpSomaCusto.Text = nSoma
          .MoveFirst
        End With
      End If
      DBGrid1.SetFocus
    End If
  Else
    dbLocal.Execute "delete from leitura_auxiliar;"
    Data1.Refresh
  End If
  Set rsAssociado = Nothing
  DoEvents
End Sub

Private Function PegaNomeAssociado(ByVal iCodigo As Long) As String
  Dim rs As Recordset
  Dim ret As String
  Set rs = db.OpenRecordset("select tipo, apartamento from associados where codigo = " & iCodigo & ";", dbOpenDynaset)
  ret = ""
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      ret = !Tipo & " " & !Apartamento
    End If
  End With
  Set rs = Nothing
  PegaNomeAssociado = ret
End Function

Private Sub ConfereSelecao(ByVal nTipo As Integer)
  If nTipo = 1 Then
    Option1.Value = True
    Option1_Click
  ElseIf nTipo = 2 Then
    Option2.Value = True
    Option2_Click
  ElseIf nTipo = 3 Then
    Option3.Value = True
    Option3_Click
  ElseIf nTipo = 4 Then
    OpFracao.Value = True
    OpFracao_Click
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub

Private Sub OpFracao_Click()
  DBGrid1.Columns(1).Visible = False
  DBGrid1.Columns(2).Visible = False
  DBGrid1.Columns(3).Locked = False
  DBGrid1.Columns(3).Caption = "Fração/Cota"
  DBGrid1.Columns(3).DataField = "FRACAO"
  DBGrid1.ReBind
  lbValor.Caption = "Valor total"
  lbConsumo.Caption = "Consumo (m3)"
End Sub

Private Sub Option1_Click()
  DBGrid1.Columns(1).Visible = True
  DBGrid1.Columns(2).Visible = True
  DBGrid1.Columns(3).Locked = True
  DBGrid1.Columns(3).Caption = "Gasto (m3)"
  DBGrid1.Columns(3).DataField = "GASTO_INDIVIDUAL"
  DBGrid1.ReBind
  lbValor.Caption = "Valor do m3"
  lbConsumo.Caption = "Consumo (m3)"
End Sub

Private Sub Option2_Click()
  DBGrid1.Columns(1).Visible = True
  DBGrid1.Columns(2).Visible = True
  DBGrid1.Columns(3).Locked = True
  DBGrid1.Columns(3).Caption = "Gasto (m3)"
  DBGrid1.Columns(3).DataField = "GASTO_INDIVIDUAL"
  DBGrid1.ReBind
  lbValor.Caption = "Valor total"
  lbConsumo.Caption = "Consumo (m3)"
End Sub

Private Sub Option3_Click()
  DBGrid1.Columns(1).Visible = False
  DBGrid1.Columns(2).Visible = False
  DBGrid1.Columns(3).Locked = False
  DBGrid1.Columns(3).Caption = "Residentes"
  DBGrid1.Columns(3).DataField = "GASTO_INDIVIDUAL"
  DBGrid1.ReBind
  lbValor.Caption = "Valor total"
  lbConsumo.Caption = "Residentes"
End Sub

Private Function MostraBloco(ByVal iCod As Long) As Integer
  Dim i As Integer
  MostraBloco = -1
  For i = 0 To cpBlocos.ListCount - 1
    If cpBlocos.ItemData(i) = iCod Then
      MostraBloco = i
      Exit For
    End If
  Next i
End Function

Private Function BlocoUnico() As Integer
  Dim i As Integer
  BlocoUnico = -1
  For i = 0 To cpBlocos.ListCount - 1
    If UCase(cpBlocos.List(i)) = "ÚNICO" Or UCase(cpBlocos.List(i)) = "UNICO" Then
      BlocoUnico = i
      Exit For
    End If
  Next i
End Function

