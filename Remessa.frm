VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Remessa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remessa"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "Remessa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton opPendente 
      Caption         =   "Pendente de baixa"
      Height          =   195
      Left            =   2040
      TabIndex        =   28
      Top             =   60
      Width           =   1635
   End
   Begin VB.ComboBox cpFiltro 
      Height          =   315
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   780
      Width           =   7815
   End
   Begin VB.CommandButton cmdMarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   315
      Left            =   9540
      TabIndex        =   25
      Top             =   4920
      Width           =   1755
   End
   Begin VB.OptionButton opAvulso 
      Caption         =   "Avulso"
      Height          =   195
      Left            =   1080
      TabIndex        =   24
      Top             =   60
      Width           =   795
   End
   Begin VB.OptionButton opNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   60
      Value           =   -1  'True
      Width           =   795
   End
   Begin VB.ComboBox cpTipoRemessa 
      Height          =   315
      ItemData        =   "Remessa.frx":000C
      Left            =   4260
      List            =   "Remessa.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4980
      Width           =   1695
   End
   Begin VB.ComboBox lsComandos 
      Height          =   1155
      Left            =   5280
      Style           =   1  'Simple Combo
      TabIndex        =   21
      Text            =   "lsComandos"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Remessa.frx":002B
      Height          =   3735
      Left            =   60
      OleObjectBlob   =   "Remessa.frx":003F
      TabIndex        =   4
      Top             =   1140
      Width           =   11235
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\Fontes\porto-real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5340
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar arquivo"
      Height          =   375
      Left            =   9540
      TabIndex        =   10
      Top             =   5760
      Width           =   1755
   End
   Begin VB.TextBox cpNomeArquivo 
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5820
      Width           =   7395
   End
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "..."
      Height          =   315
      Left            =   8280
      TabIndex        =   9
      Top             =   5820
      Width           =   555
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "Listar"
      Height          =   315
      Left            =   10140
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   360
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   11
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
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
      Left            =   1110
      TabIndex        =   0
      Top             =   360
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
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   8700
      TabIndex        =   2
      Top             =   360
      Width           =   1290
      _ExtentX        =   2275
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
      MaxLength       =   7
      TextMask        =   9
      RawText         =   9
      Mask            =   "##/####"
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpSequencia 
      Height          =   315
      Left            =   4740
      TabIndex        =   7
      Top             =   5400
      Width           =   1155
      _ExtentX        =   2037
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
   End
   Begin rdActiveText.ActiveText cpTotalDebito 
      Height          =   315
      Left            =   1260
      TabIndex        =   14
      Top             =   4980
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
      MaxLength       =   18
      Text            =   "R$ 0,00"
      TextMask        =   4
      RawText         =   4
      FloatFormat     =   2
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin MSComCtl2.DTPicker cpDebito 
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   16580611
      CurrentDate     =   41495
   End
   Begin MSComCtl2.DTPicker cpGeracao 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   5400
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   16580611
      CurrentDate     =   41495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Filtrar"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Remessa"
      Height          =   195
      Left            =   2940
      TabIndex        =   22
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data de geração"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5460
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Número sequêncial"
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   5460
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data do débito"
      Height          =   195
      Left            =   6120
      TabIndex        =   18
      Top             =   5460
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total do débito"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Diretório"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   435
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Do mês"
      Height          =   195
      Left            =   7350
      TabIndex        =   12
      Top             =   465
      Width           =   1260
   End
End
Attribute VB_Name = "Remessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colindex_lista = 4
Dim tbCondominio As Recordset
Dim nSoma As Double

Private Sub cmdArquivo_Click()
  cpNomeArquivo.Text = MinhaDll.sProcuraPorDiretorio("Diretório para geração do arquivo", Me, , True)
End Sub

Private Sub cmdGerar_Click()
  Dim iFree As Integer
  Dim rs As Recordset
  Dim sNome As String
  Dim nLote As Integer
  Dim sMen1 As String
  Dim sMen2 As String
  Dim valorTotal As Double
  Dim i As Integer
  Dim numRegArq As Integer
  Dim sLinha As String
  Dim nRemessa As Long
  Dim nomeRemessa As String
  Dim sPasta As String
  
  If Trim(cpNomeArquivo.Text) = "" Then
    MsgBox "Informe o local para geração do arquivo!", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  If Not sysFiles.FolderExists(cpNomeArquivo.Text) Then
    MsgBox "Informe o local para geração do arquivo!", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  If DBGrid1.ApproxCount = 0 Then
    MsgBox "Nenhum boleto para geração do arquivo!", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  i = 0
  With Data1.Recordset
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If !marca = "X" Then
          i = i + 1
        End If
        .MoveNext
      Loop
    End If
  End With
  
  
  If i = 0 Then
    MsgBox "Nenhum boleto marcado para geração do arquivo!", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  Resp = MsgBox("Gerar arquivo de remessa em '" & cpTipoRemessa.Text & "' de " & cpFiltro.Text & "?", vbQuestion + vbYesNo, "Gerar Arquivo")
  If Resp = vbNo Then
    Exit Sub
  End If
  
  Set rs = db.OpenRecordset("select * from condominio where codigo = " & cpCodigo.Text & " ", dbOpenDynaset)
  rs.MoveFirst
  iFree = FreeFile
  nLote = 0
  numRegArq = 0
  
  nomeRemessa = sFormataCaminho(cpNomeArquivo.Text) & "RE" & PadLeft(cpSequencia.Text, "0", 6) & ".TXT"
  Open nomeRemessa For Output As #iFree
    'header de arquivo remessa
    db.Execute "insert into remessa (remessa, data, confirmada, condominio) values ('" & PadLeft(cpSequencia.Text, "0", 6) _
        & "', #" & Format(Date, "MM/dd/yyyy") & "#, 'N', " & cpCodigo.Text & ") "
    nRemessa = RetornaCodigoUltimaRemessa(cpCodigo.Text)
    sLinha = "104" '1 - código do banco
    sLinha = sLinha & "0000" '2 - código do lote
    sLinha = sLinha & "0" '3 - tipo de registro
    sLinha = sLinha & Space(9) '4 - filler
    sLinha = sLinha & "2" '5 - Tipo de Registro beneficiário 2 Cnpj
    sLinha = sLinha & PadLeft(SoNumeros(rs!cnpjboleto), "0", 14) '6 - cnpj beneficiário
    sLinha = sLinha & String(20, "0") '7 - uso da caixa
    sLinha = sLinha & PadLeft(rs!agcedente, "0", 5) '8 - agência do beneficiário
    sLinha = sLinha & Trim(DigitoNosso(rs!agcedente)) '9 - digito verificador agencia benef.
    sNome = Right(rs!conta, 6)
    sPasta = PadLeft(sNome, "0", 6)
    sLinha = sLinha & PadLeft(sNome, "0", 6) '10 - código do cedente junto a caixa
    sLinha = sLinha & "00000000" '11, 12 - uso exclusivo da caixa
    sNome = rs!razaoboleto
    If Len(sNome) > 30 Then
      sNome = Left(sNome, 30)
    Else
      sNome = PadRight(sNome, " ", 30)
    End If
    sLinha = sLinha & sNome '13 - nome do beneficiário
    sLinha = sLinha & PadRight("CAIXA ECONOMICA FEDERAL", " ", 30) '14 - NOME DO BANCO
    sLinha = sLinha & Space(10) '15 - CNAB filler
    sLinha = sLinha & "1" '16 - código da remessa
    sLinha = sLinha & Format$(Now, "ddMMyyyy") '17 - data de geração do arquivo
    sLinha = sLinha & Format$(Now, "HHMMSS") '18 - hora de geração do arquivo
    sLinha = sLinha & PadLeft(cpSequencia.Text, "0", 6) '19 - NSA número sequencial
    sLinha = sLinha & "050" '20 - Número do layout
    sLinha = sLinha & "00000" '21 - densidade de geração do arquivo
    sLinha = sLinha & Space(20) '22 - uso exclusivo caixa filler
    If cpTipoRemessa.ListIndex = 0 Then
      sLinha = sLinha & PadRight("REMESSA-TESTE", " ", 20) '23 - situação do arquivo teste ou produção
    Else
      sLinha = sLinha & PadRight("REMESSA-PRODUCAO", " ", 20) '23 - situação do arquivo teste ou produção
    End If
    sLinha = sLinha & "    " 'versão do aplicatico caixa
    sLinha = sLinha & Space(25) 'cnab fiiler
    Print #iFree, sLinha
    numRegArq = numRegArq + 1
    db.Execute "insert into headerarquivo (id_remessa, linha) values (" & nRemessa & ", '" _
      & sLinha & "') "
    
    'header de lote
    nLote = nLote + 1
    sLinha = "104" '1 - código do banco
    sLinha = sLinha & PadLeft(nLote, "0", 4) '2 - lote de serviço - número do lote
    sLinha = sLinha & "1" '3 - tipo de registro
    sLinha = sLinha & "R" '4 - tipo de operação
    sLinha = sLinha & "01" '5 - tipo de serviço - 01 registrada 02 sem registro
    sLinha = sLinha & "00" '6 - filler
    sLinha = sLinha & "030" '7 - número da versão do layout
    sLinha = sLinha & " " '8 - CNAB filler
    sLinha = sLinha & "2" '9 - Tipo de Registro beneficiário 2 Cnpj
    sLinha = sLinha & PadLeft(SoNumeros(rs!cnpjboleto), "0", 15) '10 - cnpj beneficiário
    sLinha = sLinha & Right(rs!conta, 6) '11 - código do cedente junto a caixa
    sLinha = sLinha & String(14, "0") '11 - uso exclusivo caixa
    sLinha = sLinha & PadLeft(rs!agcedente, "0", 5) '12 - agência do beneficiário
    sLinha = sLinha & Trim(DigitoNosso(rs!agcedente)) '13 - digito verificador agencia benef.
    sLinha = sLinha & Right(rs!conta, 6) '14 - código do convênio
    sLinha = sLinha & "0000000" '15 - código do modelo do boleto
    sLinha = sLinha & "0" '16 - uso exclusivo caixa
    sNome = rs!razaoboleto
    If Len(sNome) > 30 Then
      sNome = Left(sNome, 30)
    Else
      sNome = PadRight(sNome, " ", 30)
    End If
    sLinha = sLinha & sNome '17 - nome do beneficiário
    sMen1 = Space(40)
    sMen2 = Space(40)
    sNome = ""  '!INST1 & " " & !INST2 & " " & !INST3 & " " & !INST4
    If Len(sNome) <= 40 Then
      sMen1 = PadRight(Trim(sNome), " ", 40)
    ElseIf Len(Trim(sNome)) > 40 Then
      sMen1 = Left(sNome, 40)
      sMen2 = Mid(sNome, 41)
      sMen1 = PadRight(Trim(sMen1), " ", 40)
      If Len(sMen2) > 40 Then
        sMen2 = Left(sMen2, 40)
      Else
        sMen2 = PadRight(Trim(sMen2), " ", 40)
      End If
    End If
    sLinha = sLinha & sMen1 '18 - mensagem 1
    sLinha = sLinha & sMen2 '19 - mensagem 2
    sLinha = sLinha & PadLeft(cpSequencia.Text, "0", 8) '20 - NSA número sequencial
    sLinha = sLinha & Format$(cpGeracao.Value, "ddMMyyyy") '21 - data de geração do arquivo
    sLinha = sLinha & "00000000" '22 - data de crédito
    sLinha = sLinha & Space(33) '33 - CNAB - filler
    Print #iFree, sLinha
    numRegArq = numRegArq + 1
    db.Execute "insert into headerlote (id_remessa, linha) values (" & nRemessa & ", '" _
      & sLinha & "') "
    
    'registro tipo P dados do título
    i = 0
    valorTotal = 0
    With Data1.Recordset
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          i = i + 1
          sLinha = "104" '1 - código do banco
          sLinha = sLinha & PadLeft(nLote, "0", 4) '2 - lote de serviço - número do lote
          sLinha = sLinha & "3" '3 - detalhe de lote
          sLinha = sLinha & PadLeft(i, "0", 5) '4 - núm sequencial do registro
          sLinha = sLinha & "P" '5 - código do registro
          sLinha = sLinha & " " '6 - filler
          sLinha = sLinha & !Comando '7 - código de movimento remessa
          sLinha = sLinha & PadLeft(rs!agcedente, "0", 5) '8 - agência do beneficiário
          sLinha = sLinha & Trim(DigitoNosso(rs!agcedente)) '9 - digito verificador agencia benef.
          sLinha = sLinha & Right(rs!conta, 6) '10 - código do convênio
          sLinha = sLinha & "00000000" '11 - uso exclusivo caixa
          sLinha = sLinha & "000" '12 - uso exclusivo caixa
          sLinha = sLinha & Left(!nosso, 2) '13 - carteira
          sLinha = sLinha & PadLeft(Mid$(!nosso, 3), "0", 15) ' 13 - identificação do título - noso número
          sLinha = sLinha & "1" '14 - código da carteira - 1 cobrança simples
          sLinha = sLinha & "2" '15 - tipo cobrança - 1 registrada 2 sem registro
          sLinha = sLinha & "2" '16 - tipo de documento 2 escritural
          sLinha = sLinha & "2" '17 - identificação da emissão
          sLinha = sLinha & "0" '18 - identificação da entrega
          sLinha = sLinha & PadLeft(!id, "0", 11) '19 - seu número de identificação do título
          sLinha = sLinha & "    " '19 - uso exclusivo caixa filler
          sLinha = sLinha & Format(!vcto, "ddmmyyyy") '20 - vencimento
          sLinha = sLinha & PadLeft(SoNumeros(Format(!corrigido, "#0.00")), "0", 15) '21 - valor do título
          valorTotal = valorTotal + Round(!corrigido, 2)
          sLinha = sLinha & "00000" '22 - agencia cobradora
          sLinha = sLinha & "0" '23 - digito verificador agencia
          sLinha = sLinha & "02" '24 espécie do título
          sLinha = sLinha & "N" '25 - identi título aceite o não
          sLinha = sLinha & Format(!Data, "ddmmyyyy") '26 - data de emissão
          sLinha = sLinha & "2" '27 - ident juros 2 taxa mensal
          sLinha = sLinha & Format(DateAdd("d", 1, !vcto), "ddmmyyyy") '28 - data início cobrança juros
          If rs!juros = 0 Then
            sLinha = sLinha & PadLeft(SoNumeros(Format(Parametros.juros, "#0.00000")), "0", 15) '29 - percentual de juros por mês
          Else
            sLinha = sLinha & PadLeft(SoNumeros(Format(rs!juros, "#0.00000")), "0", 15) '29 - percentual de juros por mês
          End If
          sLinha = sLinha & "0" '30 - código do desconto 0 sem
          sLinha = sLinha & "00000000" '31 - data do desconto
          sLinha = sLinha & String(15, "0") '32 - valor do desconto
          sLinha = sLinha & String(15, "0") '33 - valor do IOF
          sLinha = sLinha & String(15, "0") '34 - valor do abatimento
          sLinha = sLinha & PadLeft(!id, "0", 25) '35 - seu número de identificação do título
          sLinha = sLinha & "3" '36 - código de protesto
          sLinha = sLinha & "00" '37 - número de dias para protesto
          If (rs!dias > 0) Then
            sLinha = sLinha & "1" '38 - código para baixa devolução do título
            sLinha = sLinha & PadLeft(rs!dias, "0", 3) '39 - número de dias para caixa baixa/devolver o título
          Else
            sLinha = sLinha & "1" '38 - código para baixa devolução do título
            sLinha = sLinha & "120" '39 - número de dias para caixa baixa/devolver o título
          End If
          sLinha = sLinha & "09" '40 - código da moeda 09 real
          sLinha = sLinha & String(10, "0") '41 - uso da caixa filler
          sLinha = sLinha & " " '42 - CNAB - espaço
          Print #iFree, sLinha
          numRegArq = numRegArq + 1
          db.Execute "insert into registros (id_remessa, linha, numero) values (" & nRemessa & ", '" _
            & TiraAspa(sLinha) & "', " & i & ") "
          
          'Seguimento Q - pagador do título
          i = i + 1
          sLinha = "104"  '1 - código do banco
          sLinha = sLinha & PadLeft(nLote, "0", 4) '2 - lote de serviço - número do lote
          sLinha = sLinha & "3" '3 - detalhe de lote
          sLinha = sLinha & PadLeft(i, "0", 5) '4 - núm sequencial do registro
          sLinha = sLinha & "Q" '5 - código do registro
          sLinha = sLinha & " " '6 - filler
          sLinha = sLinha & !Comando '7 - código de movimento remessa
          sNome = SoNumeros(!cpf)
          If (Len(sNome) = 0) Then
            sLinha = sLinha & "0" '8 - tipo identificação 1 cpf 2 cnpj
          ElseIf (Len(sNome) < 14) Then
            sLinha = sLinha & "1" '8 - tipo identificação 1 cpf 2 cnpj
          Else
            sLinha = sLinha & "2" '8 - tipo identificação 1 cpf 2 cnpj
          End If
          sLinha = sLinha & PadLeft(sNome, "0", 15) '9 - número de inscrição do pagador
          sNome = !Nome
          If Len(sNome) > 40 Then
            sNome = Left(sNome, 40)
          Else
            sNome = PadRight(sNome, " ", 40)
          End If
          sLinha = sLinha & sNome '10 - nome do pagador
          sNome = !Ende
          If Len(sNome) > 40 Then
            sNome = Left(sNome, 40)
          Else
            sNome = PadRight(sNome, " ", 40)
          End If
          sLinha = sLinha & sNome '11 - endereço do pagador
          sNome = !Bair
          If Len(sNome) > 15 Then
            sNome = Left(sNome, 15)
          Else
            sNome = PadRight(sNome, " ", 15)
          End If
          sLinha = sLinha & sNome '12 - bairro do pagador
          sNome = !cep & ""
          If Len(sNome) > 0 Then
            sNome = SoNumeros(sNome)
          End If
          sLinha = sLinha & PadRight(sNome, "0", 8) '13 e 14 - cep do pagador
          sNome = !Cida
          If Len(sNome) > 15 Then
            sNome = Left(sNome, 15)
          Else
            sNome = PadRight(sNome, " ", 15)
          End If
          sLinha = sLinha & sNome '15 - cidade do pagador
          sNome = Trim(!Esta & "")
          If sNome = "" Then
            sNome = "MG"
          End If
          sLinha = sLinha & PadRight(sNome, " ", 2) '16 - uf do pagador
          'dados do avalista
          sLinha = sLinha & "0"
          sLinha = sLinha & String(15, "0")
          sLinha = sLinha & Space(40)
          'dados do avalista
          sLinha = sLinha & "000" '20 - banco correspondente
          sLinha = sLinha & Space(20) '21 - nosso número banco corresp.
          sLinha = sLinha & Space(8) '22 CNAB - filler
          Print #iFree, sLinha
          numRegArq = numRegArq + 1
          db.Execute "insert into registros (id_remessa, linha, numero) values (" & nRemessa & ", '" _
            & TiraAspa(sLinha) & "', " & i & ") "
          .MoveNext
        Loop
      End If
    End With
    
    'trailer de lote
    sLinha = "104"  '1 - código do banco
    sLinha = sLinha & PadLeft(nLote, "0", 4) '2 - coódigo do lote
    sLinha = sLinha & "5" '3 - tipo de registro
    sLinha = sLinha & Space(9) '4 CNAB filler
    sLinha = sLinha & PadLeft(i + 2, "0", 6) '5 - quantidade de registro do lote
    sLinha = sLinha & PadLeft(i / 2, "0", 6) '6 - quantidade de títulos no lote
    sLinha = sLinha & PadLeft(SoNumeros(Format$(valorTotal, "#0.00")), "0", 17) '7 valor total do lote
    sLinha = sLinha & String(46, "0") 'campos 8 a 11
    sLinha = sLinha & Space(148) 'campos 12 e 13
    Print #iFree, sLinha
    numRegArq = numRegArq + 1
    db.Execute "insert into trilerlote (id_remessa, linha) values (" & nRemessa & ", '" _
      & sLinha & "') "
    
    'trailer de arquivo
    numRegArq = numRegArq + 1
    sLinha = "104"  '1 - código do banco
    sLinha = sLinha & "9999" ' 2 - lote de serviço
    sLinha = sLinha & "9"  '3 - tipo de registro
    sLinha = sLinha & Space(9)  '4 - CNAB - filler
    sLinha = sLinha & PadLeft(nLote, "0", 6) '5 - quantidade de lotes no arquivo
    sLinha = sLinha & PadLeft(numRegArq, "0", 6) '6 - quantidade de registros no arquivo
    sLinha = sLinha & Space(211) 'campos 7 e 8
    Print #iFree, sLinha
    db.Execute "insert into trilerarquivo (id_remessa, linha) values (" & nRemessa & ", '" _
      & sLinha & "') "

  Close #iFree
  With rs
    .MoveFirst
    .Edit
    !dirremessa = cpNomeArquivo.Text
    .Update
    gravaUltimaRemessaConvenio Right(!conta, 6), cpSequencia.Text
  End With
  sPasta = sFormataCaminho(App.Path) & sPasta
  If Not sysFiles.FolderExists(sPasta) Then
    sysFiles.CreateFolder sPasta
  End If
  sysFiles.CopyFile nomeRemessa, sFormataCaminho(sPasta)
  MsgBox "Arquivo de remessa gerado com sucesso!", vbInformation + vbOKOnly, "Geração de arquivo remessa"
End Sub

Private Sub cmdListar_Click()
  Dim nSoma As Double
  Dim rs As Recordset
  Dim rsLocal As Recordset
  
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If opPendente.Value = False Then
    If Not IsDate("25/" & cpMes.Text) Then
      MsgBox "O mês/ano informado não é válido.", vbInformation + vbOKOnly, "Aviso"
      cpMes.SetFocus
      Exit Sub
    End If
  End If
  
  cmdListar.Enabled = False
  cmdMarcarTodos.Enabled = False
  cmdGerar.Enabled = False
  
  Call SetTopMostWindow(Progresso.hWnd, True)
  Progresso.Caption = "Progresso"
  Progresso.Label1.Caption = "Processando boletos..."
  Progresso.Barra.Value = 0
  Progresso.Show
  Progresso.Refresh
  
  cpFiltro.Clear
  cpFiltro.AddItem "Todos"
  cpFiltro.ListIndex = 0
    
  dbLocal.Execute "delete from boletos_auxiliar "
  Data1.Refresh
  DoEvents
  If opNormal.Value Then
    Set rs = db.OpenRecordset("select * from boletos where cond = " & cpCodigo.Text & " and tran = '" & cpMes.Text & "' " _
      & "union select * from boletos where cond = " & cpCodigo.Text & " and  tran <>  '" & cpMes.Text & "' " _
      & "and (idstatus = 4 or idstatus = 9 or idstatus = 10) order by bole;", dbOpenDynaset)
  ElseIf opPendente.Value Then
    Set rs = db.OpenRecordset("select * from boletos where cond = " & cpCodigo.Text & " and " _
      & "(idstatus = 4 or idstatus = 5 or idstatus = 9 or idstatus = 10) order by bole;", dbOpenDynaset)
  Else
    Set rs = db.OpenRecordset("select * from boletos where left(tran,2) = 'AV' and cond = " & cpCodigo.Text & " and format(vcto,'MM/yyyy') = '" & cpMes.Text & "' " _
      & "union select * from boletos where cond = " & cpCodigo.Text & " and format(vcto,'MM/yyyy') <> '" & cpMes.Text & "' " _
      & "and (idstatus = 4 or idstatus = 9 or idstatus = 10) order by bole;", dbOpenDynaset)
  End If
  
  Set rsLocal = dbLocal.OpenRecordset("boletos_auxiliar", dbOpenTable)
  With rs
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        nSoma = nSoma + !corrigido
        rsLocal.AddNew
        rsLocal!Condominio = !Condominio
        rsLocal!CGC = !CGC
        rsLocal!vcto = !vcto
        rsLocal!MENS = !MENS
        rsLocal!EXTR = !EXTR
        rsLocal!Data = !Data
        rsLocal!valr = !valr
        rsLocal!corrigido = !corrigido
        rsLocal!COTA = !COTA
        rsLocal!cdsc = !cdsc
        rsLocal!Nome = !Nome
        rsLocal!cpf = !cpf
        rsLocal!Ende = !Ende
        rsLocal!Bair = !Bair
        rsLocal!Cida = !Cida
        rsLocal!Esta = !Esta
        rsLocal!cep = !cep
        rsLocal!tran = !tran
        rsLocal!DIGITAVAL = !DIGITAVAL
        rsLocal!agcedente = !agcedente
        rsLocal!carteira = !carteira
        rsLocal!Mensagem = !Mensagem
        rsLocal!pago = !pago
        rsLocal!CDBARRA = !CDBARRA
        rsLocal!CODI = !CODI
        rsLocal!texto = !texto
        rsLocal!INST1 = !INST1
        rsLocal!INST2 = !INST2
        rsLocal!INST3 = !INST3
        rsLocal!INST4 = !INST4
        rsLocal!Banco = !Banco
        rsLocal!CdBanco = !CdBanco
        rsLocal!CANCELADO = !CANCELADO
        rsLocal!cond = !cond
        rsLocal!acumulado = !acumulado
        rsLocal!bole = !bole
        rsLocal!id = !id
        rsLocal!nosso = !nosso
        rsLocal!desconto = !desconto
        rsLocal!idStatus = !idStatus
        If rsLocal!idStatus = 0 Or IsNull(rsLocal!idStatus = 0) Then
          rsLocal!Comando = "01"
        Else
          rsLocal!Comando = RetornaComando(rsLocal!idStatus)
        End If
        rsLocal!diasprotesto = 0
        If opAvulso.Value Then
          rsLocal!marca = ""
        Else
          rsLocal!marca = "X"
        End If
        rsLocal.Update
        Progresso.Barra.Value = Int(.PercentPosition)
        Progresso.Refresh
        .MoveNext
      Loop
    End If
  End With
  
  Set rs = Nothing
  Set rsLocal = Nothing
  
  Progresso.Label1.Caption = "Calculando valores..."
  Progresso.Refresh
  
  Data1.RecordSource = "select * from boletos_auxiliar order by bole;"
  Data1.Refresh
  DBGrid1.ReBind
  nSoma = 0
  With Data1.Recordset
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        incluirFiltros RetornaDescricaoComando(!Comando)
        nSoma = nSoma + !corrigido
        Progresso.Barra.Value = Int(.PercentPosition)
        Progresso.Refresh
        .MoveNext
      Loop
    End If
  End With
  
  cpTotalDebito.Text = nSoma

  cmdListar.Enabled = True
  cmdMarcarTodos.Enabled = True
  cmdGerar.Enabled = True
  
  Call SetTopMostWindow(Progresso.hWnd, False)
  Unload Progresso
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpNome.Text = !Nome
        cpCodigo.Text = !Codigo
        cpSequencia.Text = retornaUltimaRemessaConvenio(Right(!conta, 6), !Codigo)
        If cpSequencia.Text > 999999 Then
          cpSequencia.Text = 1
        End If
        cpNomeArquivo.Text = !dirremessa & ""
        If cpMes.Enabled Then cpMes.SetFocus
      End If
    End With
  End If
End Sub

Private Sub cmdMarcarTodos_Click()
  If cmdMarcarTodos.Caption = "Desmarcar Todos" Then
    With Data1.Recordset
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          .Edit
          !marca = ""
          .Update
          .MoveNext
        Loop
      End If
    End With
    cmdMarcarTodos.Caption = "Marcar Todos"
    cpTotalDebito.Text = "0,00"
  Else
    nSoma = 0
    With Data1.Recordset
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          .Edit
          !marca = "X"
          .Update
          nSoma = nSoma + !valr
          .MoveNext
        Loop
      End If
    End With
    cmdMarcarTodos.Caption = "Desmarcar Todos"
    cpTotalDebito.Text = nSoma
  End If
End Sub

Private Sub cpFiltro_Click()
  If cpFiltro.Text = "Todos" Then
    Data1.RecordSource = "select * from boletos_auxiliar order by bole;"
  Else
    Data1.RecordSource = "select * from boletos_auxiliar where comando = '" & RetornaComandoDescricao(cpFiltro.Text) & "' order by bole;"
  End If
  Data1.Refresh
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdListar.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
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
          If cpMes.Enabled Then cpMes.SetFocus
          cpSequencia.Text = retornaUltimaRemessaConvenio(Right(!conta, 6), !Codigo)
          If cpSequencia.Text > 999999 Then
            cpSequencia.Text = 1
          End If
          cpNomeArquivo.Text = !dirremessa & ""
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
            cpSequencia.Text = retornaUltimaRemessaConvenio(Right(!conta, 6), !Codigo)
            If cpSequencia.Text > 999999 Then
              cpSequencia.Text = 1
            End If
            cpNomeArquivo.Text = !dirremessa & ""
          End If
        End With
        If cpMes.Enabled Then cpMes.SetFocus
      End If
    End If
  End If
End Sub

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  If ColIndex = colindex_lista Then  'força a seleção da lista
     Cancel = True
     DBGrid1_ButtonClick (ColIndex)
  End If
End Sub

Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
  Dim coluna As Column
  
  'mostra a lista abaixo da coluna selecionada
  If ColIndex = colindex_lista Then
     Set coluna = DBGrid1.Columns(ColIndex)
     With lsComandos
        .Left = DBGrid1.Left + coluna.Left
        .Top = DBGrid1.Top + DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.RowHeight
        '.Width = coluna.Width + 15
        .ListIndex = 0
        .Visible = True
        .ZOrder 0
        .SetFocus
     End With
  End If
End Sub

Private Sub DBGrid1_Scroll(Cancel As Integer)
   'oculta a lista se rolar o grid
   lsComandos.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 117 Then
    If (DBGrid1.Columns(6).Text = "X") Then
      DBGrid1.Columns(6).Text = ""
      cpTotalDebito.Text = cpTotalDebito.Text - Data1.Recordset!valr
    Else
      DBGrid1.Columns(6).Text = "X"
      cpTotalDebito.Text = cpTotalDebito.Text + Data1.Recordset!valr
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Data1.DatabaseName = sFormataCaminho(App.Path) & "supersoft.mdb"
  PreencheLista
  cpGeracao.Value = Now
  cpTipoRemessa.ListIndex = 1
  Refresh
  KeyPreview = True
End Sub

Private Sub PreencheLista()
  Dim rs As Recordset
  Set rs = db.OpenRecordset("comando", dbOpenTable)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        lsComandos.AddItem !Codigo & " " & !Descricao
        .MoveNext
      Loop
    End If
  End With
  Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub

Private Sub lsComandos_Click()
  'assume o valor clicado
  lsComandos_KeyPress vbKeyReturn
End Sub

Private Sub lsComandos_DblClick()
  'assume o valor clicado
  lsComandos_KeyPress vbKeyReturn
End Sub

Private Sub lsComandos_KeyPress(KeyAscii As Integer)
  'verifica a tecla pressionada e dispara ação pertinente
  Select Case KeyAscii
     Case vbKeyReturn
        DBGrid1.Columns(colindex_lista).Text = Left(lsComandos.Text, 2)
        DBGrid1.Columns(6).Text = "X"
        lsComandos.Visible = False
     Case vbKeyEscape
        lsComandos.Visible = False
  End Select
End Sub

Private Sub lsComandos_LostFocus()
   'oculta a lista ao perder foco
   lsComandos.Visible = False
End Sub

Private Sub opAvulso_Click()
  If opAvulso.Value = False Then
    Label2.Caption = "Do mês"
  Else
    Label2.Caption = "Vcto em"
  End If
  cpMes.Enabled = True
End Sub

Private Sub opNormal_Click()
  If opNormal.Value Then
    Label2.Caption = "Do mês"
  Else
    Label2.Caption = "Vcto em"
  End If
  cpMes.Enabled = True
End Sub

Private Sub incluirFiltros(ByVal sFiltro As String)
  Dim h As Integer
  Dim tem As Boolean
  tem = False
  For h = 0 To cpFiltro.ListCount - 1
    If (cpFiltro.List(h) = sFiltro) Then
      tem = True
      Exit For
    End If
  Next h
  If Not tem Then
    cpFiltro.AddItem (sFiltro)
  End If
End Sub

Public Function RetornaDescricaoComando(ByVal Comando As String) As String
  Dim ret As String
  Dim rs As Recordset
  
  ret = ""
  Set rs = db.OpenRecordset("select * from comando where codigo = '" & Comando & "';", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      ret = !Descricao
    End If
  End With
  Set rs = Nothing
  RetornaDescricaoComando = ret
End Function

Private Sub opPendente_Click()
  Label2.Caption = "N/T"
  cpMes.Enabled = False
End Sub

Public Function RetornaComandoDescricao(ByVal Descr As String) As String
  Dim ret As String
  Dim rs As Recordset
  
  ret = ""
  Set rs = db.OpenRecordset("select * from comando where descricao = '" & Descr & "';", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      ret = !Codigo
    End If
  End With
  Set rs = Nothing
  RetornaComandoDescricao = ret
End Function

