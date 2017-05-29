VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Carga 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3510
   ClientLeft      =   1350
   ClientTop       =   2865
   ClientWidth     =   5610
   ControlBox      =   0   'False
   Icon            =   "Carga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4620
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   210
      Left            =   165
      TabIndex        =   7
      Top             =   3210
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   540
      Top             =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Super Soft"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   570
      Index           =   2
      Left            =   2130
      TabIndex        =   6
      Top             =   120
      Width           =   2355
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   150
      Picture         =   "Carga.frx":000C
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Super Soft"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   180
      Width           =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando componentes do sistema. Aguarde..."
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   2985
      Width           =   3450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "31 99556-8707"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   2235
      Width           =   1290
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido em Visual Basic 6.0 por J && F Software's Ltda "
      Height          =   390
      Left            =   1440
      TabIndex        =   2
      Top             =   1545
      Width           =   2730
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Windows"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   1995
      Width           =   1035
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa de computador é protegido por tratados internacionais e leis de Copyright© descritos na caixa Sobre..."
      Height          =   405
      Left            =   165
      TabIndex        =   0
      Top             =   2535
      Width           =   5430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   180
      X2              =   5475
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   165
      X2              =   5475
      Y1              =   2955
      Y2              =   2955
   End
End
Attribute VB_Name = "Carga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim tbParametros As Recordset


Private Sub Form_Load()
  Show
  Timer1.Enabled = True
End Sub

Private Sub Iniciar()
  Dim TxjMes         As Double
  Dim StrSel2        As String
  Dim StrFile As String
  
  Inifile = sFormataCaminho(App.Path) & "supersoft.ini"
  
  Set MinhaDll = New DLLsrv
  Set sysFiles = CreateObject("Scripting.FileSystemObject")
  
  If Not sysFiles.FileExists(Inifile) Then
    MsgBox "O Arquivo de configaração não foi encontrado! O Programa não pode prosseguir.", vbCritical + vbOKOnly, "Aviso"
    Set MinhaDll = Nothing
    End
  End If
  
  Barra.Value = 1
  Barra.Value = 18
  Barra.Value = 20
  Barra.Value = 22
  Barra.Value = 25
  Barra.Value = 29
  Barra.Value = 31
  Barra.Value = 33
  Barra.Value = 35
  Barra.Value = 38
  Parametros.dados = LerIni("Config", "Dados", Inifile)
  Barra.Value = 40
  Barra.Value = 43
  Barra.Value = 45
  Barra.Value = 48
  Barra.Value = 50
  
  Call SaveDataBase(sFormataCaminho(App.Path) & "supersoft.mdb", False)
  
  If Parametros.dados = "" Then
    MsgBox "Local da base de dados não informado.", vbCritical + vbOKOnly, "Aviso"
    Set MinhaDll = Nothing
    End
  Else
Dinovo:
    StrFile = Trim(Parametros.dados)
    If sysFiles.FileExists(StrFile) Then
      sVerDb = VersaoBancoDados(StrFile)
      CriarBancoDeDados Parametros.dados, sVerDb
      If Not AbrirArquivos(StrFile) Then
        Configuracao.Show 1
        If tbParametros Is Nothing Then
          MsgBox "Houve um erro na abertura dos arquivos! O Programa não pode prosseguir.", vbCritical + vbOKOnly, "Aviso"
          Set MinhaDll = Nothing
          End
        End If
      End If
    Else
      Resp = MsgBox("Banco de dados '" & StrFile & "' não encontrado. Localizar o arquivo?", vbCritical + vbYesNo, "Aviso")
      If Resp = vbYes Then
        Configuracao.Show 1
        Parametros.dados = LerIni("Config", "Dados", Inifile)
        GoTo Dinovo
      Else
        Resp = MsgBox("Criar o banco de dados do sistema?", vbCritical + vbYesNo, "Aviso")
        If Resp = vbYes Then
          Call sysFiles.CopyFile(sFormataCaminho(App.Path) & "dados.crx", sFormataCaminho(App.Path) & "dados_1.0.1.bco", True)
          CriarBancoDeDados Parametros.dados, "1.0.0"
          GoTo Dinovo
        Else
          Set MinhaDll = Nothing
          End
          Exit Sub
        End If
      End If
    End If
  End If
  Barra.Value = 91
  Set tbParametros = db.OpenRecordset("parametros", dbOpenTable)
  With tbParametros
    If .RecordCount > 0 Then
      .MoveFirst
      If VerificaRegistro(tbParametros) = -1 Then
        MsgBox "Os dados do sistema foram alterados! Programa terminado.", vbInformation + vbOKOnly, "Aviso"
        CloseDataBase
        Set MinhaDll = Nothing
        End
      End If
    Else
      Registro.Show 1
    End If
  End With
  tbParametros.MoveFirst
  With Parametros
    Barra.Value = 95
    .Carencia = IIf(IsNull(tbParametros("Carencia").Value), 0, tbParametros("Carencia").Value)
    .juros = IIf(IsNull(tbParametros("Juros").Value), 0, tbParametros("Juros").Value)
    .Mensalidade = CSng(LerIni("Prazo", "Mensalidade", Inifile))
    .Multa = IIf(IsNull(tbParametros("DespesaAd").Value), 0, tbParametros("DespesaAd").Value)
    Barra.Value = 97
    .Periodo = IIf(IsNull(tbParametros("OutrasDesp").Value), 0, tbParametros("OutrasDesp").Value)
    .Primeira = LerIni("Mensagem", "Linha1", Inifile)
    .Segunda = LerIni("Mensagem", "Linha2", Inifile)
    .Correcao = IIf(IsNull(tbParametros!Cofins), 0, tbParametros!Cofins)
    .Visitas = CInt(LerIni("Prazo", "NumeroDeVisitas", Inifile))
    .Cortesia = IIf(IsNull(tbParametros("Linhas").Value), "", tbParametros("Linhas").Value)
    .CobDia = IIf(IsNull(tbParametros("Pis").Value), "", tbParametros("Pis").Value)
    .Empresa = tbParametros("Empresa").Value
    vbLimiteDias = IIf(IsNull(tbParametros("DiasAno").Value), 0, tbParametros("DiasAno").Value)
    SenhaSupervisor = IIf(IsNull(tbParametros("Supervisor").Value), "12345", tbParametros("Supervisor").Value)
    Empresa = tbParametros("Empresa").Value
    vbSede = tbParametros("Cidade").Value
    Barra.Value = 100
  End With
  With tbParametros
    .MoveFirst
    vbEmpresa.bairro = IIf(IsNull(!bairro), "", AcertaLetras(!bairro))
    vbEmpresa.cep = IIf(IsNull(!cep), "", !cep)
    vbEmpresa.Cidade = IIf(IsNull(!Cidade), "", AcertaLetras(!Cidade))
    vbEmpresa.Cnpj = IIf(IsNull(!Cnpj), "", !Cnpj)
    vbEmpresa.Empresa = IIf(IsNull(!Empresa), "", AcertaLetras(!Empresa))
    vbEmpresa.endereco = IIf(IsNull(!endereco), "", AcertaLetras(!endereco))
    vbEmpresa.estado = IIf(IsNull(!estado), "", !estado)
    vbEmpresa.Fantazia = ""
    vbEmpresa.Fones = IIf(IsNull(!Telefone), "", !Telefone)
    vbEmpresa.Inscricao = IIf(IsNull(!Inscricao), "", !Inscricao)
  End With
  'CalculaJuros
  DoEvents
  If sysFiles.FileExists(sFormataCaminho(App.Path) & "executar.txt") Then
    Dim iExec As Integer
    iExec = FreeFile
    Open sFormataCaminho(App.Path) & "executar.txt" For Input As #iExec
    With ExecScript
      .Text1.Text = Input(LOF(iExec), iExec)
      .Show 1
    End With
    Close #iExec
    DoEvents
    sysFiles.DeleteFile sFormataCaminho(App.Path) & "executar.txt", True
  End If
  Set tbParametros = Nothing
  DoEvents
  Principal.Show
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Iniciar
End Sub
