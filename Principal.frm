VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Principal 
   AutoShowChildren=   0   'False
   BackColor       =   &H00E0E0E0&
   Caption         =   "Super Sfot"
   ClientHeight    =   3465
   ClientLeft      =   1140
   ClientTop       =   2760
   ClientWidth     =   6315
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar Informe 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3135
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1110
      Top             =   2010
   End
   Begin VB.Menu MenuManutencoa 
      Caption         =   "&Manutenção"
      Begin VB.Menu mnuCadBlocos 
         Caption         =   "&Blocos"
      End
      Begin VB.Menu MenuCondominio 
         Caption         =   "&Condomínio"
      End
      Begin VB.Menu MenuCidades 
         Caption         =   "Ci&dades"
      End
      Begin VB.Menu mnuCadHistorico 
         Caption         =   "&Histórico padrão"
      End
      Begin VB.Menu MenuAssociados 
         Caption         =   "&Inquilinos"
      End
      Begin VB.Menu MenuParamentros 
         Caption         =   "Parâ&mentros"
      End
      Begin VB.Menu mnuManuEmailServer 
         Caption         =   "&Servidor de e-mail"
      End
      Begin VB.Menu MenuTipoDespesa 
         Caption         =   "&Tipo de despesa"
      End
      Begin VB.Menu mnuLeiturasTipo 
         Caption         =   "Ti&pos de leitura"
      End
      Begin VB.Menu MenuTipoImovel 
         Caption         =   "Tip&os do imóvel"
      End
      Begin VB.Menu MenuUsuarios 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu MenuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSair 
         Caption         =   "Sai&r"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu MenuDespesas 
      Caption         =   "&Despesas"
      Begin VB.Menu MenuDespesasCondominio 
         Caption         =   "&Despesas do condomínio"
      End
      Begin VB.Menu mnuDespAssociado 
         Caption         =   "&Incluir despesa por inquilino"
      End
      Begin VB.Menu mnusepdesp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDescParaInq 
         Caption         =   "D&esconto para inquilino"
      End
      Begin VB.Menu mnuDespFixa 
         Caption         =   "De&spesa fixa mensal"
      End
   End
   Begin VB.Menu MenuSupRotina 
      Caption         =   "Rotina &Diária"
      Begin VB.Menu mnuRotRemessa 
         Caption         =   "&Arquivo de remessa CEF"
      End
      Begin VB.Menu mnuRetornoCef 
         Caption         =   "&Arquivo de retorno CEF"
      End
      Begin VB.Menu MenuBoletosAvulsos 
         Caption         =   "Bo&letos Avulsos"
      End
      Begin VB.Menu MenuBoletos 
         Caption         =   "&Boletos (Gerar/Imprimir)"
      End
      Begin VB.Menu mnuDiaBolGenericos 
         Caption         =   "Boletos &genéricos"
      End
      Begin VB.Menu mnuRotCheques 
         Caption         =   "Emi&ssão de cheques"
      End
      Begin VB.Menu mnuDiarEnviaEmail 
         Caption         =   "&Enviar E-mail"
      End
      Begin VB.Menu mnuLeituraLancamento 
         Caption         =   "La&nçamentos de leituras"
      End
      Begin VB.Menu MenuCobranca 
         Caption         =   "&Montagem/Distribuição das despesas"
      End
      Begin VB.Menu mnuDiariaParcel 
         Caption         =   "&Parcelementos"
      End
      Begin VB.Menu MenuQuitar 
         Caption         =   "&Quitação de boletos"
      End
   End
   Begin VB.Menu MenuSupRel 
      Caption         =   "&Relatório"
      Begin VB.Menu MenuRelAtrazados 
         Caption         =   "&Boletos em atraso"
      End
      Begin VB.Menu mnuRelBolQuit 
         Caption         =   "Boletos &quitados"
      End
      Begin VB.Menu mnuRelChecaBoleto 
         Caption         =   "Checar boletos"
      End
      Begin VB.Menu MenuRelCondominio 
         Caption         =   "&Condomínio"
      End
      Begin VB.Menu mnuRelDescontos 
         Caption         =   "Desco&ntos"
      End
      Begin VB.Menu mnuRelDespFixa 
         Caption         =   "Despesas &fixas"
      End
      Begin VB.Menu mnuRelDespIndividual 
         Caption         =   "Despesa &individual"
      End
      Begin VB.Menu mnuDespesasMes 
         Caption         =   "&Despesas do mês..."
      End
      Begin VB.Menu mnuRelDistInquilinos 
         Caption         =   "Dis&tribuição para inquilinos"
      End
      Begin VB.Menu mnuRelEmailEnviados 
         Caption         =   "E-&mails enviados"
      End
      Begin VB.Menu MenuRelExtrato 
         Caption         =   "&Extrato do inquilino"
      End
      Begin VB.Menu mnuRelLeituras 
         Caption         =   "Leitura&s lançadas"
      End
      Begin VB.Menu mnuListaBoletos 
         Caption         =   "Lista de &boletos"
      End
      Begin VB.Menu mnuRelCpfsProblemas 
         Caption         =   "Lista CPF/CNPJs com problemas"
      End
      Begin VB.Menu mnuListaInq 
         Caption         =   "&Lista de inquilinos"
      End
      Begin VB.Menu mnuRelListaProp 
         Caption         =   "Lista de &proprietários"
      End
      Begin VB.Menu mnuRelRemessas 
         Caption         =   "Remessas enviadas"
      End
   End
   Begin VB.Menu MenuSupUtil 
      Caption         =   "&Utilitários"
      Begin VB.Menu MenuCancelaBoleto 
         Caption         =   "C&ancelar boletos por mês geração"
      End
      Begin VB.Menu mnuCancelPorNumero 
         Caption         =   "Cancelar &boleto por número"
      End
      Begin VB.Menu mnuUtilCancelaGenerico 
         Caption         =   "Cancelar boleto genérico"
      End
      Begin VB.Menu MenuCorrigirAtrazo 
         Caption         =   "&Corrigir Lançamentos em Atraso"
      End
      Begin VB.Menu MenuEstorno 
         Caption         =   "&Estorno de Pagamento"
      End
      Begin VB.Menu mnuUtiInfoBoleto 
         Caption         =   "&Informação do boleto"
      End
      Begin VB.Menu mnuUtilReimpBol 
         Caption         =   "&Reimpressão de boleto vencido"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu mnuLocalDados 
         Caption         =   "&Local do banco de dados"
      End
      Begin VB.Menu MenuReorganizar 
         Caption         =   "Reor&ganizar Banco de Dados"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSobre 
         Caption         =   "&Sobre..."
      End
   End
   Begin VB.Menu MenuJanelas 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu MenuCascata 
         Caption         =   "Em &cascata"
      End
      Begin VB.Menu MenuLado 
         Caption         =   "&Lado a lado"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public WithEvents balao As clsSysTray
Attribute balao.VB_VarHelpID = -1

Private Sub balao_RightClick()
  PopupMenu MenuManutencoa
End Sub

Private Sub Informe_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  balao.MouseMove Button, X, Me
End Sub

Private Sub MDIForm_Load()
  Caption = Replace(Prog, "%", App.Major & "." & App.Minor & "." & App.Revision)
  Set balao = New clsSysTray
  balao.Init Principal, Caption
  Timer1.Enabled = True
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
  balao.MouseMove Button, X, Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Resp = MsgBox("Confirma a saída do sistema?", vbQuestion + vbYesNo, "Sair")
  If Resp = vbNo Then
    Cancel = True
  Else
'    Call CloseDataBase
    Set MinhaDll = Nothing
    Set balao = Nothing
    DoEvents
    End
  End If
End Sub

Private Sub MenuAssociados_Click()
  Associados.Show
End Sub

Private Sub MenuBackup_Click()
  Backup.Show 1
End Sub

Private Sub MenuBoletos_Click()
  Boletos.Show
End Sub

Private Sub MenuBoletosAnteriores_Click()
'  BoletosOld.Show
End Sub

Private Sub MenuBoletosAvulsos_Click()
  BoletosAvulso.Show
  BoletosAvulso.ZOrder 0
End Sub

Private Sub MenuCancelaBoleto_Click()
  CanBol.Show
End Sub

Private Sub MenuCidades_Click()
  CadCidades.Show
End Sub

Private Sub MenuCobranca_Click()
  RelDesp.Caption = "Distribuição das despesas"
  RelDesp.cmdPrint.Caption = "Distribuir"
  RelDesp.Show
  RelDesp.ZOrder 0
End Sub

Private Sub MenuCondominio_Click()
  Condominio.Show
End Sub

Private Sub MenuCorrigirAtrazo_Click()
  AtualizaAtrazados.Show
  AtualizaAtrazados.ZOrder 0
End Sub

Private Sub MenuDespesasCondominio_Click()
  Despesas.Show
End Sub

Private Sub MenuEstorno_Click()
  Extorno.Show
End Sub

Private Sub MenuParamentros_Click()
Dinovo:
  Supervisor.Show vbModal
  If sSuper = SenhaSupervisor Then
    Parame.Show
  ElseIf (sSuper <> "") Then
    MsgBox "Esta não é uma senha correta! Por favor verifique e tente novamente.", vbCritical + vbOKOnly, "Aviso"
    GoTo Dinovo
  End If
End Sub

Private Sub MenuQuitar_Click()
  Quitar.Show
End Sub

Private Sub MenuRelAtrazados_Click()
  RelBolAt.Show
End Sub

Private Sub MenuRelCondominio_Click()
  RelCondo.Show
End Sub

Private Sub MenuRelExtrato_Click()
  Extrato.Show
End Sub

Private Sub MenuReorganizar_Click()
  
  Dim StrFile As String
  Dim vbRet   As Double
  
  ComDb.Show vbModeless
  StrFile = Parametros.dados
  CloseDataBase
  DoEvents
  vbRet = SaveDataBase(StrFile)
  If vbRet <> 0 Then
    MsgBox Error$(vbRet), vbInformation + vbOKOnly, "Erro"
  Else
    MsgBox "Banco de Dados reorganizado com sucesso!", vbExclamation + vbOKOnly, "Banco de dados"
  End If
  Call AbrirArquivos(StrFile, 2)
  Unload ComDb

End Sub

Private Sub MenuSair_Click()
  Unload Me
End Sub

Private Sub MenuSobre_Click()
  Sobre.Show 1
End Sub

Private Sub MenuTipoDespesa_Click()
  tpDesp.Show
End Sub

Private Sub MenuTipoImovel_Click()
  Tipos.Show
End Sub

Private Sub MenuUsuarios_Click()
  Usuarios.Show
End Sub

Private Sub mnuCadBlocos_Click()
  CadBlocos.Show
  CadBlocos.ZOrder 0
End Sub

Private Sub mnuCadHistorico_Click()
  Cadhistorico.Show
  Cadhistorico.ZOrder 0
End Sub

Private Sub mnuCancelPorNumero_Click()
  CancelaBoleto.Show
End Sub

Private Sub mnuConferiCpfs_Click()
  Dim rs As Recordset
  
  dbLocal.Execute "delete from CONFERICPF;"
  Set rs = dbLocal.OpenRecordset("CONFERICPF", dbOpenTable)
  
End Sub

Private Sub mnuDescParaInq_Click()
  Descontos.Show
  Descontos.ZOrder 0
End Sub

Private Sub mnuDespAssociado_Click()
  PorFora.Show
End Sub

Private Sub mnuDespesasMes_Click()
  RelDespesa.Show
End Sub

Private Sub mnuDespFixa_Click()
  DespesaFixa.Show
  DespesaFixa.ZOrder 0
End Sub

Private Sub mnuDiaBolGenericos_Click()
  BoletoGenerico.Show
  BoletoGenerico.ZOrder 0
End Sub

Private Sub mnuDiarEnviaEmail_Click()
  EnviarEmail.Show
  EnviarEmail.ZOrder 0
End Sub

Private Sub mnuDiariaParcel_Click()
  Parcelamento.Show
  Parcelamento.ZOrder 0
End Sub

Private Sub mnuLeituraLancamento_Click()
  LancamentoLeitura.Show
  LancamentoLeitura.ZOrder 0
End Sub

Private Sub mnuLeiturasTipo_Click()
  CadTipoLeitura.Show
  CadTipoLeitura.ZOrder 0
End Sub

Private Sub mnuListaBoletos_Click()
  RelAsBoleto.Show
End Sub

Private Sub mnuListaInq_Click()
  RelAssociados.Show
End Sub

Private Sub mnuLocalDados_Click()
  Configuracao.Show 1
End Sub

Private Sub mnuManuEmailServer_Click()
  ServidorEmail.Show
  ServidorEmail.ZOrder 0
End Sub

Private Sub mnuRelBolQuit_Click()
  RelRecebidos.Show
  RelRecebidos.ZOrder 0
End Sub

Private Sub mnuRelChecaBoleto_Click()
  frmBoletos.Show
  frmBoletos.ZOrder 0
End Sub

Private Sub mnuRelCpfsProblemas_Click()
  RelCpfs.Show
  RelCpfs.ZOrder 0
End Sub

Private Sub mnuRelDescontos_Click()
  RelDescontos.Show
  RelDescontos.ZOrder 0
End Sub

Private Sub mnuRelDespFixa_Click()
  RelDespFixa.Show
  RelDespFixa.ZOrder 0
End Sub

Private Sub mnuRelDespIndividual_Click()
  DespesaIndividual.Show
  DespesaIndividual.ZOrder 0
End Sub

Private Sub mnuRelDistInquilinos_Click()
  RelDesp.Caption = "Relatório da distribuição por inquilinos"
  RelDesp.cmdPrint.Caption = "Imprimir"
  RelDesp.Show
  RelDesp.ZOrder 0
End Sub

Private Sub mnuRelEmailEnviados_Click()
  RelEmailEnv.Show
  RelEmailEnv.ZOrder 0
End Sub

Private Sub mnuRelLeituras_Click()
  RelLeitura.Show
  RelLeitura.ZOrder 0
End Sub

Private Sub mnuRelListaProp_Click()
  RelProprietario.Show
  RelProprietario.ZOrder 0
End Sub

Private Sub mnuRelRemessas_Click()
  RelRemessas.Show
  RelRemessas.ZOrder 0
End Sub

Private Sub mnuRetornoCef_Click()
  RetornoCEF.Show
End Sub

Private Sub mnuRotCheques_Click()
  Cheques.Show
  Cheques.ZOrder 0
End Sub

Private Sub mnuRotRemessa_Click()
  Remessa.Show
  Remessa.ZOrder 0
End Sub

Private Sub mnuUtiInfoBoleto_Click()
  InfoBoleto.Show
  InfoBoleto.ZOrder 0
End Sub

Private Sub mnuUtilCancelaGenerico_Click()
  CancelaGenerico.Show
  CancelaGenerico.ZOrder 0
End Sub

Private Sub mnuUtilReimpBol_Click()
  ReBoleto.Show
  ReBoleto.ZOrder 0
End Sub

Private Sub Timer1_Timer()
On Error GoTo Errado
  
  Timer1.Enabled = False
  
  With Informe
    .Panels(1).Text = "Empresa: " & Empresa & "   "
    .Panels(1).AutoSize = sbrContents
    .Panels(2).Text = "Usuário: ? "
    .Panels(2).AutoSize = sbrContents
    .Panels(3).Text = "Data: " & Format(Date, "dd/mm/yyyy") & "  "
    .Panels(3).AutoSize = sbrContents
    .Panels(4).Text = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision & "  Banco de dados: " & Parametros.dados & " Versão: " & sVerDb & "  "
    .Panels(4).AutoSize = sbrContents
  End With
  Acesso.Show vbModal
  
  VerificaBackup
  
Fim:
  Exit Sub
  
Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim

End Sub

