VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RelatoriosRPT 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10020
   Icon            =   "RelatoriosRPT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstEnvioEmail 
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   4140
      Visible         =   0   'False
      Width           =   9135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9675
      lastProp        =   500
      _cx             =   17066
      _cy             =   6588
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Menu mnuImpressora 
      Caption         =   "Impressora"
      Begin VB.Menu mnuConfigPrint 
         Caption         =   "Configurar"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "Fechar"
      End
   End
   Begin VB.Menu mnuSupExport 
      Caption         =   "Exportar"
      Begin VB.Menu mnuExportPDF 
         Caption         =   "Adobe PDF"
      End
      Begin VB.Menu mnuExportExcel 
         Caption         =   "MicroSoft Excel"
      End
      Begin VB.Menu mnuExportWord 
         Caption         =   "MicroSoft Word"
      End
   End
   Begin VB.Menu mnuExportBoleto 
      Caption         =   "Exportar"
      Visible         =   0   'False
      Begin VB.Menu mnuExAdobe 
         Caption         =   "Adobe PDF"
         Begin VB.Menu mnuExAdobeEste 
            Caption         =   "Somente este"
         End
         Begin VB.Menu mnuExAdobeTodos 
            Caption         =   "Todos"
         End
      End
   End
   Begin VB.Menu mnuOrdenar 
      Caption         =   "Ordenar"
      Visible         =   0   'False
      Begin VB.Menu mnuOrdNome 
         Caption         =   "Nome"
      End
      Begin VB.Menu mnuOrdEndereco 
         Caption         =   "Endereço"
      End
      Begin VB.Menu mnuOrdBairro 
         Caption         =   "Bairro"
      End
      Begin VB.Menu mnuOrdeApartamento 
         Caption         =   "Unidade/Apartamento"
      End
   End
   Begin VB.Menu mnuEviarEmail 
      Caption         =   "Enviar por E-mail"
      Visible         =   0   'False
      Begin VB.Menu mnuEviarEamilEste 
         Caption         =   "Somente este boleto"
      End
      Begin VB.Menu mnuEnviarEmailTodos 
         Caption         =   "Todos os boletos"
      End
   End
End
Attribute VB_Name = "RelatoriosRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Cancela As Boolean

Dim WithEvents Report As CRAXDRT.Report
Attribute Report.VB_VarHelpID = -1
Dim CRApp2 As New CRAXDRT.Application
Dim sOrdem As String
Dim i As Integer
Dim sInforma As String
Dim rsSql As String
Dim noClique As Boolean
Dim rs As Recordset
Dim oTabela As DatabaseTable
Dim oField As DatabaseFieldDefinition

'para boletos
Dim localLcond As Long
Dim localStran As String

Public Sub Carregar(ByVal sCampos As String, _
                  ByVal sBancoDados As String, _
                  ByVal sFiltro As String, _
                  ByVal sTitulo As String, _
                  ByVal sReport As String, _
                  Optional ByVal sGroupFiltro As String = "", _
                  Optional ByVal SelBoleto As String = "", _
                  Optional ByVal params As String = "", _
                  Optional ByVal TabsLocal As String = "", _
                  Optional ByVal lCond As Long = -1, _
                  Optional ByVal sTran As String = "", _
                  Optional Parametros As String = "", _
                  Optional ByVal SubReports As String = "")
  
On Error GoTo Errado
  Dim i As Integer
  Dim sCP() As String
  Dim sSort() As CRcampos
  Dim rValor As Double
  Dim sPar As String
  Dim pNome As String
  Dim crParametro As ParameterFieldDefinition
  Dim sTabs() As String
  Dim arParams() As String
  Dim vlpar As ParameterValue
  Dim sTit() As String
  Dim crxSubreport As CRAXDRT.Report
  
'  Progre.Label1.Caption = "Carregando..."
'  Progre.Show
'  Call SetTopMostWindow(Me.hWnd, True)
  
  DoEvents
  
  Me.MousePointer = vbHourglass
  
  sTit = Split(sTitulo, vbCrLf)
  Me.Caption = sTit(0)
  Me.Show
  sInforma = Principal.Informe.Panels(1).Text
  Principal.Informe.Panels(1).Text = sTit(0)
  
  If SelBoleto <> "" Then
    rsSql = SelBoleto
    localLcond = lCond
    localStran = sTran
    Set rs = db.OpenRecordset(rsSql, dbOpenDynaset)
    If rs.RecordCount > 0 Then
      rs.MoveFirst
    End If
    mnuOrdNome.Checked = True
  End If
  
  '"{relentradas.numero};crAscendingOrder;relentradas|"
  If sCampos <> "" Then
    sCP = Split(sCampos, "|")
    For i = 0 To UBound(sCP)
      ReDim Preserve sSort(i)
      sSort(i).Campo = GetPiece(sCP(i), ";", 1)
      If GetPiece(sCP(i), ";", 2) = "crAscendingOrder" Then
        sSort(i).Orden = crAscendingOrder
      Else
        sSort(i).Orden = crDescendingOrder
      End If
      sSort(i).tabela = GetPiece(sCP(i), ";", 3)
    Next i
  End If
  
  Set Report = CRApp2.OpenReport(sReport)
  Report.EnableParameterPrompting = False
  Report.DiscardSavedData
  Report.ReportTitle = sTitulo
  Report.RecordSelectionFormula = sFiltro
  If sGroupFiltro <> "" Then
    Report.GroupSelectionFormula = sGroupFiltro
  End If
  
  If Parametros <> "" Then
    Report.EnableParameterPrompting = False
    Set crParametro = Report.ParameterFields.GetItemByName(GetPiece(Parametros, "|", 1))
    With crParametro
      .ClearCurrentValueAndRange
      .DiscreteOrRangeKind = crDiscreteValue
      .EnableMultipleValues = False
      .DisallowEditing = True
      .AddDefaultValue (CDbl(GetPiece(Parametros, "|", 2)))
      .SetNthDefaultValue 1, CDbl(GetPiece(Parametros, "|", 2))
    End With
  End If
  
  If params <> "" Then
    arParams = Split(params, ";")
    For i = 0 To UBound(arParams)
      If GetPiece(arParams(i), "|", 1) <> "" Then
        pNome = GetPiece(arParams(i), "|", 1)
        sPar = GetPiece(arParams(i), "|", 2)
        Report.FormulaFields.GetItemByName(pNome).Text = Replace(sPar, ",", ".")
      End If
    Next i
  End If
  
  If TabsLocal = "" Then
    For Each oTabela In Report.Database.Tables
      oTabela.Location = sBancoDados
    Next
  Else
    sTabs = Split(TabsLocal, ";")
    For i = 0 To UBound(sTabs)
      For Each oTabela In Report.Database.Tables
        If UCase(oTabela.Name) = sTabs(i) Then
          oTabela.Location = sFormataCaminho(App.Path) & "supersoft.mdb"
        Else
          oTabela.Location = sBancoDados
        End If
      Next
    Next i
  End If
  
  'teste
  If SubReports <> "" Then
    sTabs = Split(SubReports, ";")
    For i = 0 To UBound(sTabs)
      Set crxSubreport = Report.OpenSubreport(sTabs(i))
      For Each oTabela In crxSubreport.Database.Tables
        oTabela.Location = sBancoDados
      Next
    Next i
  End If
  
  If sCampos <> "" Then
    For i = 1 To Report.RecordSortFields.Count
      Report.RecordSortFields.Delete (1)
    Next i
  
    For i = 0 To UBound(sSort)
      For Each oTabela In Report.Database.Tables
        If oTabela.Name = sSort(i).tabela Then
          For Each oField In oTabela.Fields
            If oField.Name = sSort(i).Campo Then
              Report.RecordSortFields.Add oField, sSort(i).Orden
            End If
          Next
        End If
      Next
    Next i
  End If
  
  Set crxSubreport = Nothing
  Set crParametro = Nothing
  
  Report.VerifyOnEveryPrint = True
  Report.ReadRecords
  
  CRViewer91.ReportSource = Report
  CRViewer91.Zoom 100
  CRViewer91.ViewReport
  While CRViewer91.IsBusy
    DoEvents
  Wend
  
Sair:
  Call SetTopMostWindow(Me.hWnd, False)
'  Unload Progre
  Me.MousePointer = vbDefault
  Exit Sub

Errado:
  If Err.Number <> 445 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  End If
  Set CRApp2 = Nothing
  Set Report = Nothing
  Set rs = Nothing
  Principal.Informe.Panels(1).Text = sInforma
  Resume Sair

End Sub

Private Sub CRViewer91_FirstPageButtonClicked(UseDefault As Boolean)
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.MoveFirst
    End If
  End If
End Sub

Private Sub CRViewer91_LastPageButtonClicked(UseDefault As Boolean)
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.MoveLast
    End If
  End If
End Sub

Private Sub CRViewer91_NextPageButtonClicked(UseDefault As Boolean)
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.AbsolutePosition = CRViewer91.GetCurrentPageNumber - 1
'      rs.MoveNext
    End If
  End If
End Sub

Private Sub CRViewer91_PrevPageButtonClicked(UseDefault As Boolean)
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.AbsolutePosition = CRViewer91.GetCurrentPageNumber - 1
'      rs.MovePrevious
    End If
  End If
End Sub

Private Sub CRViewer91_PrintButtonClicked(UseDefault As Boolean)
  If Not rs Is Nothing Then
    With rs
      .AbsolutePosition = CRViewer91.GetCurrentPageNumber - 1
      If rs!vcto < Date Then
        Principal.balao.ShowBalloonTip "Para reimprimir este boleto, você deve ir em reimpressão de boletos vencidos.", beError, "Boletos", 10000
        MsgBox "Este boleto está vencido. Não pode ser impresso!", vbExclamation + vbOKOnly, "Aviso"
        UseDefault = False
      End If
    End With
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    Cancela = False
  End If
End Sub

Private Sub Form_Load()
  noClique = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set CRApp2 = Nothing
  Set Report = Nothing
  Set rs = Nothing
  Principal.Informe.Panels(1).Text = sInforma
End Sub

Private Sub Form_Resize()
  CRViewer91.Top = 0
  CRViewer91.Left = 0
  CRViewer91.Height = ScaleHeight
  CRViewer91.Width = ScaleWidth
  lstEnvioEmail.Top = ScaleHeight - lstEnvioEmail.Height
  lstEnvioEmail.Left = 0
  lstEnvioEmail.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set CRApp2 = Nothing
  Set Report = Nothing
  Set rs = Nothing
  Set oTabela = Nothing
  Set oField = Nothing
  Principal.Informe.Panels(1).Text = sInforma
End Sub

Private Sub lstEnvioEmail_DblClick()
  lstEnvioEmail.Visible = False
  Refresh
End Sub

Private Sub mnuConfigPrint_Click()
  Report.PrinterSetup (Me.hWnd)
  CRViewer91.Refresh
End Sub

Private Sub mnuEnviarEmailTodos_Click()
On Error GoTo Erro

  Dim sFile As String
  Dim i As Integer
  Dim sNome As String
  Dim sEmail As String
  Dim tPags As Long
  Dim nEnviados As Integer
  Dim NomeCond As String
  Dim quantasVezes As Integer
  
  Cancela = False
  
  Dim rep As String
  rep = "Internet"
  
  If Not IsWebConnected(rep) Then
    MsgBox "Você não está conectado a internet.", vbInformation + vbOKOnly, "Internet Status"
    Exit Sub
  End If
  
  Progre.Label1.Caption = "Enviando e-mail... Tecle ESC para cancelar."
  Progre.Show
  Call SetTopMostWindow(Progre.hWnd, True)
  Progre.Refresh
    
5 If mnuOrdNome.Checked = False Then
    mnuOrdNome_Click
  End If
  
  While CRViewer91.IsBusy
    DoEvents
  Wend
  
  quantasVezes = 0
  UltimoErro = ""
  rs.MoveFirst
  
  Call CRViewer91.GetLastPageNumber(tPags, True)
  
  If (rs.RecordCount <> tPags) Then
    Resp = MsgBox("Houve um erro na tentativa de enviar os boleto por email. Tentar novamente?" & tPags, vbCritical + vbYesNo, "Aviso")
    If Resp = vbYes Then
      GoTo 5
    Else
      Exit Sub
    End If
  End If
    
  nEnviados = 0
  
  For i = 1 To tPags
    If Cancela = True Then
      GoTo 10
    End If
    lstEnvioEmail.Visible = True
    Me.Refresh
    sNome = AcertaLetras(NomeCompleto(rs!cdsc))
    sEmail = EmailAssociado(rs!cdsc)
    NomeCond = NomeCondominio(rs!cond)
    Progre.Label1.Caption = "Enviando e-mail... Tecle ESC para cancelar." & vbCrLf & sNome
    Progre.Refresh
    
    If Trim(sEmail) <> "" Then
      Progre.Label1.Caption = "Enviando e-mail... Tecle ESC para cancelar." & vbCrLf & sNome & vbCrLf & sEmail
      Progre.Refresh
      
      sFile = sFormataCaminho(GetTempDir) & sNome & ".pdf"
      
      If Left(rs!tran, 2) = "AV" Then
        NomeCond = NomeCond & vbCrLf & "do mês " & Format$(rs!Data, "MM/yyyy") & " (" & rs!DIGITAVAL & ")"
      Else
        NomeCond = NomeCond & vbCrLf & "do mês " & rs!tran & " (" & rs!DIGITAVAL & ")"
      End If
        
      If Cancela = True Then
        GoTo 10
      End If
      
      If sysFiles.FileExists(sFile) Then
        Call sysFiles.DeleteFile(sFile, True)
      End If
    
      Report.DisplayProgressDialog = False
      Report.EnableParameterPrompting = False
      Report.ExportOptions.DestinationType = crEDTDiskFile
      Report.ExportOptions.DiskFileName = sFile
      Report.ExportOptions.FormatType = crEFTPortableDocFormat
      Report.ExportOptions.PDFExportAllPages = False
      Report.ExportOptions.PDFFirstPageNumber = i
      Report.ExportOptions.PDFLastPageNumber = i
      Report.Export False
    
      If Cancela = True Then
        GoTo 10
      End If
      
50    If EnvioDeEmail(sEmail, sFile, NomeCond, rs!cdsc) Then
        acrescentaItem sNome & " - " & sEmail, rs!id
        nEnviados = nEnviados + 1
        quantasVezes = 0
        DoEvents
      Else
        DoEvents
        If quantasVezes < 3 Then
          quantasVezes = quantasVezes + 1
          GoTo 50
        End If
        acrescentaItem UltimoErro & " - " & sNome & " - " & sEmail, rs!id
      End If
      
      If sysFiles.FileExists(sFile) Then
        Call sysFiles.DeleteFile(sFile, True)
      End If
    End If
    DoEvents
    rs.MoveNext
  Next i
  
10  Call SetTopMostWindow(Progre.hWnd, False)
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  DoEvents
  MsgBox nEnviados & " boleto(s) enviado(s) com sucesso!", vbInformation + vbOKOnly, "Aviso"
  
Fim:
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  
  tPags = CRViewer91.GetCurrentPageNumber
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      For i = 2 To tPags
        rs.MoveNext
      Next i
    End If
  End If
  
  Report.DisplayProgressDialog = True
  Exit Sub

Erro:
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  If Err.Number <> 32755 And Err.Number <> 0 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim

End Sub

Private Sub acrescentaItem(ByVal sLinha As String, ByVal sId As Long)
  lstEnvioEmail.AddItem (sLinha)
  lstEnvioEmail.ItemData(lstEnvioEmail.NewIndex) = sId
  lstEnvioEmail.ListIndex = lstEnvioEmail.NewIndex
End Sub

Private Sub mnuEviarEamilEste_Click()
On Error GoTo Erro

  Dim sFile As String
  Dim sNome As String
  Dim sEmail As String
  Dim NomeCond As String
  Dim pgAtual As Integer
  Dim i As Integer
  
  Dim rep As String
  rep = "Internet"
  
  If Not IsWebConnected(rep) Then
    MsgBox "Você não está conectado a internet.", vbInformation + vbOKOnly, "Internet Status"
    Exit Sub
  End If
  
  pgAtual = CRViewer91.GetCurrentPageNumber
  
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      For i = 2 To pgAtual
        rs.MoveNext
      Next i
    End If
  End If

  sNome = AcertaLetras(NomeCompleto(rs!cdsc))
  sEmail = EmailAssociado(rs!cdsc)
  NomeCond = NomeCondominio(rs!cond)
  
  If sEmail = "" Then
    MsgBox sNome & " não possui email cadastrado!", vbInformation + vbOKOnly, "Aviso"
    GoTo Fim
  End If
  
  Resp = MsgBox("Enviar este boleto para " & sNome & " no email " & sEmail & "?", vbQuestion + vbYesNo, "Enviar")
  If Resp = vbNo Then
    GoTo Fim
  End If
  
  Progre.Label1.Caption = "Enviando e-mail..." & vbCrLf & sNome & vbCrLf & sEmail
  Progre.Show
  
  Call SetTopMostWindow(Progre.hWnd, True)
  Progre.Refresh
  
  If Left(rs!tran, 2) = "AV" Then
    NomeCond = NomeCond & vbCrLf & "do mês " & Format$(rs!Data, "MM/yyyy") & " (" & rs!DIGITAVAL & ")"
  Else
    NomeCond = NomeCond & vbCrLf & "do mês " & rs!tran & " (" & rs!DIGITAVAL & ")"
  End If
  
  sFile = sFormataCaminho(GetTempDir) & sNome & ".pdf"
  
  If sysFiles.FileExists(sFile) Then
    Call sysFiles.DeleteFile(sFile, True)
  End If

  Report.DisplayProgressDialog = False
  Report.EnableParameterPrompting = False
  Report.ExportOptions.DestinationType = crEDTDiskFile
  Report.ExportOptions.DiskFileName = sFile
  Report.ExportOptions.FormatType = crEFTPortableDocFormat
  Report.ExportOptions.PDFExportAllPages = False
  Report.ExportOptions.PDFFirstPageNumber = CRViewer91.GetCurrentPageNumber
  Report.ExportOptions.PDFLastPageNumber = CRViewer91.GetCurrentPageNumber
  Report.Export False

  If EnvioDeEmail(sEmail, sFile, NomeCond, rs!cdsc) Then
    Call SetTopMostWindow(Progre.hWnd, False)
    Unload Progre
    MsgBox "Boleto enviado para '" & sEmail & "' com sucesso!", vbExclamation + vbOKOnly, "E-mail"
  Else
    Call SetTopMostWindow(Progre.hWnd, False)
    Unload Progre
    MsgBox UltimoErro, vbExclamation + vbOKOnly, "E-mail"
  End If
  
  If sysFiles.FileExists(sFile) Then
    Call sysFiles.DeleteFile(sFile, True)
  End If

Fim:
  Report.DisplayProgressDialog = True
  Exit Sub

Erro:
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  If Err.Number <> 32755 And Err.Number <> 0 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim
End Sub

Private Sub mnuExAdobeEste_Click()
On Error GoTo Erro

  Dim sFile As String
  Dim i As Integer
  Dim sNome As String
  Dim pgAtual As Integer
  
  pgAtual = CRViewer91.GetCurrentPageNumber
  
  If Not rs Is Nothing Then
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      For i = 2 To pgAtual
        rs.MoveNext
      Next i
    End If
  End If
  
  sNome = AcertaLetras(NomeCompleto(rs!cdsc))
  
  sFile = sFormataCaminho(PastaSistema(5)) & sNome & ".pdf"
  
  With CommonDialog1
    .CancelError = True
    .DialogTitle = "Exportação Adobe PDF"
    .filename = sFile
    .Filter = "Adobe PDF|*.pdf"
    .InitDir = sFormataCaminho(PastaSistema(5))
    .DefaultExt = "pdf"
    .ShowSave
    sFile = .filename
  End With

  If sysFiles.FileExists(sFile) Then
    Call sysFiles.DeleteFile(sFile, True)
  End If

  Report.EnableParameterPrompting = False
  Report.ExportOptions.DestinationType = crEDTDiskFile
  Report.ExportOptions.DiskFileName = sFile
  Report.ExportOptions.FormatType = crEFTPortableDocFormat
  Report.ExportOptions.PDFExportAllPages = False
  Report.ExportOptions.PDFFirstPageNumber = CRViewer91.GetCurrentPageNumber
  Report.ExportOptions.PDFLastPageNumber = CRViewer91.GetCurrentPageNumber
  Report.Export False

  MsgBox "Arquivo '" & sFile & "' gerado com sucesso.", vbExclamation + vbOKOnly, "Exportar"

Fim:
  Exit Sub

Erro:
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim
End Sub

Private Sub mnuExAdobeTodos_Click()
On Error GoTo Erro

  Dim sFile As String
  Dim sFolder As String
  Dim i As Long
  Dim nPaginas As Long
  Dim pdfFolder As Folder
  Dim pdfFiles As File
  Dim sNome As String
  
  Progre.Label1.Caption = "Exportando..."
  Progre.Show
  Call SetTopMostWindow(Progre.hWnd, True)
  
  If mnuOrdNome.Checked = False Then
    mnuOrdNome_Click
  End If
  
  While CRViewer91.IsBusy
    DoEvents
  Wend
  
  sFolder = sFormataCaminho(PastaSistema(5)) & "boletos\"

  If Not sysFiles.FolderExists(sFolder) Then
    sysFiles.CreateFolder (sFolder)
  End If
  
  Set pdfFolder = sysFiles.GetFolder(sFolder)
  For Each pdfFiles In pdfFolder.Files
    pdfFiles.Delete (True)
  Next
  
  Report.EnableParameterPrompting = False
  Report.ExportOptions.DestinationType = crEDTDiskFile
  Report.ExportOptions.FormatType = crEFTPortableDocFormat
  Report.ExportOptions.PDFExportAllPages = False
  Call CRViewer91.GetLastPageNumber(nPaginas, True)
  Report.DisplayProgressDialog = False
  rs.MoveFirst
  For i = 1 To nPaginas
    sNome = AcertaLetras(NomeCompleto(rs!cdsc))
    Progre.Label1.Caption = "Exportando Página " & PadLeft(i, "0", 6) & vbCrLf & sNome
    Progre.Label1.Refresh
    sFile = sFormataCaminho(sFolder) & sNome & ".pdf"
    Report.ExportOptions.DiskFileName = sFile
    Report.ExportOptions.PDFFirstPageNumber = i
    Report.ExportOptions.PDFLastPageNumber = i
    Report.Export False
    DoEvents
    rs.MoveNext
  Next i
  Report.DisplayProgressDialog = True
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  DoEvents
  MsgBox "Os boletos foram exportados com sucesso para pasta '" & sFolder & "'.", vbExclamation + vbOKOnly, "Exportar"

Fim:
  Report.DisplayProgressDialog = True
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  Set pdfFolder = Nothing
  Set pdfFiles = Nothing
  Exit Sub

Erro:
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim
End Sub

Private Sub mnuExportExcel_Click()
On Error GoTo Erro

  Dim sFile As String

  sFile = sFormataCaminho(PastaSistema(5)) & "rel_" & Format$(Now, "ddMMMyyyHHMM") & ".xls"
  
  With CommonDialog1
    .CancelError = True
    .DialogTitle = "Exportação Excel"
    .filename = sFile
    .Filter = "Microsoft Excel|.xls"
    .InitDir = sFormataCaminho(PastaSistema(5))
    .DefaultExt = "xls"
    .ShowSave
    sFile = .filename
  End With

  If sysFiles.FileExists(sFile) Then
    Call sysFiles.DeleteFile(sFile, True)
  End If

  Report.ExportOptions.DestinationType = crEDTDiskFile
  Report.ExportOptions.DiskFileName = sFile
  Report.ExportOptions.FormatType = crEFTExcel97
  Report.ExportOptions.ExcelExportAllPages = True
  Report.Export False

  MsgBox "Arquivo '" & sFile & "' gerado com sucesso.", vbExclamation + vbOKOnly, "Exportar"

Fim:
  Exit Sub

Erro:
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim

End Sub

Private Sub mnuExportPDF_Click()
On Error GoTo Erro

  Dim sFile As String

  sFile = sFormataCaminho(PastaSistema(5)) & "rel_" & Format$(Now, "ddMMMyyyHHMM") & ".pdf"
  
  With CommonDialog1
    .CancelError = True
    .DialogTitle = "Exportação Adobe PDF"
    .filename = sFile
    .Filter = "Adobe PDF|*.pdf"
    .InitDir = sFormataCaminho(PastaSistema(5)) '& "pdf"
    .DefaultExt = "pdf"
    .ShowSave
    sFile = .filename
  End With

  If sysFiles.FileExists(sFile) Then
    Call sysFiles.DeleteFile(sFile, True)
  End If

  Report.ExportOptions.DestinationType = crEDTDiskFile
  Report.ExportOptions.DiskFileName = sFile
  Report.ExportOptions.FormatType = crEFTPortableDocFormat
  Report.ExportOptions.PDFExportAllPages = True
  Report.Export False

  MsgBox "Arquivo '" & sFile & "' gerado com sucesso.", vbExclamation + vbOKOnly, "Exportar"

Fim:
  Exit Sub

Erro:
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim
End Sub

Private Sub mnuExportWord_Click()
On Error GoTo Erro

  Dim sFile As String

  sFile = sFormataCaminho(PastaSistema(5)) & "rel_" & Format$(Now, "ddMMMyyyHHMM") & ".doc"
  
  With CommonDialog1
    .CancelError = True
    .DialogTitle = "Exportação Word"
    .filename = sFile
    .Filter = "Microsoft Word|*.doc"
    .InitDir = sFormataCaminho(PastaSistema(5))
    .DefaultExt = "doc"
    .ShowSave
    sFile = .filename
  End With

  If sysFiles.FileExists(sFile) Then
    Call sysFiles.DeleteFile(sFile, True)
  End If

  Report.ExportOptions.DestinationType = crEDTDiskFile
  Report.ExportOptions.DiskFileName = sFile
  Report.ExportOptions.FormatType = crEFTWordForWindows
  Report.ExportOptions.WORDWExportAllPages = True
  Report.Export False

  MsgBox "Arquivo '" & sFile & "' gerado com sucesso.", vbExclamation + vbOKOnly, "Exportar"

Fim:
  Exit Sub

Erro:
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
    If sysFiles.FileExists(sFile) Then
      Call sysFiles.DeleteFile(sFile, True)
    End If
  End If
  Resume Fim

End Sub

Private Sub mnuFechar_Click()
  Unload Me
End Sub

Private Sub mnuOrdBairro_Click()
  Me.MousePointer = 11
  Set rs = db.OpenRecordset(Left(rsSql, InStrRev(rsSql, "by") + 2) & " bair;", dbOpenDynaset)
  For i = 1 To Report.RecordSortFields.Count
    Report.RecordSortFields.Delete (1)
  Next i
  For Each oTabela In Report.Database.Tables
    If oTabela.Name = "BOLETOS" Then
      For Each oField In oTabela.Fields
        If oField.Name = "{BOLETOS.BAIR}" Then
          Report.RecordSortFields.Add oField, crAscendingOrder
        End If
      Next
    End If
  Next
  CRViewer91.Refresh
  While CRViewer91.IsBusy
    DoEvents
  Wend
  CRViewer91.ShowFirstPage
  rs.MoveFirst
  mnuOrdNome.Checked = False
  mnuOrdEndereco.Checked = False
  mnuOrdeApartamento.Checked = False
  mnuOrdBairro.Checked = True
  Me.MousePointer = 0
End Sub

Private Sub mnuOrdeApartamento_Click()
  Me.MousePointer = 11
  Set rs = db.OpenRecordset("SELECT BOLETOS.*,  BLOCOS.NOME_BLOCO+' '+ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO+' '+BOLETOS.NOME AS ORDENAR " _
      & "FROM BLOCOS INNER JOIN (ASSOCIADOS INNER JOIN BOLETOS ON ASSOCIADOS.CODIGO = BOLETOS.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO " _
      & "WHERE (((BOLETOS.COND)=" & localLcond & ") AND ((BOLETOS.TRAN)= '" & localStran & "')) ORDER BY  BLOCOS.NOME_BLOCO+' '+ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO+' '+BOLETOS.NOME;", dbOpenDynaset)
  Report.RecordSortFields(1).Field = Report.FormulaFields.GetItemByName("nomeinquilino")
  Report.RecordSortFields(1).SortDirection = crAscendingOrder
  CRViewer91.Refresh
  While CRViewer91.IsBusy
    DoEvents
  Wend
  CRViewer91.ShowFirstPage
  rs.MoveFirst
  mnuOrdNome.Checked = False
  mnuOrdEndereco.Checked = False
  mnuOrdeApartamento.Checked = True
  mnuOrdBairro.Checked = False
  Me.MousePointer = 0
End Sub

Private Sub mnuOrdEndereco_Click()
  Me.MousePointer = 11
  Set rs = db.OpenRecordset(Left(rsSql, InStrRev(rsSql, "by") + 2) & " ende;", dbOpenDynaset)
  For i = 1 To Report.RecordSortFields.Count
    Report.RecordSortFields.Delete (1)
  Next i
  For Each oTabela In Report.Database.Tables
    If oTabela.Name = "BOLETOS" Then
      For Each oField In oTabela.Fields
        If oField.Name = "{BOLETOS.ENDE}" Then
          Report.RecordSortFields.Add oField, crAscendingOrder
        End If
      Next
    End If
  Next
  CRViewer91.Refresh
  While CRViewer91.IsBusy
    DoEvents
  Wend
  CRViewer91.ShowFirstPage
  rs.MoveFirst
  mnuOrdNome.Checked = False
  mnuOrdEndereco.Checked = True
  mnuOrdeApartamento.Checked = False
  mnuOrdBairro.Checked = False
  Me.MousePointer = 0
End Sub

Private Sub mnuOrdNome_Click()
  Me.MousePointer = 11
  Set rs = db.OpenRecordset(rsSql, dbOpenDynaset)
  For i = 1 To Report.RecordSortFields.Count
    Report.RecordSortFields.Delete (1)
  Next i
  For Each oTabela In Report.Database.Tables
    If oTabela.Name = "BOLETOS" Then
      For Each oField In oTabela.Fields
        If oField.Name = "{BOLETOS.BOLE}" Then
          Report.RecordSortFields.Add oField, crAscendingOrder
        End If
      Next
    End If
  Next
  CRViewer91.Refresh
  While CRViewer91.IsBusy
    DoEvents
  Wend
  CRViewer91.ShowFirstPage
  rs.MoveFirst
  mnuOrdNome.Checked = True
  mnuOrdEndereco.Checked = False
  mnuOrdeApartamento.Checked = False
  mnuOrdBairro.Checked = False
  Me.MousePointer = 0
End Sub

'Private Sub Report_NoData(pCancel As Boolean)
'  If noClique Then
'    MsgBox "Nunhum dado para imprimir.", vbInformation + vbOKOnly, "Aviso"
'    noClique = False
'    Unload Me
'  End If
'End Sub
