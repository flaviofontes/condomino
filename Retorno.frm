VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RetornoCEF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analise de arquivos de retorno"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   HelpContextID   =   9660
   Icon            =   "Retorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstTipo 
      Height          =   1185
      Left            =   720
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   5400
      Width           =   6375
   End
   Begin VB.ComboBox cpFiltro 
      Height          =   315
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   8655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Retorno.frx":000C
      Height          =   3135
      Left            =   60
      OleObjectBlob   =   "Retorno.frx":0020
      TabIndex        =   7
      Top             =   2160
      Width           =   9315
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Programas desenvolvidos\Vicosa\Reca\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "retorno"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdBaixa 
      Caption         =   "&Processar este retorno"
      Height          =   375
      Left            =   7260
      TabIndex        =   6
      Top             =   5820
      Width           =   2115
   End
   Begin VB.CommandButton cmdRelatorio 
      Caption         =   "&Relatorio"
      Height          =   375
      Left            =   7260
      TabIndex        =   5
      Top             =   5400
      Width           =   2115
   End
   Begin VB.CommandButton cmdProcessar 
      Caption         =   "&Carregar"
      Height          =   375
      Left            =   7860
      TabIndex        =   3
      Top             =   300
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog AbrirArq 
      Left            =   8220
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbrir 
      Height          =   375
      Left            =   7140
      Picture         =   "Retorno.frx":0F13
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   615
   End
   Begin VB.TextBox cpArquivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   7035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Imprimir"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   5400
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Filtrar"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label lbRegistros 
      Caption         =   "Número de registros: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   9105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo de retorno"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1305
   End
End
Attribute VB_Name = "RetornoCEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UltimoDir As String
Dim jaPro As Boolean
Dim uRetorno As String
Dim NumeroSeq As Double
Dim sysFiles As New FileSystemObject
Dim rsCliente As Recordset
Dim rsBaixados As Recordset
Dim TipoRel As Integer
Dim dtArquivo As String
Dim HoraArquivo As String
Dim sSequencial As String
Dim sCond As String
Dim Tamanho As Integer


Public KeyRemove As Long


Private Sub cmdAbrir_Click()
  db.Execute "Delete From Retorno;"
  Data1.Refresh
  DBGrid1.ReBind
  cpFiltro.Clear
  lstTipo.Clear
  DoEvents
  AbrirRetorno.Show vbModal
End Sub


Private Sub ProcessaTodos()
  Dim i As Integer
  
  With AbrirRetorno.listaRetornos
    For i = 1 To .ListItems.Count - 1
      RetornoCaixa .ListItems(i).SubItems(3), 1
    Next i
    RetornoCaixa .ListItems(.ListItems.Count).SubItems(3), 0
  End With
  
End Sub

Private Sub cmdBaixa_Click()

On Error GoTo errHandler
  Dim sValor As String
  Dim sCred As String
  Dim sNome As String
  Dim nJuros As Double
  Dim nDias As Integer
  Dim i As Integer
  Dim sArquivo As String
  Dim rs As Recordset
  Dim sNao As String
  Dim rsRetorno As Recordset
  Dim sBoletosNao As String
  
  cmdAbrir.Enabled = False
  cmdProcessar.Enabled = False
  cmdRelatorio.Enabled = False
  cmdBaixa.Enabled = False
  
  Set rsRetorno = db.OpenRecordset("select * from retorno order by titulo;", dbOpenDynaset)
  
  Resp = MsgBox("Todos os títulos listados aqui serão quitados. Prosseguir?", vbQuestion + vbYesNo, "Baixa")
  If (Resp = vbYes) Then
    Me.Refresh
    sBoletosNao = ""
    Set rsBaixados = db.OpenRecordset("baixados", dbOpenTable)
    With rsRetorno
      If .RecordCount > 0 Then
        .MoveFirst
        sNao = ""
        DBEngine.BeginTrans
        Do While Not .EOF
          If (!movimento = "06") Then
            Set rs = db.OpenRecordset("select * from boletos where id = " & !id_boleto & ";", dbOpenDynaset)
            If rs.RecordCount > 0 Then
              rs.MoveFirst
              rs.Edit
              rs!tarifas = !tarifas
              rs!vlpago = !valorpgto
              rs!pago = "S"
              rs!dtpgto = !dataocorre
              rs!idStatus = 7
              If rs!acumulado = "S" Then
                db.Execute "update boletos set " & _
                  "pago = 'S'" & _
                  "idstatus = 7" & _
                  ", dtpgto = #" & Format$(.Fields("dataocorre").Value, "mm/dd/yyyy") & _
                  "#, tarifas = 0, vlpago = 0 " & _
                  " where cdsc = " & !cliente & " and left(tran,2) <> 'AV' and pago = 'N' and vcto < #" _
                  & Format$(rs!vcto, "MM/dd/yyyy") & "#;"
              End If
              rs.Update
              Pagamento !id_boleto, !dataocorre, !valorpgto, !Descricao, !tarifas
            Else
              sBoletosNao = sBoletosNao & .Fields("titulo").Value & " " & PadRight(.Fields("nome").Value, " ", 40) & " " & PadLeft(Format$(.Fields("valorpgto").Value, "#,##0.00"), " ", 15) & vbCrLf
            End If
          ElseIf (!movimento = "02") Then
            db.Execute "update boletos set idstatus = 2 where id = " & !id_boleto & ";"
          ElseIf (!movimento = "09") Then
            db.Execute "delete from boletos where idstatus = 9 and cancelado = 'S' and id = " & !id_boleto & ";"
          End If
          sNome = Mid$(cpArquivo.Text, InStrRev(cpArquivo.Text, "\") + 1)
          
          If Not RetornoJaBaixado(sNome, dtArquivo, HoraArquivo, sSequencial) Then
            rsBaixados.AddNew
            rsBaixados!Titulo = !Titulo
            rsBaixados!Documento = !Documento
            rsBaixados!Vencimento = !Vencimento
            rsBaixados!valor = !valor
            rsBaixados!NUMEROTITULO = !NUMEROTITULO
            rsBaixados!COTA = ""
            If !Nome & "" = "" Then
              rsBaixados!Nome = GetNome(Mid$(!Titulo, 11))
              rsBaixados!COTA = ""
            Else
              rsBaixados!Nome = !Nome
            End If
            rsBaixados!tarifas = !tarifas
            rsBaixados!motivo = !motivo
            rsBaixados!juros = !juros
            rsBaixados!desconto = !desconto
            rsBaixados!ABATIMENTO = !ABATIMENTO
            rsBaixados!iof = !iof
            rsBaixados!valorpgto = !valorpgto
            rsBaixados!VALORCREDITO = !VALORCREDITO
            rsBaixados!dataocorre = !dataocorre
            rsBaixados!datacred = !datacred
            rsBaixados!Data = !Data
            rsBaixados!VALOROC = !VALOROC
            rsBaixados!cliente = !cliente
            rsBaixados!boleto = !Titulo
            rsBaixados!pago = !valorpgto
            rsBaixados!Codigo = !cliente
            rsBaixados!DESPESA = 0 '!DESPESA
            rsBaixados!Referencia = "" '!Referencia
            rsBaixados!idConta = 1
            rsBaixados.Update
          End If
          DoEvents
          .MoveNext
        Loop
        If (sBoletosNao <> "") Then
          frmMensagem.lbMensagem.Caption = "Os títulos abaixo não foram quitados favor verificar." & vbCrLf & vbCrLf & sBoletosNao
          frmMensagem.Show 1
        End If
        sNome = Mid$(cpArquivo.Text, InStrRev(cpArquivo.Text, "\") + 1)
        If Not RetornoJaBaixado(sNome, dtArquivo, HoraArquivo, sSequencial) Then
            db.Execute "insert into retbaixados (arquivo, data, hora, sequencial) values('" & _
                sNome & "', #" & dtArquivo & "#, '" & HoraArquivo & "', '" & sSequencial & "');"
        End If
        DBEngine.CommitTrans
        MsgBox "Retorno processado com sucesso!", vbExclamation + vbOKOnly, "Quitação"
        Dim sDestino As String
        If KeyRemove = -5 Then
          With AbrirRetorno.listaRetornos
            For i = 1 To .ListItems.Count
              sNome = "RET" & Format$(Now, "ddmmyyyy hhmmss") & ".ret"
              sDestino = sFormataCaminho(App.Path) & "retorno\" & Trim(AcertaLetras(sCond)) & "\"
              If Not (sysFiles.FolderExists(sDestino)) Then
                Call sysFiles.CreateFolder(sDestino)
              End If
              Call sysFiles.CopyFile(.ListItems(i).SubItems(3), sFormataCaminho(sDestino) & sNome)
              Call sysFiles.DeleteFile(.ListItems(i).SubItems(3), True)
            Next i
          End With
        Else
          sNome = "RET" & sSequencial & "-" & Format$(Now, "ddmmyyyy hhmmss") & ".ret"
          sDestino = sFormataCaminho(App.Path) & "retorno\" & Trim(AcertaLetras(sCond)) & "\"
          If Not (sysFiles.FolderExists(sDestino)) Then
            Call sysFiles.CreateFolder(sDestino)
          End If
          If sysFiles.FileExists(cpArquivo.Text) Then
            Call sysFiles.CopyFile(cpArquivo.Text, sFormataCaminho(sDestino) & sNome)
            Call sysFiles.DeleteFile(cpArquivo.Text, True)
          End If
        End If
      End If
    End With
    Set rsBaixados = Nothing
    db.Execute "Delete from retorno;"
    Data1.Refresh
    cpArquivo.Text = ""
    If KeyRemove = -5 Then
      AbrirRetorno.listaRetornos.ListItems.Clear
    Else
      If KeyRemove <= AbrirRetorno.listaRetornos.ListItems.Count Then
        AbrirRetorno.listaRetornos.ListItems.Remove KeyRemove
      End If
    End If
  End If
Sair:
  cmdAbrir.Enabled = True
  cmdProcessar.Enabled = True
  cmdRelatorio.Enabled = True
  cmdBaixa.Enabled = True
  Set rsRetorno = Nothing
  Exit Sub
 
errHandler:
  MsgBox trataErros(Err), vbCritical + vbOKOnly, "Erro"
  DBEngine.Rollback
  Resume Sair
End Sub

Private Sub cmdProcessar_Click()

On Error GoTo errHandler
  Dim sPathArq As String
  cmdAbrir.Enabled = False
  cmdProcessar.Enabled = False
  cmdRelatorio.Enabled = False
  cmdBaixa.Enabled = False
  db.Execute "Delete from retorno;"
  Data1.Refresh
  lstTipo.Clear
  cpFiltro.Clear
  cpFiltro.AddItem "Todos"
  cpFiltro.ListIndex = 0
  
  If cpArquivo.Text <> Empty Then
    If Dir$(cpArquivo.Text) <> Empty Then
      sPathArq = Left$(cpArquivo.Text, InStrRev(cpArquivo.Text, "\") - 1)
      If RetornoCaixa Then
        Data1.Refresh
        DBGrid1.ReBind
        uRetorno = Format$(Val(uRetorno) + 1, "0000000")
        If Not sysFiles.FolderExists(App.Path & "\retorno") Then
          MkDir App.Path & "\retorno"
        End If
      End If
    End If
  End If

Sair:
  cmdAbrir.Enabled = True
  cmdProcessar.Enabled = True
  cmdRelatorio.Enabled = True
  cmdBaixa.Enabled = True
  Exit Sub
 
errHandler:
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Mensagem: " & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Sair

End Sub

Private Sub cmdRelatorio_Click()
  Dim i As Integer
  Dim sFormula As String
  
  sFormula = ""
  For i = 0 To lstTipo.ListCount - 1
    If lstTipo.Selected(i) = True Then
      sFormula = sFormula & "{retorno.descricao} = '" & lstTipo.List(i) & "' or "
    End If
  Next i
  If Right(sFormula, 4) = " or " Then
    sFormula = Left(sFormula, Len(sFormula) - 4)
  End If
  If cpFiltro.Text <> "Todos" Then
    sFormula = "left({retorno.descricao}, " & Len(cpFiltro.Text) & ") = '" & cpFiltro.Text & "'"
  End If
  RelatoriosRPT.Carregar "", Parametros.dados, sFormula, "Retorno: " & cpArquivo.Text, sFormataCaminho(App.Path) & "retorno.rpt"
End Sub

Private Sub cpFiltro_Click()
  If cpFiltro.Text = "Todos" Then
    Data1.RecordSource = "select * from retorno order by titulo;"
    Data1.Refresh
  Else
    Tamanho = Len(cpFiltro.Text)
    Data1.RecordSource = "select * from retorno where left(descricao," & Tamanho & ") = '" & cpFiltro.Text & "' order by titulo;"
    Data1.Refresh
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
  db.Execute "Delete From Retorno;"
  Refresh
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
  Set rsCliente = db.OpenRecordset("associados", dbOpenTable)
  rsCliente.Index = "codigoid"
End Sub

Private Function RetornoCaixa(Optional sNomeFile As String = "", Optional Todos As Integer = 0) As Boolean

On Error GoTo errHandler
  Dim sDetalhe As String
  Dim nFree As Integer
  Dim sFile As String
  Dim sCampo As String
  Dim nArq As Double
  Dim nReg As Double
  Dim nData As Date
  Dim sTipos As Recordset
  Dim sCliente As rNome
  Dim i As Integer
  Dim sNome As String
  Dim txtDados() As String
  Dim nRemessa As String
  Dim sErros() As String
  Dim Count As Integer
  Dim idCondominio As Long
  Dim linha As String
  Dim sDescricao As String
  Dim sPasta As String
  Dim Parcial As Boolean
  
  Parcial = False
  
  If sNomeFile = "" Then
    sFile = cpArquivo.Text
  Else
    sFile = sNomeFile
  End If
  
  nArq = 0
  RetornoCaixa = True
  dtArquivo = ""
  
  nFree = FreeFile
 
  Open sFile For Input As #nFree
    txtDados = Split(Input(LOF(nFree), nFree), vbCrLf)
  Close #nFree
  
  sDetalhe = txtDados(0)
  If Mid$(sDetalhe, 1, 3) <> "104" Then
    If sNomeFile = "" Then
      MsgBox "Este retorno não é da Caixa Econômica Federal.", vbCritical + vbOKOnly, "Aviso"
    End If
    Exit Function
  End If
  
  sCampo = Mid$(sDetalhe, 144, 8)
  sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
  dtArquivo = sCampo
  sCampo = Mid$(sDetalhe, 152, 6)
  sCampo = Left$(sCampo, 2) & ":" & Mid$(sCampo, 3, 2) & ":" & Right$(sCampo, 2)
  HoraArquivo = sCampo
  sCampo = Mid$(sDetalhe, 158, 6)
  sSequencial = sCampo
  nData = CDate(dtArquivo)
    
  Dim nRegs As Integer
  nRegs = 0
  
  For i = 0 To UBound(txtDados)
    If Mid$(txtDados(i), 8, 1) = "3" And Mid$(txtDados(i), 14, 1) = "T" Then
      nRegs = nRegs + 1
    End If
  Next i
  
  sPasta = Mid$(txtDados(0), 59, 6)
  sCond = GetCondominioConvenio(sPasta, txtDados(0))
  idCondominio = GetCodigoCondominioConvenio(Mid$(sDetalhe, 59, 6))
    
  sDetalhe = txtDados(UBound(txtDados) - 1)
  sCampo = Mid$(sDetalhe, 24, 6)
  lbRegistros.Caption = "Condomínio: " & sPasta & " " & sCond & vbCrLf & "Arquivo. Data: " & Format$(nData, "dd/MM/yyyy") _
                        & ", Hora: " & HoraArquivo & ", Sequencial: " & sSequencial & vbCrLf & "Número de boletos: " & nRegs
  lbRegistros.Refresh
  nReg = Val(sCampo)
  nArq = Val(sSequencial)
  i = 1
  DoEvents
  If sNomeFile = "" Then
    sNome = Mid$(sFile, InStrRev(cpArquivo.Text, "\") + 1)
    If RetornoJaBaixado(sNome, dtArquivo, HoraArquivo, sSequencial) Then
      Resp = MsgBox("Retorno já processado. Processar novamente?", vbQuestion + vbYesNo, "Retorno")
      If Resp = vbNo Then
        Exit Function
      End If
    End If
  End If
  
  db.Execute "Delete from retorno;"
  DoEvents
  Data1.Refresh
  DBGrid1.ReBind
  DoEvents
  
  If nReg <= 0 Then
    Exit Function
  End If
  
  sPasta = Mid$(txtDados(0), 59, 6)
  sCampo = Mid$(txtDados(0), 143, 1)
  Select Case sCampo
    Case "2"
      GoTo Processa
    Case "3"
      MsgBox Mid$(txtDados(0), 172, 20), vbInformation + vbOKOnly, "Aviso"
      sNome = Mid$(cpArquivo.Text, InStrRev(cpArquivo.Text, "\") + 1)
      nRemessa = Mid$(txtDados(0), 158, 6)
      processarRemessa nRemessa, sPasta
      salvaRetorno sPasta
      Exit Function
    Case "4"
      MsgBox Mid$(txtDados(0), 172, 20) & "ARCIAL. Favor verificar erros informados.", vbInformation + vbOKOnly, "Aviso"
      Parcial = True
      GoTo ProcessaRejeitada
    Case "5"
      MsgBox Mid$(txtDados(0), 172, 20), vbInformation + vbOKOnly, "Aviso"
      GoTo ProcessaRejeitada
    Case "6"
      MsgBox Mid$(txtDados(0), 172, 20), vbInformation + vbOKOnly, "Aviso"
      salvaRetorno sPasta
      Exit Function
  End Select
  
  
Processa:
  SetTopMostWindow Progresso.hWnd, True
  Progresso.Barra.Max = nReg
  Progresso.Caption = "Lendo arquivo..."
  Progresso.Show
  Me.SetFocus
  For i = 0 To UBound(txtDados)
    Progresso.Barra.Value = i
    sDetalhe = txtDados(i)
    'Analisa linha
    sCampo = Mid$(sDetalhe, 8, 1)
    
    If sCampo = "1" Then
      sCampo = Mid$(sDetalhe, 9, 1)
      nRemessa = Mid$(sDetalhe, 186, 6)
    ElseIf sCampo = "3" Then
      sCampo = Mid$(sDetalhe, 14, 1)
    End If
    
    If Mid$(sDetalhe, 8, 1) = "3" Then
      sCampo = Mid$(sDetalhe, 14, 1)
      If sCampo = "T" Then
        Data1.Recordset.AddNew
        sCampo = Mid$(sDetalhe, 16, 2)
        Data1.Recordset.Fields("Movimento").Value = sCampo
        sCampo = Mid$(sDetalhe, 214, 10)
        Data1.Recordset.Fields("Motocorre").Value = sCampo
        sCampo = Mid$(sDetalhe, 18, 6)
        Data1.Recordset.Fields("AgConta").Value = sCampo
        sCampo = Mid$(sDetalhe, 24, 13)
        Data1.Recordset.Fields("Conta").Value = sCampo
        sCampo = Mid$(sDetalhe, 40, 17)
        Data1.Recordset.Fields("Titulo").Value = Trim(sCampo)
        sCampo = Mid$(sDetalhe, 40, 2)
        Data1.Recordset.Fields("CARTEIRA").Value = Val(sCampo)
        sCampo = Mid$(sDetalhe, 59, 15)
        Data1.Recordset.Fields("Documento").Value = sCampo
        sCampo = Mid$(sDetalhe, 74, 8)
        If IsNumeric(sCampo) Then
          sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
          If IsDate(sCampo) Then
            Data1.Recordset.Fields("Vencimento").Value = CDate(sCampo)
          End If
        End If
        sCampo = Mid$(sDetalhe, 82, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("Valor").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("Valor").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 106, 25)
        Data1.Recordset.Fields("NumeroTitulo").Value = Trim(sCampo)
        sCampo = Mid$(sDetalhe, 131, 2)
        Data1.Recordset.Fields("MOEDA").Value = sCampo
        sCampo = Mid$(sDetalhe, 134, 15)
        Data1.Recordset.Fields("Inscricao").Value = sCampo
        sCampo = Mid$(sDetalhe, 149, 40)
        Data1.Recordset.Fields("Nome").Value = sCampo
        sCampo = Mid$(sDetalhe, 189, 10)
        Data1.Recordset.Fields("Contrato").Value = sCampo
        sCampo = Mid$(sDetalhe, 199, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("Tarifas").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("Tarifas").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 214, 10)
        Data1.Recordset.Fields("motivo").Value = sCampo
        sDescricao = RetornaDescricaoMovimento(Data1.Recordset.Fields("movimento").Value)
        incluirFiltros sDescricao
        If InStr("02,03,26,30", Data1.Recordset.Fields("movimento").Value) > 0 Then
          sDescricao = sDescricao & " " & RetornaDescricaoErro(Mid$(sDetalhe, 214, 2))
        ElseIf Data1.Recordset.Fields("movimento").Value = "28" Then
          sDescricao = sDescricao & " " & RetornaDescricaoC047B(Mid$(sDetalhe, 214, 2))
        ElseIf InStr("06,09,17", Data1.Recordset.Fields("movimento").Value) > 0 Then
          sDescricao = sDescricao & " " & RetornaDescricaoC047C(Left(Data1.Recordset.Fields("motocorre").Value, 2))
          sDescricao = sDescricao & " " & RetornaDescricaoC047D(Mid(Data1.Recordset.Fields("motocorre").Value, 3, 2))
        End If
        Data1.Recordset.Fields("Descricao").Value = sDescricao
        incluirTipos Data1.Recordset.Fields("descricao").Value
      ElseIf sCampo = "U" Then
        sCampo = Mid$(sDetalhe, 18, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("Juros").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("Juros").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 33, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("Desconto").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("Desconto").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 48, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("Abatimento").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("Abatimento").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 63, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("IOF").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("IOF").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 78, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("valorpgto").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("valorpgto").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 93, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("ValorCredito").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("ValorCredito").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 138, 8)
        If IsNumeric(sCampo) Then
          sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
          If IsDate(sCampo) Then
            Data1.Recordset.Fields("DATAOCORRE").Value = CDate(sCampo)
          End If
        End If
        sCampo = Mid$(sDetalhe, 146, 8)
        If IsNumeric(sCampo) Then
          sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
          If IsDate(sCampo) Then
            Data1.Recordset.Fields("DataCred").Value = CDate(sCampo)
          End If
        End If
        sCampo = Mid$(sDetalhe, 154, 4)
        Data1.Recordset.Fields("Ocorrencia").Value = sCampo
        sCampo = Mid$(sDetalhe, 158, 8)
        If IsNumeric(sCampo) Then
          sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
          If IsDate(sCampo) Then
            Data1.Recordset.Fields("Data").Value = CDate(sCampo)
          End If
        End If
        sCampo = Mid$(sDetalhe, 166, 15)
        If IsNumeric(sCampo) Then
          Data1.Recordset.Fields("ValorOc").Value = CDbl(sCampo) / 100
        Else
          Data1.Recordset.Fields("ValorOc").Value = 0
        End If
        sCampo = Mid$(sDetalhe, 181, 30)
        Data1.Recordset.Fields("ComplOcorre").Value = sCampo
        '====== nome do cliente  =========
        sCliente = AchaNome(Trim(Data1.Recordset.Fields("Titulo").Value), Data1.Recordset.Fields("valor").Value)
        Data1.Recordset.Fields("Nome").Value = sCliente.Nome
        Data1.Recordset.Fields("Cliente").Value = sCliente.Codigo
        Data1.Recordset.Fields("ID_BOLETO").Value = sCliente.id_boleto
        '=================================
        Data1.Recordset.Update
      ElseIf sCampo = "W" Then
        sCampo = Mid$(sDetalhe, 18, 6)
        If (Val(sCampo) = 1) Then
          Unload Progresso
          MsgBox Trim(Mid$(sDetalhe, 34, 50)), vbInformation + vbOKOnly, "Aviso"
          salvaRetorno sPasta
        Else
          If Val(nRemessa) <= 0 Then
            nRemessa = retornaUltimaRemessa(idCondominio)
          End If
          sCampo = Mid$(sDetalhe, 18, 6)
          linha = Val(sCampo)
          sCampo = Mid$(sDetalhe, 28, 1)
          Data1.Recordset.AddNew
          If sCampo = "P" Or sCampo = "Q" Then
            sCampo = RetornaLinhaRemessa(linha, nRemessa, sPasta, "3", sCampo)
            Data1.Recordset.Fields("Titulo").Value = GetPiece(sCampo, "|", 1)
            Data1.Recordset.Fields("valor").Value = GetPiece(sCampo, "|", 2)
          Else
            Data1.Recordset.Fields("Titulo").Value = "0"
            Data1.Recordset.Fields("valor").Value = 0
          End If
          sCampo = Mid$(sDetalhe, 16, 2)
          sCampo = RetornaDescricaoMovimento(sCampo)
          incluirFiltros sCampo
          sCampo = Trim(Mid$(sDetalhe, 25, 133))
          If (Len(sCampo) > 0) Then
            sErros = Split(sCampo, " ")
            sCampo = ""
            For Count = 0 To UBound(sErros)
              sCampo = sCampo & Trim(RetornaDescricaoErro(Right(sErros(Count), 2))) & " "
            Next Count
          End If
          Data1.Recordset.Fields("Descricao").Value = Trim(sCampo)
          incluirTipos Data1.Recordset.Fields("descricao").Value
          '====== nome do cliente  =========
          sCliente = AchaNome(Trim(Data1.Recordset.Fields("Titulo").Value), Data1.Recordset.Fields("valor").Value)
          Data1.Recordset.Fields("Nome").Value = sCliente.Nome
          Data1.Recordset.Fields("Cliente").Value = sCliente.Codigo
          Data1.Recordset.Fields("ID_BOLETO").Value = sCliente.id_boleto
          '=================================
          Data1.Recordset.Update
        End If
      End If
    End If
  Next i
  DoEvents
  
  NumeroSeq = nArq
  TipoRel = 1
  GoTo Sair
  
ProcessaRejeitada:

  sDetalhe = txtDados(0)
  
  sCampo = Mid$(sDetalhe, 144, 8)
  sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
  dtArquivo = sCampo
  sCampo = Mid$(sDetalhe, 152, 6)
  sCampo = Left$(sCampo, 2) & ":" & Mid$(sCampo, 3, 2) & ":" & Right$(sCampo, 2)
  HoraArquivo = sCampo
  sCampo = Mid$(sDetalhe, 158, 6)
  sSequencial = sCampo
  nData = CDate(dtArquivo)
    
  sDetalhe = txtDados(UBound(txtDados) - 1)
  sCampo = Mid$(sDetalhe, 24, 6)
  nReg = (CDbl(sCampo) - 2) / 2
  nArq = Val(sSequencial)
  i = 1
  DoEvents
  db.Execute "Delete from retorno;"
  DoEvents
  Data1.Refresh
  DBGrid1.ReBind
  DoEvents
  If nReg <= 0 Then
    Exit Function
  End If
  SetTopMostWindow Progresso.hWnd, True
  Progresso.Barra.Max = (nReg * 2) + 2
  Progresso.Caption = "Lendo arquivo..."
  Progresso.Show
  Me.SetFocus
  For i = 0 To UBound(txtDados)
    If i <= ((nReg * 2) + 2) Then Progresso.Barra.Value = i
    sDetalhe = txtDados(i)
    'Analisa linha
    
    sCampo = Mid$(sDetalhe, 8, 1)
    If sCampo = "1" Then
      sCampo = Mid$(sDetalhe, 9, 1)
    ElseIf sCampo = "3" Then
      sCampo = Mid$(sDetalhe, 14, 1)
    End If
    
    If sCampo = "T" Then
      nRemessa = Mid$(sDetalhe, 186, 6)
    End If
    
    If sCampo = "W" Then
      If Val(nRemessa) <= 0 Then
        nRemessa = retornaUltimaRemessa(idCondominio)
      End If
      sCampo = Mid$(sDetalhe, 18, 6)
      linha = Val(sCampo)
      sCampo = Mid$(sDetalhe, 28, 1)
      Data1.Recordset.AddNew
      If sCampo = "P" Or sCampo = "Q" Then
        sCampo = RetornaLinhaRemessa(linha, nRemessa, sPasta, "3", sCampo)
        Data1.Recordset.Fields("Titulo").Value = GetPiece(sCampo, "|", 1)
        Data1.Recordset.Fields("valor").Value = Val(GetPiece(sCampo, "|", 2))
      Else
        Data1.Recordset.Fields("Titulo").Value = "0"
        Data1.Recordset.Fields("valor").Value = 0
      End If
      sCampo = Trim(Mid$(sDetalhe, 25, 133))
      If (Len(sCampo) > 0) Then
        sErros = Split(sCampo, " ")
        sCampo = ""
        incluirFiltros Trim(RetornaDescricaoErro(Right(sErros(0), 2)))
        For Count = 0 To UBound(sErros)
          sCampo = sCampo & Trim(RetornaDescricaoErro(Right(sErros(Count), 2))) & " "
        Next Count
      End If
      Data1.Recordset.Fields("Descricao").Value = Left(Trim(sCampo), 100)
      incluirTipos Data1.Recordset.Fields("descricao").Value
      '====== nome do cliente  =========
      sCliente = AchaNome(Trim(Data1.Recordset.Fields("Titulo").Value), Data1.Recordset.Fields("valor").Value)
      Data1.Recordset.Fields("Nome").Value = sCliente.Nome
      Data1.Recordset.Fields("Cliente").Value = sCliente.Codigo
      Data1.Recordset.Fields("Vencimento").Value = sCliente.Vencimento
      Data1.Recordset.Fields("ID_BOLETO").Value = sCliente.id_boleto
      '=================================
      Data1.Recordset.Update
    End If
  Next i
  DoEvents
  Close #nFree
  Unload Progresso
  DoEvents
  NumeroSeq = nArq
  TipoRel = 1
  If Parcial Then
    processarRemessa nRemessa, sPasta, True
  End If
  salvaRetorno sPasta
  If KeyRemove < AbrirRetorno.listaRetornos.ListItems.Count Then
'    AbrirRetorno.listaRetornos.ListItems.Remove KeyRemove
  End If

  cmdBaixa.Enabled = False


Sair:
  Unload Progresso
  Exit Function
 
errHandler:
  Unload Progresso
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Mensagem: " & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Function

Private Function AchaNome(ByVal sTitulo As String, Optional nVal As Double) As rNome

On Error GoTo errHandler
  Dim sRet As rNome
  Dim sVal As String
  Dim PegaTitulo As Recordset
  Dim sNomeCond As String
  Dim sInfo As String
  Dim sAchado As String
  
  
'  If Val(Mid(sTitulo, 3, 4)) = 0 Then
'    sVal = Replace(Format$(nVal, "#0.00"), ",", ".")
'    Set PegaTitulo = db.OpenRecordset("Select * From boletos Where valr = " & sVal & " and nosso = '" & sTitulo & "' order by vcto;", dbOpenDynaset)
'  Else
    Set PegaTitulo = db.OpenRecordset("Select * From boletos Where nosso = '" & sTitulo & "' order by vcto;", dbOpenDynaset)
''  End If
  
  With PegaTitulo
    If .RecordCount > 0 Then
      .MoveLast
      sAchado = !nosso
      sRet.Nome = NomeCompleto(.Fields("cdsc").Value) & " " & GetCondominioCodigo(!cond)
      sRet.Codigo = .Fields("cdsc").Value
      sRet.Vencimento = .Fields("vcto").Value
      sRet.Documento = ""  'Left$(.Fields("").Value, 50)
      sRet.id_boleto = !id
    Else
      sNomeCond = ""
      If IsNumeric(sTitulo) Then
        If Len(Trim(sTitulo)) > 7 Then
          sNomeCond = GetCondominioCodigo(Mid(sTitulo, 3, 4))
        End If
      End If
      sAchado = "NAO"
      sRet.Nome = "NÃO ENCONTRADO - " & sNomeCond
      sRet.Codigo = 0
      sRet.Vencimento = Date
      sRet.Documento = ""
      sRet.id_boleto = -1
    End If
  End With
  sInfo = sTitulo & "|" & nVal & "|" & sRet.Nome & "|" & sRet.Codigo & "|" & sRet.Vencimento & "|" & sAchado
  ArquivoInfo sInfo
  AchaNome = sRet

Sair:
  Exit Function
 
errHandler:
 
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Mensagem: " & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Function

Private Sub Form_Unload(Cancel As Integer)
'  db.Execute "Delete From Retorno;"
  Unload AbrirRetorno
End Sub

Private Function GetNome(ByVal nCod As Long) As String
  Dim rs As Recordset
  Dim sRet As String
  Set rs = db.OpenRecordset("select * from associados where codigo = " & nCod & ";", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      If !boleto = 1 Then
        sRet = !Nome & "  " & GetCondominioCodigo(!Condominio)
      Else
        If !Proprietario = "" Then
          sRet = !Nome & "  " & GetCondominioCodigo(!Condominio)
        Else
          sRet = !Proprietario & "  " & GetCondominioCodigo(!Condominio)
        End If
      End If
    Else
      sRet = ""
    End If
  End With
  rs.Close
  Set rs = Nothing
  GetNome = sRet
End Function

Public Function GetCondominioConvenio(ByVal sConvenio As String, sLinha As String) As String
  Dim rs As Recordset
  Dim sRet As String
  Set rs = db.OpenRecordset("select * from condominio where right(conta,6) = '" & sConvenio & "';", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      If !titularboleto = 1 Then
        sRet = !Nome
      Else
        sRet = !razaoboleto
      End If
    Else
      If Trim(sLinha) <> "" Then
        sRet = Mid$(sLinha, 73, 30) & "***"
      Else
        sRet = "CONVÊNIO " & sConvenio & " NÃO ENCONTRADO"
      End If
    End If
  End With
  rs.Close
  Set rs = Nothing
  GetCondominioConvenio = sRet
End Function

Private Function GetCondominioCodigo(ByVal lCod As Long) As String
  Dim rs As Recordset
  Dim sRet As String
  Set rs = db.OpenRecordset("select * from condominio where codigo = " & lCod & ";", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      sRet = !Nome
    Else
      sRet = ""
    End If
  End With
  rs.Close
  Set rs = Nothing
  GetCondominioCodigo = sRet
End Function

Public Function GetCodigoCondominioConvenio(ByVal sConvenio As String) As Long
  Dim rs As Recordset
  Dim sRet As Long
  Set rs = db.OpenRecordset("select * from condominio where conta = '" & sConvenio & "';", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      sRet = !Codigo
    Else
      sRet = -1
    End If
  End With
  rs.Close
  Set rs = Nothing
  GetCodigoCondominioConvenio = sRet
End Function

Private Function RetornaLinhaRemessa(ByVal nLinha As Integer, _
                                     ByVal sUltima As String, _
                                     ByVal sConvenio As String, _
                                     ByVal Registro As String, _
                                     ByVal segmento As String) As String
  Dim sNomeRemessa As String
  Dim txtDados() As String
  Dim nFree As Integer
  Dim i As Integer
  Dim sRet As String

  sNomeRemessa = sFormataCaminho(App.Path) & sConvenio & "\RE" & PadLeft(sUltima, "0", 6) & ".TXT"
  
  If Not sysFiles.FileExists(sNomeRemessa) Then
    If InStr(cpArquivo.Text, "retorno") > 0 Then
      sNomeRemessa = Left(cpArquivo.Text, InStr(cpArquivo.Text, "retorno") - 1)
      sNomeRemessa = sNomeRemessa & "seguranc\RE" & PadLeft(sUltima, "0", 6) & ".TXT"
    End If
  End If
  
  sRet = ""
  If sysFiles.FileExists(sNomeRemessa) Then
    nFree = FreeFile
    Open sNomeRemessa For Input As #nFree
      txtDados = Split(Input(LOF(nFree), nFree), vbCrLf)
    Close #nFree
    For i = 0 To UBound(txtDados)
      If (Mid$(txtDados(i), 8, 1) = "3") Then
        If (Val(Mid$(txtDados(i), 9, 5)) = nLinha) Then
          If Mid$(txtDados(i), 14, 1) = "Q" Then
            sRet = Mid$(txtDados(i - 1), 41, 17) & "|" & CStr((Val(Mid$(txtDados(i - 1), 86, 15)) / 100))
          Else
            sRet = Mid$(txtDados(i), 41, 17) & "|" & CStr((Val(Mid$(txtDados(i), 86, 15)) / 100))
          End If
        End If
      End If
    Next i
  End If
  RetornaLinhaRemessa = sRet
End Function


Private Sub salvaRetorno(ByVal sConvenio As String)
  Dim sDestino As String
  sDestino = sFormataCaminho(App.Path) & "retorno\" & sConvenio & "\"
  If Not (sysFiles.FolderExists(sDestino)) Then
    Call sysFiles.CreateFolder(sDestino)
  End If
  Call sysFiles.CopyFile(cpArquivo.Text, sFormataCaminho(sDestino))
  Call sysFiles.DeleteFile(cpArquivo.Text, True)
  If KeyRemove <= AbrirRetorno.listaRetornos.ListItems.Count Then
    AbrirRetorno.listaRetornos.ListItems.Remove KeyRemove
    cmdBaixa.Enabled = False
  End If
End Sub

Private Sub processarRemessa(ByVal numeroRemessa As String, ByVal sConvenio As String, Optional ByVal olharData1 As Boolean = False)
  Dim sNomeRemessa As String
  Dim txtDados() As String
  Dim nFree As Integer
  Dim i As Integer
  Dim NossoNumero As String

  sNomeRemessa = sFormataCaminho(App.Path) & sConvenio & "\RE" & PadLeft(numeroRemessa, "0", 6) & ".TXT"
  
  If Not sysFiles.FileExists(sNomeRemessa) Then
    If (InStr(cpArquivo.Text, "retorno") > 0) Then
      sNomeRemessa = Left(cpArquivo.Text, InStr(cpArquivo.Text, "retorno") - 1)
    Else
      sNomeRemessa = Left(cpArquivo.Text, InStrRev(cpArquivo.Text, "\"))
    End If
    sNomeRemessa = sNomeRemessa & "seguranc\RE" & PadLeft(numeroRemessa, "0", 6) & ".TXT"
  End If
  
  If sysFiles.FileExists(sNomeRemessa) Then
    nFree = FreeFile
    Open sNomeRemessa For Input As #nFree
      txtDados = Split(Input(LOF(nFree), nFree), vbCrLf)
    Close #nFree
    For i = 0 To UBound(txtDados)
      If (Mid$(txtDados(i), 14, 1) = "P") Then
        NossoNumero = Mid$(txtDados(i), 41, 17)
        If olharData1 Then
          Data1.Recordset.FindFirst "titulo = '" & NossoNumero & "'"
          If Data1.Recordset.NoMatch Then
            db.Execute "update boletos set idstatus = 2 where bole = '" & NossoNumero & "';"
          End If
        Else
          db.Execute "update boletos set idstatus = 2 where bole = '" & NossoNumero & "';"
        End If
      End If
    Next i
  End If

End Sub

Private Sub incluirTipos(ByVal sTipo As String)
  Dim h As Integer
  Dim tem As Boolean
  tem = False
  For h = 0 To lstTipo.ListCount - 1
    If (lstTipo.List(h) = sTipo) Then
      tem = True
      Exit For
    End If
  Next h
  If Not tem Then
    lstTipo.AddItem (sTipo)
  End If
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

