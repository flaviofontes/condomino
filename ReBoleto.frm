VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form ReBoleto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpimir boleta vencida"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   Icon            =   "ReBoleto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpVencimento 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Top             =   4680
      Width           =   1395
      _ExtentX        =   2461
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
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0000C000&
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CheckBox chCorrigir 
      Caption         =   "Corrigir valor do boleto"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   4740
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1035
      Top             =   3630
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ReBoleto.frx":000C
      Height          =   3945
      Left            =   45
      OleObjectBlob   =   "ReBoleto.frx":0020
      TabIndex        =   3
      Top             =   600
      Width           =   10005
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programas desenvolvidos\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1410
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   120
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
      Left            =   1140
      TabIndex        =   0
      Top             =   120
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Novo vencimento"
      Height          =   195
      Left            =   2220
      TabIndex        =   8
      Top             =   4740
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "ReBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbCondominio As Recordset

Private Sub Refaz()
  Dim nValor As Double
  Dim selCondominio As Recordset
  Dim rsBoleto As Recordset
  Dim sBarra  As String
  Dim nFator  As String
  Dim fValor  As String
  Dim vbDig   As String
  Dim sLivre As String
  Dim Campo1  As String
  Dim Campo2  As String
  Dim Campo3  As String
  Dim Campo4  As String
  Dim Campo5  As String
  Dim NossoNumero As String
  Dim sTemp As String
  Dim qDig As Integer
  Dim sSql As String
  Dim DadoAnterior As String
  
  If IsNumeric(DBGrid1.Columns(1).Text) Then
    Data1.Recordset.Bookmark = DBGrid1.Bookmark
    
    'recriar boleto atualizando data de vencimento e valor
    'trocar linha digitavél e código de barras
    
    With Data1.Recordset
      
      If chCorrigir.Value = 1 Then
        nValor = Reajustar(!cond, !valr, CDate(cpVencimento.Text), !vcto)
      Else
        nValor = !valr
      End If
      
      Resp = MsgBox("O boleto de '" & NomeCompleto(!cdsc) & "' será atualizado e impresso com vencimento " _
        & Format$(cpVencimento.Text, "dd/MM/yyyy") & " e valor de " & Format$(nValor, "#,##0.00") _
        & ". Continuar?", vbQuestion + vbYesNo, "Reimpressão")
      
      If Resp = vbNo Then
        Exit Sub
      End If
      
      fValor = Format(nValor, "#0.00")
      fValor = Left(fValor, InStr(fValor, ",") - 1) & Right(fValor, 2)
      fValor = PadLeft(fValor, "0", 10)
      nFator = CStr(CDate(cpVencimento.Text) - CDate("07/10/1997"))
      
      Set selCondominio = db.OpenRecordset("select * from condominio where codigo = " & !cond & ";", dbOpenDynaset)
      selCondominio.MoveFirst
      
      If selCondominio!tipoboleto = 1 Then
        
        NossoNumero = !nosso
        
        sTemp = selCondominio!Conta & ""
        qDig = DigitosCedente(1)
        If Len(sTemp) > qDig Then
          sTemp = Right(sTemp, qDig)
        ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
          sTemp = PadLeft(sTemp, "0", qDig)
        End If
        
        sLivre = sTemp & DigitoNosso(sTemp)
        sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
        sLivre = sLivre & Mid$(NossoNumero, 6, 3)
        sLivre = sLivre & Mid$(NossoNumero, 2, 1)
        sLivre = sLivre & Mid$(NossoNumero, 9)
            
        sLivre = sLivre & DigitoNosso(sLivre)
      
      ElseIf selCondominio!tipoboleto = 2 Then
      
        NossoNumero = !nosso
        sTemp = selCondominio!Conta & ""
        qDig = DigitosCedente(2)
        If Len(sTemp) > qDig Then
          sTemp = Right(sTemp, qDig)
        ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
          sTemp = PadLeft(sTemp, "0", qDig)
        End If
        
        sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & sTemp
      
      End If
      
      sBarra = Left(selCondominio!CdAgencia, 3) & "9" & nFator & fValor & sLivre
      
      vbDig = DigitoBarra(sBarra)
      
      sBarra = Left(sBarra, 4) & vbDig & Mid(sBarra, 5)
      Campo1 = Mid(sBarra, 1, 4) & Mid(sBarra, 20, 5)
      Campo1 = Campo1 & DigitoCodigo(Campo1)
      Campo1 = Left(Campo1, 5) & "." & Mid(Campo1, 6)
      
      Campo2 = Mid(sBarra, 25, 10)
      Campo2 = Campo2 & DigitoCodigo(Campo2)
      Campo2 = Left(Campo2, 5) & "." & Mid(Campo2, 6)
      
      Campo3 = Mid(sBarra, 35)
      Campo3 = Campo3 & DigitoCodigo(Campo3)
      Campo3 = Left(Campo3, 5) & "." & Mid(Campo3, 6)
      
      Campo4 = Mid(sBarra, 5, 1)
      
      Campo5 = Mid(sBarra, 6, 4) & Mid(sBarra, 10, 10)
      
      
      Set rsBoleto = db.OpenRecordset("select * from boletos where cdsc=" & !cdsc & " and bole = '" & !bole & "' and tran = '" & !tran & "';", dbOpenDynaset)
      If rsBoleto.RecordCount > 0 Then
        rsBoleto.MoveFirst
        DadoAnterior = rsBoleto!tran & " Vencida em " & Format(rsBoleto!vcto, "dd/MM/yyyy")
        rsBoleto.Edit
        rsBoleto!vcto = cpVencimento.Text
        rsBoleto!valr = nValor
        rsBoleto!corrigido = nValor
        rsBoleto!CDBARRA = sBarra
        rsBoleto!CODI = Campo1 & " " & Campo2 & " " & Campo3 & " " & Campo4 & " " & Campo5
        If InStr(rsBoleto!texto, "2a via de boleto.") = 0 Then
          rsBoleto!texto = rsBoleto!texto & vbCrLf & "2a via do boleto ref. a " & DadoAnterior
        End If
        If !idStatus = 2 Or !idStatus = 4 Or !idStatus = 6 Then
          rsBoleto!idStatus = 10
        ElseIf !idStatus = 8 Then
          rsBoleto!idStatus = 1
        End If
        rsBoleto.Update
      End If
      
      Set selCondominio = Nothing
      
      sSql = "Select * from boletos where cdsc = " & !cdsc & " and cond = " & !cond & " and nosso = '" & NossoNumero & "' order by nome;"
      
      If Left$(!tran, 2) = "AV" Then
        RelatoriosRPT.mnuSupExport.Visible = False
        RelatoriosRPT.mnuExportBoleto.Visible = True
        RelatoriosRPT.mnuEviarEmail.Visible = True
        RelatoriosRPT.Carregar "", Parametros.dados, "{boletos.cdsc} = " & Data1.Recordset!cdsc & " and {boletos.nosso} = '" _
            & Data1.Recordset!nosso & "'", "Boletos", sFormataCaminho(App.Path) & "avulso.rpt", , sSql
      Else
        Dim PerSind As Double
        
        PerSind = PersentualSindico(CodigoSindico(!cond))
        
        RelatoriosRPT.mnuSupExport.Visible = False
        RelatoriosRPT.mnuExportBoleto.Visible = True
        RelatoriosRPT.mnuEviarEmail.Visible = True
        RelatoriosRPT.Carregar "", Parametros.dados, "{boletos.tran} = '" & !tran & "' and {boletos.nosso} = '" & NossoNumero & "' and {boletos.cond} = " _
            & !cond, "Boletos", sFormataCaminho(App.Path) & "cobranca.rpt", , sSql, "persind|" & PerSind
      End If
      cmdSelecionar_Click
    End With
  End If
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
      End If
    End With
  End If
End Sub

Private Sub cmdPrint_Click()
  If DBGrid1.ApproxCount = 0 Then
    MsgBox "Nenhum boleto para reimpressão.", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  If Not IsDate(cpVencimento.Text) Then
    MsgBox "Nova data de vencimento não informada ou não é válida.", vbInformation + vbOKOnly, "Aviso"
    cpVencimento.SetFocus
    Exit Sub
  End If
    
  If CDate(cpVencimento.Text) < Date Then
    MsgBox "Nova data de vencimento não pode ser menor que a data atual.", vbInformation + vbOKOnly, "Aviso"
    cpVencimento.SetFocus
    Exit Sub
  End If
  
  Refaz
End Sub

Private Sub cmdSelecionar_Click()
  If Val(cpCodigo.Text) > 0 Then
    Selecionar cpCodigo.Text
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
          cmdSelecionar.SetFocus
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
        cmdSelecionar.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpVencimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
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
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
End Sub

Private Sub Selecionar(ByVal nCod As Long)
  Data1.RecordSource = "SELECT BOLETOS.*, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO+' '+ASSOCIADOS.PROPRIETARIO AS nomecomp, BLOCOS.NOME_BLOCO " _
    & "FROM BLOCOS RIGHT JOIN (ASSOCIADOS RIGHT JOIN BOLETOS ON ASSOCIADOS.CODIGO = BOLETOS.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO " _
    & "WHERE (((BOLETOS.COND)=" & nCod & ") AND ((BOLETOS.PAGO)='N') AND ((Boletos.CANCELADO)='N') AND ((BOLETOS.VCTO)<#" & Format$(Date, "MM/dd/yyyy") & "#)) " _
    & "ORDER BY ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO, BOLETOS.VCTO;"
  Data1.Refresh
End Sub

Public Function Reajustar(ByVal lCod As Long, _
                          ByVal oldValor As Double, _
                          ByVal NovoVenc As Date, _
                          ByVal oldVenc As Date) As Double
On Error GoTo Errado
  
  Dim dData     As Date
  Dim nDias     As Long
  Dim nJuros    As Single
  Dim valor     As String
  Dim Carencia  As Double
  Dim juros     As Double
  Dim Correcao  As Double
  Dim Multa     As Double
  Dim dbAnterior As Double
  
  Carencia = 30
  dData = NovoVenc
  dbAnterior = 0
  
  With tbCondominio
    .Index = "codigoid"
    .Seek "=", lCod
    If Not .NoMatch Then
      juros = !juros
      Multa = !Multa
      nJuros = juros / 30
    Else
      juros = Parametros.juros
      Multa = Parametros.Multa
      nJuros = juros / 30
    End If
  End With
  
  nDias = NovoVenc - oldVenc
  If nDias > 0 Then
    Correcao = Round(oldValor * (Multa / 100), 2)
    valor = Round(oldValor * (nJuros * nDias) / 100, 2)
    dbAnterior = oldValor + Correcao + valor
  Else
    dbAnterior = oldValor
  End If
  
Fim:
  Reajustar = dbAnterior
  Exit Function
  
Errado:
  MsgBox "Erro: " & Err.Number & vbCrLf & "Descrição: " & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim
  
End Function


Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
