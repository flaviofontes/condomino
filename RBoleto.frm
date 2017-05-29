VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form RBoleto 
   Caption         =   "Relatório"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "RBoleto.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8.758
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   10.372
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Escolha 
      Left            =   4395
      Top             =   555
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar Horizontal 
      Height          =   270
      Left            =   210
      Max             =   22
      TabIndex        =   4
      Top             =   4665
      Width           =   5130
   End
   Begin VB.VScrollBar Vertical 
      Height          =   3420
      Left            =   5220
      Max             =   30
      TabIndex        =   3
      Top             =   420
      Width           =   285
   End
   Begin VB.PictureBox Rel 
      BackColor       =   &H00C0E0FF&
      Height          =   3480
      Left            =   0
      ScaleHeight     =   6.033
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   7.435
      TabIndex        =   1
      Top             =   420
      Width           =   4275
      Begin VB.PictureBox Pagina 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   11906
         Left            =   390
         ScaleHeight     =   20.955
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   29.66
         TabIndex        =   2
         Top             =   300
         Width           =   16838
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3105
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RBoleto.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RBoleto.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RBoleto.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RBoleto.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RBoleto.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RBoleto.frx":099C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Atalhos 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primeira página"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Página anterior."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proxima página."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Última página."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fechar."
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "RBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim Atual As Integer
Dim Max As Integer
Dim ChamaLpt As Integer
Dim pData1 As Date
Dim pData2 As Date
Dim nRel As String
Dim rsSelecao As DAO.Recordset
Const MaxView As Integer = 25
Dim sMesprn As String
Dim fechando As Boolean


Private Sub Atalhos_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1
      Impressao
    Case 3
      Atual = 0
      If rsSelecao.RecordCount > 0 Then
        rsSelecao.AbsolutePosition = Atual
        Pagina.Cls
        BoletoComp
      End If
    Case 4
      Pagina.Cls
      Atual = Atual - 1
      If Atual < 0 Then
        Atual = 0
        If rsSelecao.RecordCount > 0 Then
          rsSelecao.AbsolutePosition = 0
          BoletoComp
        End If
      Else
        If rsSelecao.RecordCount > 0 Then
          rsSelecao.AbsolutePosition = Int(Atual)
          BoletoComp
        End If
      End If
    Case 5
      Pagina.Cls
      Atual = Atual + 1
      If Atual > Max Then
        Atual = Max
        If rsSelecao.RecordCount > 0 Then
          rsSelecao.AbsolutePosition = Int(Atual)
          BoletoComp
        End If
      Else
        If rsSelecao.RecordCount > 0 Then
          rsSelecao.AbsolutePosition = Int(Atual)
          BoletoComp
        End If
      End If
    Case 6
      Pagina.Cls
      Atual = Max
      If rsSelecao.RecordCount > 0 Then
        rsSelecao.AbsolutePosition = Int(Atual)
        BoletoComp
      End If
    Case 8
      fechando = True
      Unload Me
  End Select
  AnunciaPagina Atual + 1
End Sub

Private Sub Form_Load()
  Refresh
  KeyPreview = True
  fechando = False
End Sub

Public Sub Iniciar()
  Pagina.ScaleMode = vbCentimeters
  Pagina.AutoRedraw = True
  Pagina.Height = 29.7
  Pagina.Width = 21
  Pagina.Left = Rel.Left + 0.5
  Pagina.Top = Rel.Top + 0.5
End Sub

Public Function NovaPagina() As Integer
  Pagina.Cls
End Function

Private Sub AnunciaPagina(ByVal nPage As Integer)
  Caption = nRel & " - página " & nPage & " de " & (Max + 1)
End Sub

'=============================================================
'Imprime o texto a direita
'=============================================================
Public Sub PrintDireita(ByVal Y As Single, ByVal X As Single, ByVal sTexto As String, ByVal sFonte As String, ByVal sSize As Integer, Optional ByVal Bold As Boolean, Optional ByVal Italico As Boolean)
On Error Resume Next
  Pagina.FontName = sFonte
  Pagina.FontSize = sSize
  Pagina.FontBold = Bold
  Pagina.FontItalic = Italico
  Pagina.CurrentX = X
  Pagina.CurrentY = Y
  Pagina.Print sTexto
End Sub

'=============================================================
'Imprime o texto a esquerda
'=============================================================
Public Sub PrintEsquerda(ByVal Y As Single, ByVal X As Single, ByVal sTexto As String, ByVal sFonte As String, ByVal sSize As Integer, Optional ByVal Bold As Boolean, Optional Italico As Boolean)
On Error Resume Next
  Pagina.FontName = sFonte
  Pagina.FontSize = sSize
  Pagina.FontBold = Bold
  Pagina.FontItalic = Italico
  Pagina.CurrentX = X - Pagina.TextWidth(sTexto)
  Pagina.CurrentY = Y
  Pagina.Print sTexto
End Sub

'=============================================================
'Imprime o texto centralizado
'=============================================================
Public Sub PrintCentralizado(ByVal Y As Single, ByVal X As Single, ByVal sTexto As String, ByVal sFonte As String, ByVal sSize As Integer, Optional ByVal Bold As Boolean, Optional Italico As Boolean)
On Error Resume Next
  Pagina.FontName = sFonte
  Pagina.FontSize = sSize
  Pagina.FontBold = Bold
  Pagina.FontItalic = Italico
  Pagina.CurrentX = X - (Pagina.TextWidth(sTexto) / 2)
  Pagina.CurrentY = Y
  Pagina.Print sTexto
End Sub

'=============================================================
'Imprime uma linha
'=============================================================
Public Sub PrintLinha(ByVal Y1 As Single, ByVal X1 As Single, ByVal Y2 As Single, ByVal X2 As Single, Optional vbDraw As DrawStyleConstants)
On Error Resume Next
  If Not IsMissing(vbDraw) Then
    Pagina.DrawStyle = vbDraw
  Else
    Pagina.DrawStyle = vbSolid
  End If
  Pagina.Line (X1, Y1)-(X2, Y2), 0, BF
  Pagina.DrawStyle = vbSolid
End Sub

'=============================================================
'Imprime um quadro
'=============================================================
Public Sub PrintQuadro(ByVal Y1 As Single, ByVal X1 As Single, ByVal Y2 As Single, ByVal X2 As Single, vbDraw As DrawStyleConstants)
On Error Resume Next
  If Not IsMissing(vbDraw) Then
    Pagina.DrawStyle = vbDraw
  Else
    Pagina.DrawStyle = vbSolid
  End If
  Pagina.Line (X1, Y1)-(X2, Y2), 0, B
  Pagina.DrawStyle = vbSolid
End Sub

'=============================================================
'Imprime desenho
'=============================================================
Public Sub PrintDesenho(ByVal Y1 As Single, ByVal X1 As Single, ByVal Y2 As Single, ByVal X2 As Single, ByVal sFigura As String)
On Error Resume Next
  If Dir$(sFigura) <> "" Then
    Pagina.PaintPicture VB.LoadPicture(sFigura), X1, Y1, Y2, X2
  End If
End Sub

Private Sub Form_Resize()
  If Not fechando Then
    Horizontal.Left = 0
    Horizontal.Top = Me.ScaleHeight - Horizontal.Height
    Horizontal.Width = Me.ScaleWidth - Vertical.Width
    Vertical.Top = 0.741
    Vertical.Height = Me.ScaleHeight - 0.741
    Vertical.Left = Me.ScaleWidth - Vertical.Width
    Rel.Move 0, 0.741, Me.ScaleWidth - Vertical.Width, Me.ScaleHeight - Horizontal.Height
  End If
End Sub

Private Sub Horizontal_Change()
  Pagina.Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
End Sub

Private Sub Vertical_Change()
  Pagina.Top = Rel.Top + 0.5 + (Vertical.Value * -1)
End Sub

Private Sub Impressao()
On Error GoTo Errado

  Dim iCounter As Integer
  Dim X        As Printer
  Dim iPg1     As Integer
  Dim iPg2     As Integer
  
  With Escolha
    .DialogTitle = "Selecione a impressora"
    .CancelError = True
    If Max > 0 Then
      .FromPage = 1
      .ToPage = 1
      .Max = Max + 1
      .FLAGS = cdlPDDisablePrintToFile Or cdlPDReturnDC Or cdlPDAllPages Or cdlPDUseDevModeCopies Or cdlPDNoSelection
    ElseIf Max = 0 Then
      .FLAGS = cdlPDDisablePrintToFile Or cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDUseDevModeCopies Or cdlPDNoSelection
    Else
      MsgBox "Nenhum bloqueto para imprimir.", vbInformation + vbOKOnly, Caption
      Exit Sub
    End If
    .ShowPrinter
    
    
    Printer.Orientation = cdlPortrait
    
    If Not VerificaFlags(.FLAGS) Then
      MousePointer = vbHourglass
      BloquetosLpt 1, Max + 1
      MousePointer = vbDefault
    Else
      iPg1 = .FromPage
      iPg2 = .ToPage
      MousePointer = vbHourglass
      BloquetosLpt iPg1, iPg2
      MousePointer = vbDefault
    End If
  End With

Fim:
  Exit Sub

Errado:
  If Err.Number = 32755 Then
    Resume Fim
  Else
    MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbCritical + vbOKOnly, Caption
    Resume Fim
  End If

End Sub

'=============================================================
'Bloquetos bancários
'=============================================================
Public Sub Bloquetos(ByVal sMes As String, ByVal nCond As Long)
  
  If nCond > 0 Then
    Set rsSelecao = db.OpenRecordset("Select * From Boletos Where Tran = '" & sMes & "' and cond = " & nCond & " Order By Nome;", dbOpenDynaset)
  Else
    Set rsSelecao = db.OpenRecordset("Select * From Boletos Where Tran = '" & sMes & "' Order By Nome;", dbOpenDynaset)
  End If
  DoEvents
  sMesprn = sMes
  
  With rsSelecao
    nRel = "Bloquetos bancários"
    If .RecordCount > 0 Then
      Iniciar
      .MoveLast
      .MoveFirst
      Atual = 0
      BoletoComp
      Max = .RecordCount - 1
'      nPage = 1
    End If
  End With
  
  AnunciaPagina Atual + 1

End Sub
  
Private Sub BoletoComp()
  Dim iSize     As Integer
  Dim iAtual    As Single
  Dim sFonte    As String
  Dim SelDebito As Recordset
  Dim dTotal    As Double
  Dim NpER      As Double
  Dim nFree     As Integer
  Dim sMens     As String
  Dim sMensagem() As String
  Dim i As Integer
  Dim iCounter  As Integer
  Dim SelAtraso As Recordset
  Dim selFora   As Recordset
  Dim somafora As Double
  
  nFree = FreeFile
  iCounter = 1
  
  tbAssociados.Index = "codigoid"
  tbCondominio.Index = "codigoid"
  
  With rsSelecao
    If .RecordCount > 0 Then
      'Cabeçalho
      If Dir(sFormataCaminho(App.Path) & "logotipo.jpg") <> "" Then
        PrintDesenho 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "logotipo.jpg"
      End If
      PrintDireita 0.6, 3, vbEmpresa.Empresa, "Arial", 14, True
      PrintDireita 1.2, 3, vbEmpresa.Endereco, "Arial", 9, False
      PrintDireita 1.2, 10, vbEmpresa.Bairro, "Arial", 9, False
      PrintDireita 1.2, 15, vbEmpresa.Cidade, "Arial", 9, False
      PrintDireita 1.6, 3, vbEmpresa.Estado, "Arial", 9, False
      PrintDireita 1.6, 4, vbEmpresa.Cep, "Arial", 9, False
      PrintDireita 1.6, 6.5, vbEmpresa.Cnpj, "Arial", 9, False
      PrintDireita 1.6, 10.5, vbEmpresa.Inscricao, "Arial", 9, False
      PrintDireita 1.6, 15, vbEmpresa.Fones, "Arial", 9, False
      PrintLinha 2, 1, 2.05, 20
      
      'Associado
      PrintDireita 2.4, 1, !cdsc, "Arial", 10, True
      PrintDireita 2.4, 2.5, !Nome, "Arial", 10, True
      tbAssociados.Seek "=", !cdsc
      If Not tbAssociados.NoMatch Then
        PrintEsquerda 2.4, 16, tbAssociados!Tipo, "Arial", 10, True
        PrintDireita 2.4, 16.4, tbAssociados!Apartamento, "Arial", 10, True
      End If
      PrintDireita 2.9, 1, !Ende, "Arial", 10, False
      PrintDireita 2.9, 11, !Bair, "Arial", 10, False
      PrintDireita 3.4, 1, !Cida, "Arial", 10, False
      PrintDireita 3.4, 11, !Esta, "Arial", 10, False
      PrintDireita 3.4, 12.5, !Cep, "Arial", 10, False
      
      'Mensagem
      iAtual = 5
      sMens = !Texto & ""
      sMensagem = Split(sMens, "|")
      For i = 0 To UBound(sMensagem)
        PrintDireita iAtual, 1, sMensagem(i), "Arial", 9, False
        iAtual = iAtual + 0.4
      Next i
      
      'Débitos do mês
      PrintLinha 5, 14.5, 9, 14.52
      Set SelDebito = db.OpenRecordset("Select TIPO, sum(VALOR) AS NVALOR From DESPESAS Where CONDOMINIO = " & !Cond & " And MES = '" & !TRAN & "' Group By TIPO Order By TIPO;", dbOpenDynaset)
      
      iAtual = 5.3
      dTotal = 0
      With SelDebito
        PrintDireita 5, 15, "Resumo", "Arial", 6, True
        If .RecordCount > 0 Then
          .MoveLast
          .MoveFirst
          Do While Not .EOF
            PrintDireita iAtual, 15, AchaTipo(!Tipo), "Arial Narrow", 6, False
            PrintEsquerda iAtual, 19.5, Format$(!nValor, "#,##0.00"), "Arial Narrow", 6, False
            dTotal = dTotal + !nValor
            iAtual = iAtual + 0.3
            .MoveNext
          Loop
          'fundo de caixa
          tbCondominio.Seek "=", rsSelecao!Cond
          If Not tbCondominio.NoMatch Then
            NpER = Round(dTotal * tbCondominio!Condominio / 100, 2)
          Else
            NpER = Round(dTotal * 0.1)
          End If
          PrintDireita iAtual, 15, "Fundo de reserva", "Arial Narrow", 6
          PrintEsquerda iAtual, 19.5, Format$(NpER, "#,##0.00"), "Arial Narrow", 6
          dTotal = dTotal + NpER
          iAtual = iAtual + 0.3
          PrintDireita iAtual, 15, "Total", "Arial Narrow", 6, True
          PrintEsquerda iAtual, 19.5, Format$(dTotal, "#,##0.00"), "Arial Narrow", 6, True
        End If
      End With
      Set SelDebito = Nothing
      
'      'Boletos em atraso
'      Set SelAtraso = db.OpenRecordset("Select * From BOLETOS Where CDSC = " & !cdsc & " And (PAGO = 'N' Or PAGO Is Null) And VCTO < #" & Format$(!Data, "mm/dd/yyyy") & "# Order By BOLE;", dbOpenDynaset)
'      iAtual = 5.3
'      dTotal = 0
'      With SelAtraso
'        If .RecordCount > 0 Then
'          PrintDireita 5, 1, "Atrasados", "Arial", 6, True
'          .MoveLast
'          .MoveFirst
'          Do While Not .EOF And (iAtual < 9)
'            Set selFora = db.OpenRecordset("Select * From porfora Where associado = " & !cdsc & " And mes = '" & !TRAN & "';", dbOpenDynaset)
'            somafora = 0
'            With selFora
'              If .RecordCount > 0 Then
'                .MoveLast
'                .MoveFirst
'                Do While Not .EOF
'                  somafora = somafora + !Valor
'                  .MoveNext
'                Loop
'              End If
'            End With
'            Set selFora = Nothing
'            PrintDireita iAtual, 1, Format$(!BOLE, "#0"), "Arial Narrow", 6
'            PrintEsquerda iAtual, 3, Format$(!MENS + somafora, "#,##0.00"), "Arial Narrow", 6
'            dTotal = dTotal + !MENS + somafora
'            iAtual = iAtual + 0.3
'            .MoveNext
'          Loop
'          If Not .EOF Then
'            PrintDireita iAtual, 1, "?", "Arial Narrow", 6, True
'            PrintEsquerda iAtual, 3, "******", "Arial Narrow", 6, True
'          Else
'            PrintDireita iAtual, 1, "Total", "Arial Narrow", 6, True
'            PrintEsquerda iAtual, 3, Format$(dTotal, "#,##0.00"), "Arial Narrow", 6, True
'          End If
'          PrintDireita iAtual, 1, "Total", "Arial Narrow", 6, True
'          PrintEsquerda iAtual, 3, Format$(dTotal, "#,##0.00"), "Arial Narrow", 6, True
'        End If
'      End With
'      Set SelAtraso = Nothing
      
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", rsSelecao!Cond
        If Not .NoMatch Then
          If (!Frase & "") <> "" Then
            PrintDireita 9.5, 1, !Frase & "", "Arial", 7, False
          Else
            PrintDireita 9.5, 1, "O(s) valore(s) em atraso consta(m) de correção monetária.", "Arial", 7, False
          End If
        Else
          PrintDireita 9.5, 1, "O(s) valore(s) em atraso consta(m) de correção monetária.", "Arial", 7, False
        End If
      End With
      
      PrintEsquerda 9.5, 20, "Maiores informações favor procurar a secretaria.", "Arial", 7, False
      PrintEsquerda 10, 20, "Esta parte não precisa ser levada ao banco.", "Arial", 7, False
      PrintDireita 10, 1, "Cortar", "Arial", 7, False
      PrintLinha 10.3, 1, 10.32, 20, vbDot
      
      'Recibo do sacado
      PrintLinha 10.5, 1, 10.53, 20
      PrintLinha 11.1, 1, 11.13, 20
      PrintLinha 12.1, 1, 12.13, 20
      PrintLinha 12.8, 1, 12.83, 20
      PrintLinha 14.5, 1, 14.53, 20
      PrintLinha 17, 1, 17.02, 20, vbDot
      PrintLinha 10.5, 4, 11.1, 4.03
      PrintLinha 10.5, 6.5, 11.1, 6.53
      PrintLinha 11.1, 15.5, 14.5, 15.53
      PrintLinha 13.65, 15.5, 13.68, 20
      If Dir(sFormataCaminho(App.Path) & "banco.jpg") <> "" Then
        PrintDesenho 10.55, 1.1, 2, 0.5, sFormataCaminho(App.Path) & "banco.jpg"
      End If
'        PrintEsquerda   10.55, 3.7, "CAIXA", "Arial Narrow", 13, True
      PrintDireita 10.55, 4.5, vbIdent & "-" & DigAgencia, "Arial Narrow", 13, True
      PrintDireita 10.55, 7, !CODI, "Arial Narrow", 13, True
      PrintDireita 11.13, 1.3, "Local de pagamento", "Arial", 6, False
      PrintDireita 11.5, 1.5, "Pagável preferencialmente nas casas lotéricas", "Arial", 10, True
      PrintDireita 11.13, 15.8, "Vencimento", "Arial", 6, False
      PrintEsquerda 11.5, 19.5, Format$(!vcto, "dd/mm/yyyy"), "Arial", 12, True
      PrintDireita 12.13, 1.3, "Cedente", "Arial", 6, False
      PrintDireita 12.4, 1.5, !Condominio & "", "Arial", 9, True
      PrintEsquerda 12.4, 15, Format$(!CGC, String(18, "0")), "Arial", 9, True
      PrintDireita 12.13, 15.8, "Agência/Código cedente", "Arial", 6, False
      PrintEsquerda 12.36, 19.5, !MENSAGEM, "Arial", 9, False
      PrintDireita 12.83, 1.3, "Sacado", "Arial", 6, False
      PrintDireita 13, 2.5, !cdsc, "Arial", 8, True
      PrintDireita 13, 3.7, !Nome, "Arial", 8, True
      PrintEsquerda 13, 15, !Cpf & "", "Arial", 8, True
      PrintDireita 13.3, 2.5, !Ende, "Arial", 8, False
      PrintDireita 13.3, 9.3, !Bair, "Arial", 8, False
      PrintDireita 13.6, 2.5, !Cida, "Arial", 8, False
      PrintDireita 13.6, 9.3, !Esta, "Arial", 8, False
      PrintDireita 13.6, 10.8, !Cep, "Arial", 8, False
      PrintDireita 14.1, 1.3, "Sacador/Avalista", "Arial", 6, False
      PrintDireita 12.83, 15.8, "Nosso número", "Arial", 6, False
      PrintEsquerda 13.25, 19.5, !DIGITAVAL, "Arial", 9, True
      PrintDireita 13.68, 15.8, "(=) Valor do documento", "Arial", 6, False
      PrintEsquerda 13.98, 19.5, Format$(!VALR, "#,##0.00"), "Arial", 8, True
      PrintDireita 14.53, 1.53, "Autenticação mecânica/RECIBO DO SACADO", "Arial", 6, False
      PrintEsquerda 14.53, 20, "J & F Software's Ltda. 31-3891-4298/0413", "Arial", 6, False, True
      PrintDireita 16.7, 1.3, "Cortar", "Arial", 6, False
      
      
      'Ficha de compensação
      PrintLinha 17.5, 1, 17.53, 20
      PrintLinha 18.5, 1, 18.53, 20
      PrintLinha 19.5, 1, 19.53, 20
      PrintLinha 20.3, 1, 20.33, 20
      PrintLinha 21.1, 1, 21.13, 20
      PrintLinha 21.9, 1, 21.93, 20
      
      PrintLinha 17.5, 4, 18.5, 4.03
      PrintLinha 17.5, 6.5, 18.5, 6.53
      
      PrintLinha 18.5, 15.5, 27.5, 15.53
      
      PrintLinha 20.3, 4.2, 21.1, 4.22
      PrintLinha 20.3, 8.8, 21.1, 8.82
      PrintLinha 20.3, 10.8, 21.1, 10.82
      PrintLinha 20.3, 12.3, 21.1, 12.32
      
      PrintLinha 21.1, 5.3, 21.9, 5.32
      PrintLinha 21.1, 7.4, 21.9, 7.42
      PrintLinha 21.1, 9.1, 21.9, 9.12
      PrintLinha 21.1, 12.5, 21.9, 12.52
      
      PrintLinha 22.7, 15.5, 22.72, 20
      PrintLinha 23.5, 15.5, 23.52, 20
      PrintLinha 24.3, 15.5, 24.32, 20
      PrintLinha 25.1, 15.5, 25.12, 20
      
      PrintLinha 25.8, 1, 25.83, 20
      
      PrintLinha 27.5, 1, 27.53, 20
      
      PrintDireita 27.7, 1, !CDBARRA, "Interleaved2of5-regular", 25, False
      PrintDireita 28.1, 1, !CDBARRA, "Interleaved2of5-regular", 25, False
      PrintDireita 27.55, 14, "Autenticação mecânica/FICHA COMPENSAÇÃO", "Arial Narrow", 7
      
      If Dir(sFormataCaminho(App.Path) & "banco.jpg") <> "" Then
        PrintDesenho 17.7, 1.1, 2.6, 0.7, sFormataCaminho(App.Path) & "banco.jpg"
      End If
'        PrintEsquerda   17.9, 3.7, "CAIXA", "Arial Narrow", 13, True
      PrintDireita 17.9, 4.5, vbIdent & "-" & DigAgencia, "Arial Narrow", 13, True
      PrintDireita 17.9, 7, !CODI, "Arial Narrow", 13, True
      PrintDireita 18.54, 1.3, "Local de pagamento", "Arial", 6, False
      PrintDireita 18.8, 1.5, "Pagável preferencialmente nas casas lotéricas", "Arial", 10, True
      PrintDireita 18.54, 15.8, "Vencimento", "Arial", 6, False
      PrintEsquerda 18.8, 19.5, Format$(!vcto, "dd/mm/yyyy"), "Arial", 12, True
      PrintDireita 19.53, 1.3, "Cedente", "Arial", 6, False
      PrintDireita 19.78, 1.5, !Condominio & "", "Arial", 9, True
      PrintEsquerda 19.78, 15, Format(!CGC, String(18, "0")), "Arial", 9, True
      PrintDireita 19.53, 15.8, "Agência/Código cedente", "Arial", 6, False
      PrintEsquerda 19.78, 19.5, !MENSAGEM, "Arial", 9, False
      PrintDireita 20.33, 1.3, "Data do documento", "Arial Narrow", 6, False
      PrintDireita 20.33, 4.5, "Número do documento", "Arial Narrow", 6, False
      PrintDireita 20.33, 9, "Esp. docum.", "Arial Narrow", 6, False
      PrintDireita 20.33, 11, "Aceite", "Arial Narrow", 6, False
      PrintDireita 20.33, 12.5, "Data do processamento", "Arial Narrow", 6, False
      PrintDireita 20.55, 1.5, Format$(!Data, "dd/mm/yyyy"), "Arial Narrow", 9, False
      PrintDireita 20.55, 12.5, Format$(!Data, "dd/mm/yyyy"), "Arial Narrow", 9, False
      PrintDireita 21.13, 1.3, "Uso do banco", "Arial Narrow", 6, False
      PrintDireita 21.13, 5.5, "Carteira", "Arial Narrow", 6, False
      PrintDireita 21.13, 7.6, "Espécie", "Arial Narrow", 6, False
      PrintDireita 21.13, 9.3, "Quantidade", "Arial Narrow", 6, False
      PrintDireita 21.13, 12.7, "Valor", "Arial Narrow", 6, False
      PrintDireita 21.4, 5.5, "SR", "Arial", 9, False
      PrintDireita 21.4, 7.6, "R$", "Arial", 9, False
      
      PrintDireita 20.33, 15.8, "Nosso número", "Arial Narrow", 6, False
      PrintEsquerda 20.55, 19.5, !DIGITAVAL, "Arial", 8, True

      PrintDireita 21.13, 15.8, "(=) Valor do documento", "Arial Narrow", 6, False
      PrintEsquerda 21.4, 19.5, Format$(!VALR, "#,##0.00"), "Arial", 8, True

      PrintDireita 21.93, 1.3, "Instruções:   (Todas as informações deste bloqueto são de exclusiva responsabilidade do cedente)", "Arial", 7, True
      PrintDireita 22.5, 2, !INST1, "Arial", 10, False
      PrintDireita 23, 2, !INST2, "Arial", 10, False
      PrintDireita 23.5, 2, !INST3, "Arial", 10, False
      PrintDireita 24, 2, !INST4, "Arial", 10, False
      
      PrintDireita 21.93, 15.8, "(-) Desconto/Abatimento", "Arial Narrow", 6, False
      PrintDireita 22.73, 15.8, "(-) Outras deduções", "Arial Narrow", 6, False
      PrintDireita 23.53, 15.8, "(+) Mora/Multa", "Arial Narrow", 6, False
      PrintDireita 24.33, 15.8, "(+) Outros acréscimos", "Arial Narrow", 6, False
      'PrintEsquerda 24.7, 19.5, Format$(dTotal, "#,##0.00"), "Arial", 8, True
      PrintDireita 25.13, 15.8, "(=) Valor cobrado", "Arial Narrow", 6, False
      
      PrintDireita 25.83, 1.3, "Sacado", "Arial", 6, False
      PrintDireita 26, 2.5, !cdsc, "Arial", 8, True
      PrintDireita 26, 3.7, !Nome, "Arial", 8, True
      PrintDireita 26.3, 2.5, !Ende, "Arial", 8, False
      PrintDireita 26.3, 9.3, !Bair, "Arial", 8, False
      PrintDireita 26.6, 2.5, !Cida, "Arial", 8, False
      PrintDireita 26.6, 9.3, !Esta, "Arial", 8, False
      PrintDireita 26.6, 10.8, !Cep, "Arial", 8, False
      PrintDireita 27.1, 1.3, "Sacador/Avalista", "Arial", 6, False
      
      PrintDireita 25.83, 15.8, "CPF/CNPJ", "Arial", 6, False
      PrintEsquerda 26, 19.5, !Cpf & "", "Arial", 8, True
      PrintDireita 27.1, 15.8, "Código de baixa", "Arial", 6, False
    End If
  End With
  
End Sub

Private Sub BloquetosLpt(ByVal GoPage As Integer, ByVal EndPage As Integer)
  
  Dim iSize     As Integer
  Dim iAtual    As Single
  Dim sFonte    As String
  Dim SelDebito As Recordset
  Dim dTotal    As Double
  Dim NpER      As Double
  Dim nFree     As Integer
  Dim sMens     As String
  Dim sMensagem() As String
  Dim iCounter  As Integer
  Dim SelBol    As Recordset
  Dim SelAtraso As Recordset
  
  
  nFree = FreeFile
  iCounter = 1
  
  tbCondominio.Index = "codigoid"
  
  Printer.ScaleMode = vbCentimeters
  'Printer.PaperSize = vbPRPSLegal
  With rsSelecao
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If iCounter < GoPage Then
          iCounter = iCounter + 1
          .MoveNext
          GoTo Proximo
        End If
        If iCounter > EndPage Then
          GoTo Parar
        End If
        Progre.Label1.Caption = "Imprimindo página: " + Str(iCounter)
        Progre.Label1.Refresh
        'Cabeçalho
        If Dir(sFormataCaminho(App.Path) & "logotipo.jpg") <> "" Then
          PrnDesenho 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "logotipo.jpg"
        End If
        PrnDireita 0.6, 3, vbEmpresa.Empresa, "Arial", 14, True
        PrnDireita 1.2, 3, vbEmpresa.Endereco, "Arial", 9, False
        PrnDireita 1.2, 10, vbEmpresa.Bairro, "Arial", 9, False
        PrnDireita 1.2, 15, vbEmpresa.Cidade, "Arial", 9, False
        PrnDireita 1.6, 3, vbEmpresa.Estado, "Arial", 9, False
        PrnDireita 1.6, 4, vbEmpresa.Cep, "Arial", 9, False
        PrnDireita 1.6, 6.5, vbEmpresa.Cnpj, "Arial", 9, False
        PrnDireita 1.6, 10.5, vbEmpresa.Inscricao, "Arial", 9, False
        PrnDireita 1.6, 15, vbEmpresa.Fones, "Arial", 9, False
        PrnLinha 2, 1, 2.02, 20
        
        'Associado
        PrnDireita 2.4, 1, !cdsc, "Arial", 10, True
        PrnDireita 2.4, 2.5, !Nome, "Arial", 10, True
        tbAssociados.Seek "=", !cdsc
        If Not tbAssociados.NoMatch Then
          PrnEsquerda 2.4, 16, tbAssociados!Tipo, "Arial", 10, True
          PrnDireita 2.4, 16.4, tbAssociados!Apartamento, "Arial", 10, True
        End If
        PrnDireita 2.9, 1, !Ende & "", "Arial", 10, False
        PrnDireita 2.9, 11, !Bair & "", "Arial", 10, False
        PrnDireita 3.4, 1, !Cida & "", "Arial", 10, False
        PrnDireita 3.4, 11, !Esta & "", "Arial", 10, False
        PrnDireita 3.4, 12.5, !Cep & "", "Arial", 10, False
        
        'Mensagem
        iAtual = 5
        sMens = !Texto & ""
        sMensagem = Split(sMens, "|")
        For i = 0 To UBound(sMensagem)
          PrnDireita iAtual, 1, sMensagem(i), "Arial", 9, False
          iAtual = iAtual + 0.4
        Next i
        
        PrnLinha 5, 14.5, 9, 14.52
        Set SelDebito = db.OpenRecordset("Select TIPO, sum(VALOR) AS NVALOR From DESPESAS Where CONDOMINIO = " & !Cond & " And MES = '" & !TRAN & "' Group By TIPO Order By TIPO;", dbOpenDynaset)
        
        iAtual = 5.3
        dTotal = 0
        With SelDebito
          PrnDireita 5, 15, "Resumo", "Arial", 6, True
          If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While Not .EOF
              PrnDireita iAtual, 15, AchaTipo(!Tipo), "Arial Narrow", 6, False
              PrnEsquerda iAtual, 19.5, Format$(!nValor, "#,##0.00"), "Arial Narrow", 6, False
              dTotal = dTotal + !nValor
              iAtual = iAtual + 0.3
              .MoveNext
            Loop
            'fundo de caixa
            tbCondominio.Seek "=", rsSelecao!Cond
            If Not tbCondominio.NoMatch Then
              NpER = Round(dTotal * tbCondominio!Condominio / 100, 2)
            Else
              NpER = Round(dTotal * 0.1)
            End If
            PrnDireita iAtual, 15, "Fundo de reserva", "Arial Narrow", 6
            PrnEsquerda iAtual, 19.5, Format$(NpER, "#,##0.00"), "Arial Narrow", 6
            dTotal = dTotal + NpER
            iAtual = iAtual + 0.3
            PrnDireita iAtual, 15, "Total", "Arial Narrow", 6, True
            PrnEsquerda iAtual, 19.5, Format$(dTotal, "#,##0.00"), "Arial Narrow", 6, True
          End If
        End With
        Set SelDebito = Nothing
        
'        'Boletos em atraso
'        Set SelAtraso = db.OpenRecordset("Select * From BOLETOS Where CDSC = " & !cdsc & " And (PAGO = 'N' Or PAGO Is Null) And VCTO < #" & Format$(!Data, "mm/dd/yyyy") & "# Order By BOLE;", dbOpenDynaset)
'
'        iAtual = 5.3
'        dTotal = 0
'        With SelAtraso
'          If .RecordCount > 0 Then
'            PrnDireita 5, 1, "Atrasados", "Arial", 6, True
'            .MoveLast
'            .MoveFirst
'            Do While Not .EOF And (iAtual < 9)
'              PrnDireita iAtual, 1, Format$(!BOLE, "#0"), "Arial Narrow", 6
'              PrnEsquerda iAtual, 3, Format$(!Corrigido, "#,##0.00"), "Arial Narrow", 6
'              dTotal = dTotal + !Corrigido
'              iAtual = iAtual + 0.3
'              .MoveNext
'            Loop
'            If Not .EOF Then
'              PrnDireita iAtual, 1, "?", "Arial Narrow", 6, True
'              PrnEsquerda iAtual, 3, "******", "Arial Narrow", 6, True
'            Else
'              PrnDireita iAtual, 1, "Total", "Arial Narrow", 6, True
'              PrnEsquerda iAtual, 3, Format$(dTotal, "#,##0.00"), "Arial Narrow", 6, True
'            End If
'          End If
'        End With
'        Set SelAtraso = Nothing
        
        With tbCondominio
          .Index = "codigoid"
          .Seek "=", rsSelecao!Cond
          If Not .NoMatch Then
            If (!Frase & "") <> "" Then
              PrnDireita 9.5, 1, !Frase & "", "Arial", 7, False
            Else
              PrnDireita 9.5, 1, "O(s) valore(s) em atraso consta(m) de correção monetária.", "Arial", 7, False
            End If
          Else
            PrnDireita 9.5, 1, "O(s) valore(s) em atraso consta(m) de correção monetária.", "Arial", 7, False
          End If
        End With
        
        PrnEsquerda 9.5, 20, "Maiores informações favor procurar a secretaria.", "Arial", 7, False
        PrnEsquerda 10, 20, "Esta parte não precisa ser levada ao banco.", "Arial", 7, False
        
        PrnDireita 10, 1, "Cortar", "Arial", 7, False
        PrnLinha 10.3, 1, 10.32, 20, vbDot
        
        'Recibo do sacado
        PrnLinha 10.5, 1, 10.53, 20
        PrnLinha 11.1, 1, 11.13, 20
        PrnLinha 12.1, 1, 12.13, 20
        PrnLinha 12.8, 1, 12.83, 20
        PrnLinha 14.5, 1, 14.53, 20
        PrnLinha 10.5, 4, 11.1, 4.03
        PrnLinha 10.5, 6.5, 11.1, 6.53
        PrnLinha 11.1, 15.5, 14.5, 15.53
        PrnLinha 13.65, 15.5, 13.68, 20
        If Dir(sFormataCaminho(App.Path) & "banco.jpg") <> "" Then
          PrnDesenho 10.55, 1.1, 2, 0.5, sFormataCaminho(App.Path) & "banco.jpg"
        End If
'        PrnEsquerda 10.55, 3.7, "CAIXA", "Arial Narrow", 13, True
        PrnDireita 10.55, 4.5, vbIdent & "-" & DigAgencia, "Arial Narrow", 13, True
        PrnDireita 10.55, 7, !CODI, "Arial Narrow", 13, True
        PrnDireita 11.13, 1.3, "Local de pagamento", "Arial", 6, False
        PrnDireita 11.5, 1.5, "Pagável preferencialmente nas casas lotéricas", "Arial", 10, True
        PrnDireita 11.13, 15.8, "Vencimento", "Arial", 6, False
        PrnEsquerda 11.5, 19.5, Format$(!vcto, "dd/mm/yyyy"), "Arial", 12, True
        PrnDireita 12.13, 1.3, "Cedente", "Arial", 6, False
        
        PrnDireita 12.4, 1.5, !Condominio, "Arial", 9, True
        PrnEsquerda 12.4, 15, Format(!CGC, String(18, "0")), "Arial", 9, True
        
        PrnDireita 12.13, 15.8, "Agência/Código cedente", "Arial", 6, False
        PrnEsquerda 12.36, 19.5, !MENSAGEM, "Arial", 9, False
        PrnDireita 12.83, 1.3, "Sacado", "Arial", 6, False
        PrnDireita 13, 2.5, !cdsc, "Arial", 8, True
        PrnDireita 13, 3.7, !Nome, "Arial", 8, True
        PrnEsquerda 13, 15, !Cpf & "", "Arial", 8, True
        PrnDireita 13.3, 2.5, !Ende & "", "Arial", 8, False
        PrnDireita 13.3, 9.3, !Bair & "", "Arial", 8, False
        PrnDireita 13.6, 2.5, !Cida & "", "Arial", 8, False
        PrnDireita 13.6, 9.3, !Esta & "", "Arial", 8, False
        PrnDireita 13.6, 10.8, !Cep & "", "Arial", 8, False
        PrnDireita 14.1, 1.3, "Sacador/Avalista", "Arial", 6, False
        PrnDireita 12.83, 15.8, "Nosso número", "Arial", 6, False
        PrnEsquerda 13.25, 19.5, !DIGITAVAL, "Arial", 9, True
        PrnDireita 13.68, 15.8, "(=) Valor do documento", "Arial", 6, False
        PrnEsquerda 13.98, 19.5, Format$(!VALR, "#,##0.00"), "Arial", 8, True
        PrnDireita 14.53, 1.53, "Autenticação mecânica/RECIBO DO SACADO", "Arial", 6, False
        PrnEsquerda 14.53, 20, "J & F Software's Ltda. 31-3891-4298/0413", "Arial", 6, False, True
        
        PrnDireita 15.7, 1.3, "Cortar", "Arial", 6, False
        PrnLinha 16, 1, 16.02, 20, vbDot
        
        'Ficha de compensação
        PrnLinha 16.5, 1, 16.52, 20
        PrnLinha 17.5, 1, 17.52, 20
        PrnLinha 18.5, 1, 18.52, 20
        PrnLinha 19.3, 1, 19.32, 20
        PrnLinha 20.1, 1, 20.12, 20
        PrnLinha 20.9, 1, 20.92, 20
        
        PrnLinha 16.5, 4, 17.5, 4.03
        PrnLinha 16.5, 6.5, 17.5, 6.53
        
        PrnLinha 17.5, 15.5, 26.5, 15.53
        
        PrnLinha 19.3, 4.2, 20.1, 4.22
        PrnLinha 19.3, 8.8, 20.1, 8.82
        PrnLinha 19.3, 10.8, 20.1, 10.82
        PrnLinha 19.3, 12.3, 20.1, 12.32
        
        PrnLinha 20.1, 5.3, 20.9, 5.32
        PrnLinha 20.1, 7.4, 20.9, 7.42
        PrnLinha 20.1, 9.1, 20.9, 9.12
        PrnLinha 20.1, 12.5, 20.9, 12.52
        
        PrnLinha 21.7, 15.5, 21.72, 20
        PrnLinha 22.5, 15.5, 22.52, 20
        PrnLinha 23.3, 15.5, 23.32, 20
        PrnLinha 24.1, 15.5, 24.12, 20
        
        PrnLinha 24.8, 1, 24.83, 20
        
        PrnLinha 26.5, 1, 26.53, 20
        
        If Dir(sFormataCaminho(App.Path) & "banco.jpg") <> "" Then
          PrnDesenho 16.7, 1.1, 2.6, 0.7, sFormataCaminho(App.Path) & "banco.jpg"
        End If
'        PrnEsquerda 17.9, 3.7, "CAIXA", "Arial Narrow", 13, True
        PrnDireita 16.9, 4.5, vbIdent & "-" & DigAgencia, "Arial Narrow", 13, True
        PrnDireita 16.9, 7, !CODI, "Arial Narrow", 13, True
        PrnDireita 17.54, 1.3, "Local de pagamento", "Arial", 6, False
        PrnDireita 17.8, 1.5, "Pagável preferencialmente nas casas lotéricas", "Arial", 10, True
        PrnDireita 17.54, 15.8, "Vencimento", "Arial", 6, False
        PrnEsquerda 17.8, 19.5, Format$(!vcto, "dd/mm/yyyy"), "Arial", 12, True
        PrnDireita 18.53, 1.3, "Cedente", "Arial", 6, False
        
        PrnDireita 18.78, 1.5, !Condominio & "", "Arial", 9, True
        PrnEsquerda 18.78, 15, Format(!CGC, String(18, "0")), "Arial", 9, True
        
        PrnDireita 18.53, 15.8, "Agência/Código cedente", "Arial", 6, False
        PrnEsquerda 18.78, 19.5, !MENSAGEM, "Arial", 9, False
        PrnDireita 19.33, 1.3, "Data do documento", "Arial Narrow", 6, False
        PrnDireita 19.33, 4.5, "Número do documento", "Arial Narrow", 6, False
        PrnDireita 19.33, 9, "Esp. docum.", "Arial Narrow", 6, False
        PrnDireita 19.33, 11, "Aceite", "Arial Narrow", 6, False
        PrnDireita 19.33, 12.5, "Data do processamento", "Arial Narrow", 6, False
        PrnDireita 19.55, 1.5, Format$(!Data, "dd/mm/yyyy"), "Arial Narrow", 9, False
        PrnDireita 19.55, 12.5, Format$(!Data, "dd/mm/yyyy"), "Arial Narrow", 9, False
        PrnDireita 20.13, 1.3, "Uso do banco", "Arial Narrow", 6, False
        PrnDireita 20.13, 5.5, "Carteira", "Arial Narrow", 6, False
        PrnDireita 20.13, 7.6, "Espécie", "Arial Narrow", 6, False
        PrnDireita 20.13, 9.3, "Quantidade", "Arial Narrow", 6, False
        PrnDireita 20.13, 12.7, "Valor", "Arial Narrow", 6, False
        PrnDireita 20.4, 5.5, "SR", "Arial", 9, False
        PrnDireita 20.4, 7.6, "R$", "Arial", 9, False
        
        PrnDireita 19.33, 15.8, "Nosso número", "Arial Narrow", 6, False
        PrnEsquerda 19.55, 19.5, !DIGITAVAL, "Arial", 8, True

        PrnDireita 20.13, 15.8, "(=) Valor do documento", "Arial Narrow", 6, False
        PrnEsquerda 20.4, 19.5, Format$(!VALR, "#,##0.00"), "Arial", 8, True

        PrnDireita 20.93, 1.3, "Instruções:   (Todas as informações deste bloqueto são de exclusiva responsabilidade do cedente)", "Arial", 7, True
        PrnDireita 21.5, 2, !INST1, "Arial", 10, False
        PrnDireita 22, 2, !INST2, "Arial", 10, False
        PrnDireita 22.5, 2, !INST3, "Arial", 10, False
        PrnDireita 23, 2, !INST4, "Arial", 10, False
        
        PrnDireita 20.93, 15.8, "(-) Desconto/Abatimento", "Arial Narrow", 6, False
        PrnDireita 21.73, 15.8, "(-) Outras deduções", "Arial Narrow", 6, False
        PrnDireita 22.53, 15.8, "(+) Mora/Multa", "Arial Narrow", 6, False
        PrnDireita 23.33, 15.8, "(+) Outros acréscimos", "Arial Narrow", 6, False
        'PrnEsquerda 24.7, 19.5, Format$(dTotal, "#,##0.00"), "Arial", 8, True
        PrnDireita 24.13, 15.8, "(=) Valor cobrado", "Arial Narrow", 6, False
        
        PrnDireita 24.83, 1.3, "Sacado", "Arial", 6, False
        PrnDireita 25, 2.5, !cdsc, "Arial", 8, True
        PrnDireita 25, 3.7, !Nome, "Arial", 8, True
        PrnDireita 25.3, 2.5, !Ende & "", "Arial", 8, False
        PrnDireita 25.3, 9.3, !Bair & "", "Arial", 8, False
        PrnDireita 25.6, 2.5, !Cida & "", "Arial", 8, False
        PrnDireita 25.6, 9.3, !Esta & "", "Arial", 8, False
        PrnDireita 25.6, 10.8, !Cep & "", "Arial", 8, False
        PrnDireita 26.1, 1.3, "Sacador/Avalista", "Arial", 6, False
        
        PrnDireita 25.83, 15.8, "CPF/CNPJ", "Arial", 6, False
        PrnEsquerda 25, 19.5, !Cpf & "", "Arial", 8, True
        PrnDireita 26.1, 15.8, "Código de baixa", "Arial", 6, False
        'JÁ
        PrnDireita 26.55, 14, "Autenticação mecânica/FICHA COMPENSAÇÃO", "Arial Narrow", 7
        PrnDireita 26.7, 1, !CDBARRA, "Interleaved2of5-regular", 25, False
        PrnDireita 27.1, 1, !CDBARRA, "Interleaved2of5-regular", 25, False
        
        iCounter = iCounter + 1
        .MoveNext
        
        If Not .EOF Then
          Printer.NewPage
        End If
Proximo:
      Loop
Parar:
      Printer.EndDoc
    End If
  End With
  
End Sub

Private Function AchaTipo(ByVal iTipo As Integer) As String
  With tbTipoDesp
    .Index = "codigoid"
    .Seek "=", iTipo
    If Not .NoMatch Then
      AchaTipo = !descricao
    Else
      AchaTipo = "Outras"
    End If
  End With
End Function

