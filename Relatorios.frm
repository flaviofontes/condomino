VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Relatorios 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "Relatorios.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8.758
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   10.266
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
      TabIndex        =   4
      Top             =   4665
      Width           =   5130
   End
   Begin VB.VScrollBar Vertical 
      Height          =   3420
      Left            =   5070
      TabIndex        =   3
      Top             =   1170
      Width           =   285
   End
   Begin VB.PictureBox Rel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   3480
      Left            =   540
      ScaleHeight     =   6.033
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   7.435
      TabIndex        =   1
      Top             =   1065
      Width           =   4275
      Begin VB.PictureBox Pagina 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   11906
         Index           =   0
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
      Left            =   6075
      Top             =   180
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
            Picture         =   "Relatorios.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Relatorios.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Relatorios.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Relatorios.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Relatorios.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Relatorios.frx":099C
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
      Width           =   5820
      _ExtentX        =   10266
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
Attribute VB_Name = "Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Atual As Integer
Dim Max As Integer
Dim ChamaLpt As Integer
Dim rsSelecao As DAO.Recordset


Private Sub Atalhos_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1
      Impressao
    Case 3
      Pagina(Atual).Visible = False
      Atual = 0
      Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
      Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
      Pagina(Atual).ZOrder
      Pagina(Atual).Visible = True
      Pagina(Atual).Refresh
    Case 4
      Pagina(Atual).Visible = False
      Atual = Atual - 1
      If Atual <= 0 Then
        Atual = 0
        Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
        Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
        Pagina(Atual).ZOrder
        Pagina(Atual).Visible = True
        Pagina(Atual).Refresh
      ElseIf Atual > 0 Then
        Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
        Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
        Pagina(Atual).ZOrder
        Pagina(Atual).Visible = True
        Pagina(Atual).Refresh
      End If
    Case 5
      Pagina(Atual).Visible = False
      Atual = Atual + 1
      If Atual >= Max Then
        Atual = Max
        Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
        Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
        Pagina(Atual).ZOrder
        Pagina(Atual).Visible = True
        Pagina(Atual).Refresh
      ElseIf Atual < Max Then
        Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
        Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
        Pagina(Atual).ZOrder
        Pagina(Atual).Visible = True
        Pagina(Atual).Refresh
      End If
    Case 6
      Pagina(Atual).Visible = False
      Atual = Max
      Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
      Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
      Pagina(Atual).ZOrder
      Pagina(Atual).Visible = True
      Pagina(Atual).Refresh
    Case 8
      VB.Unload Me
  End Select
End Sub

Private Sub Form_Load()
  i = 0
  Atalhos.Buttons(1).Enabled = False
  Atalhos.Buttons(3).Enabled = False
  Atalhos.Buttons(4).Enabled = False
  Atalhos.Buttons(5).Enabled = False
  Atalhos.Buttons(6).Enabled = False
  Atalhos.Buttons(8).Enabled = False
  MousePointer = vbHourglass
End Sub

Private Sub Form_Resize()
  Horizontal.Left = 0
  Horizontal.Top = Me.ScaleHeight - Horizontal.Height
  Horizontal.Width = Me.ScaleWidth - Vertical.Width
  Vertical.Top = Atalhos.Height
  Vertical.Height = Me.ScaleHeight - Atalhos.Height
  Vertical.Left = Me.ScaleWidth - Vertical.Width
  Rel.Move 0, Atalhos.Height, Me.ScaleWidth - Vertical.Width, Me.ScaleHeight - Horizontal.Height
End Sub

Public Sub Iniciar()
  Pagina(0).ScaleMode = vbCentimeters
  Pagina(0).AutoRedraw = True
  Pagina(0).Height = 29.7
  Pagina(0).Width = 21
  Pagina(0).Left = Rel.Left + 0.5
  Pagina(0).Top = Rel.Top + 0.5
End Sub

Public Function NovaPagina() As Integer
  i = i + 1
  Load Pagina(i)
  Pagina(i).ScaleMode = vbCentimeters
  Pagina(i).AutoRedraw = True
  Pagina(i).Height = 29.7
  Pagina(i).Width = 21
  Pagina(i).Left = Rel.Left + 0.5
  Pagina(i).Top = Rel.Top + 0.5
  NovaPagina = i
End Function

Public Function FimDeImpressao()
  Pagina(0).Visible = True
  Pagina(0).ZOrder
  Pagina(0).Refresh
  Vertical.Min = -0.5
  Vertical.Max = 30
  Vertical.SmallChange = 1
  Horizontal.Min = -0.5
  Horizontal.Max = 22.5
  Horizontal.SmallChange = 1
  Atual = 0
  Max = i
  If Max = 1 Then
    Atalhos.Buttons(1).Enabled = True
    Atalhos.Buttons(3).Enabled = False
    Atalhos.Buttons(4).Enabled = False
    Atalhos.Buttons(5).Enabled = False
    Atalhos.Buttons(6).Enabled = False
    Atalhos.Buttons(8).Enabled = True
  Else
    Atalhos.Buttons(1).Enabled = True
    Atalhos.Buttons(3).Enabled = True
    Atalhos.Buttons(4).Enabled = True
    Atalhos.Buttons(5).Enabled = True
    Atalhos.Buttons(6).Enabled = True
    Atalhos.Buttons(8).Enabled = True
  End If
  Caption = Me.Caption & " - Total de páginas: " & (Max + 1)
  MousePointer = vbDefault
  Show
End Function


'=============================================================
'Imprime o texto a direita
'=============================================================
Public Sub PrintDireita(ByVal Index As Integer, ByVal Y As Single, ByVal X As Single, ByVal sTexto As String, ByVal sFonte As String, ByVal sSize As Integer, Optional ByVal Bold As Boolean, Optional ByVal Italico As Boolean)
  Pagina(Index).FontName = sFonte
  Pagina(Index).FontSize = sSize
  Pagina(Index).FontBold = Bold
  Pagina(Index).FontItalic = Italico
  Pagina(Index).CurrentX = X
  Pagina(Index).CurrentY = Y
  Pagina(Index).Print sTexto
End Sub

'=============================================================
'Imprime o texto a esquerda
'=============================================================
Public Sub PrintEsquerda(ByVal Index As Integer, ByVal Y As Single, ByVal X As Single, ByVal sTexto As String, ByVal sFonte As String, ByVal sSize As Integer, Optional ByVal Bold As Boolean, Optional Italico As Boolean)
  Pagina(Index).FontName = sFonte
  Pagina(Index).FontSize = sSize
  Pagina(Index).FontBold = Bold
  Pagina(Index).FontItalic = Italico
  Pagina(Index).CurrentX = X - Pagina(Index).TextWidth(sTexto)
  Pagina(Index).CurrentY = Y
  Pagina(Index).Print sTexto
End Sub

'=============================================================
'Imprime o texto centralizado
'=============================================================
Public Sub PrintCentralizado(ByVal Index As Integer, ByVal Y As Single, ByVal X As Single, ByVal sTexto As String, ByVal sFonte As String, ByVal sSize As Integer, Optional ByVal Bold As Boolean, Optional Italico As Boolean)
  Pagina(Index).FontName = sFonte
  Pagina(Index).FontSize = sSize
  Pagina(Index).FontBold = Bold
  Pagina(Index).FontItalic = Italico
  Pagina(Index).CurrentX = X - (Pagina(Index).TextWidth(sTexto) / 2)
  Pagina(Index).CurrentY = Y
  Pagina(Index).Print sTexto
End Sub

'=============================================================
'Imprime uma linha
'=============================================================
Public Sub PrintLinha(ByVal Index As Integer, ByVal Y1 As Single, ByVal X1 As Single, ByVal Y2 As Single, ByVal X2 As Single, Optional vbDraw As DrawStyleConstants)
  If Not IsMissing(vbDraw) Then
    Pagina(Index).DrawStyle = vbDraw
  Else
    Pagina(Index).DrawStyle = vbSolid
  End If
  Pagina(Index).Line (X1, Y1)-(X2, Y2), 0, BF
  Pagina(Index).DrawStyle = vbSolid
End Sub

'=============================================================
'Imprime um quadro
'=============================================================
Public Sub PrintQuadro(ByVal Index As Integer, ByVal Y1 As Single, ByVal X1 As Single, ByVal Y2 As Single, ByVal X2 As Single, vbDraw As DrawStyleConstants)
  If Not IsMissing(vbDraw) Then
    Pagina(Index).DrawStyle = vbDraw
  Else
    Pagina(Index).DrawStyle = vbSolid
  End If
  Pagina(Index).Line (X1, Y1)-(X2, Y2), 0, B
  Pagina(Index).DrawStyle = vbSolid
End Sub

'=============================================================
'Imprime desenho
'=============================================================
Public Sub PrintDesenho(ByVal Index As Integer, ByVal Y1 As Single, ByVal X1 As Single, ByVal Y2 As Single, ByVal X2 As Single, ByVal sFigura As String)
  If Dir$(sFigura) <> "" Then
    Pagina(Index).PaintPicture VB.LoadPicture(sFigura), X1, Y1, Y2, X2
  End If
End Sub

Private Sub Horizontal_Change()
  Pagina(Atual).Left = Rel.Left + 0.5 + (Horizontal.Value * -1)
End Sub

Private Sub Vertical_Change()
  Pagina(Atual).Top = Rel.Top + 0.5 + (Vertical.Value * -1)
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
      .Flags = cdlPDDisablePrintToFile Or cdlPDReturnDC Or cdlPDAllPages Or cdlPDUseDevModeCopies Or cdlPDNoSelection
    ElseIf Max = 0 Then
      .Flags = cdlPDDisablePrintToFile Or cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDUseDevModeCopies Or cdlPDNoSelection
    Else
      MsgBox "Nenhum bloqueto para imprimir.", vbInformation + vbOKOnly, Caption
      Exit Sub
    End If
    .ShowPrinter
'    For Each X In Printers
'      If X.hDC = .hDC Then
'        Set Printer = X
'        Exit For
'      End If
'    Next
    Printer.Orientation = cdlPortrait
    
    If Not VerificaFlags(.Flags) Then
      Select Case ChamaLpt
        Case 1
          DebitosReceberLpt 1, Max + 1
        Case 2
          BoletosQuitadosLpt 1, Max + 1
        Case 3
          PaisRespLpt 1, Max + 1
      End Select
    Else
      iPg1 = .FromPage
      iPg2 = .ToPage
      Select Case ChamaLpt
        Case 1
          DebitosReceberLpt iPg1, iPg2
        Case 2
          BoletosQuitadosLpt iPg1, iPg2
        Case 3
          PaisRespLpt iPg1, iPg2
      End Select
    End If
  End With

Fim:
  Exit Sub

Errado:
  If Err.Number = 32755 Then
    Resume Fim
  Else
    ErroPadrao Err.Number, Err.Description, Err.Source
    Resume Fim
  End If

End Sub


'=================================================
'Contas a receber
'=================================================
Public Sub DebitosReceber()
  
  Dim iAtual  As Single
  Dim iPage   As Integer
  Dim iSoma   As Double
  Dim iSub    As Double
  Dim iRef    As Integer
  Dim sIdent  As String
  Dim sSql    As String
  
  iPage = 1
  iSoma = 0
  iSub = 0
  iRef = 1
  sIdent = ""
  
  sSql = "SELECT DISTINCT Boletos.Data, Boletos.VCTO, Boletos.VALR, Boletos.CDSC, Boletos.NOME, Mensalidades.Situacao " & _
        "FROM Boletos LEFT JOIN Mensalidades ON Boletos.EXTR = Mensalidades.Boleto " & _
        "WHERE (((Mensalidades.Situacao)='N')) ORDER BY Boletos.Nome;"
  
  Set rsSelecao = dbfDados.OpenRecordset(sSql, dbOpenDynaset)
  DoEvents
  
  With rsSelecao
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Caption = "Listagem de débitos"
      Iniciar
      Do While Not .EOF
        DoEvents
        'Cabeçalho
        If Dir(sFormataCaminho(App.Path) & "coeducar.jpg") <> "" Then
          PrintDesenho iPage - 1, 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "coeducar.jpg"
        End If
        PrintDireita iPage - 1, 0.6, 3, vbCabecalho.Nome, "Arial", 14, True
        PrintDireita iPage - 1, 1.2, 3, vbCabecalho.Endereco, "Arial", 9, False
        PrintDireita iPage - 1, 1.2, 10, vbCabecalho.Bairro, "Arial", 9, False
        PrintDireita iPage - 1, 1.2, 15, vbCabecalho.Cidade, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 3, vbCabecalho.Estado, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 4, vbCabecalho.Cep, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 6.5, vbCabecalho.Cgc, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 10.5, vbCabecalho.Inscricao, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 15, vbCabecalho.Telefone, "Arial", 9, False
        PrintLinha iPage - 1, 2, 1, 2.03, 20, vbSolid
        PrintDireita iPage - 1, 2.5, 1, "Relatório de débitos a receber", "Arial", 10, True
        PrintDireita iPage - 1, 3, 1, "Data", "Arial", 10, True
        PrintDireita iPage - 1, 3, 3, "Ident.", "Arial", 10, True
        PrintDireita iPage - 1, 3, 4.3, "Nome", "Arial", 10, True
        PrintDireita iPage - 1, 3, 14, "Vencimento", "Arial", 10, True
        PrintEsquerda iPage - 1, 3, 20, "Valor", "Arial", 10, True
        PrintLinha iPage - 1, 3.4, 1, 3.42, 20, vbSolid
        iAtual = 3.6
        Do While Not .EOF And iAtual < 25.5
          PrintDireita iPage - 1, iAtual, 1, Format$(!Data, "dd/mm/yyyy"), "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 3, !CDSC & "", "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 4.3, !Nome & "", "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 14, Format$(!VCTO, "dd/mm/yyyy"), "Arial", 9, False
          PrintEsquerda iPage - 1, iAtual, 20, Format$(!Valr, "#,##0.00"), "Arial", 9, False
          iAtual = iAtual + 0.4
          iSoma = iSoma + (0 & !Valr)
          iSub = iSub + (0 & !Valr)
          iRef = iRef + 1
          .MoveNext
          DoEvents
        Loop
        If Not .EOF Then
          PrintLinha iPage - 1, 26, 1, 26.02, 20
          PrintDireita iPage - 1, 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
          PrintEsquerda iPage - 1, 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
          NovaPagina
          iPage = iPage + 1
        Else
          PrintDireita iPage - 1, 25.5, 10, "Total a receber", "Arial", 10, False
          PrintEsquerda iPage - 1, 25.5, 20, Format$(iSoma, "#,##0.00"), "Arial", 10, True
          PrintLinha iPage - 1, 26, 1, 26.02, 20
          PrintDireita iPage - 1, 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
          PrintEsquerda iPage - 1, 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
        End If
      Loop
    End If
  End With
         
  FimDeImpressao
  
  Set rsSelecao = Nothing
  
  ChamaLpt = 1

End Sub

Private Sub DebitosReceberLpt(ByVal Page1 As Single, ByVal Page2 As Single)
  
  Dim iAtual  As Single
  Dim iPage   As Integer
  Dim iSoma   As Double
  Dim iSub    As Double
  Dim iRef    As Integer
  Dim sIdent  As String
  Dim sSql    As String
  
  iPage = 1
  iSoma = 0
  iSub = 0
  iRef = 1
  sIdent = ""
  
  sSql = "SELECT DISTINCT Boletos.VCTO, Boletos.VALR, Boletos.CDSC, Boletos.NOME, Boletos.DATA, Boletos.BOLE, Mensalidades.Situacao " & _
        "FROM Boletos RIGHT JOIN Mensalidades ON Boletos.CDSC = Mensalidades.Identificacao " & _
        "Where (((Mensalidades.Situacao) = 'N')) ORDER BY Boletos.NOME;"

  Set rsSelecao = dbfDados.OpenRecordset(sSql, dbOpenDynaset)
  DoEvents
  
  With rsSelecao
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Printer.ScaleMode = vbCentimeters
      Do While Not .EOF
        DoEvents
        'Cabeçalho
        If iPage < Page1 Then
          GoTo Nova
        End If
        If iPage > Page2 Then
          GoTo Final
        End If
        If Dir(sFormataCaminho(App.Path) & "coeducar.jpg") <> "" Then
          PrnDesenho 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "coeducar.jpg"
        End If
        PrnDireita 0.6, 3, vbCabecalho.Nome, "Arial", 14, True
        PrnDireita 1.2, 3, vbCabecalho.Endereco, "Arial", 9, False
        PrnDireita 1.2, 10, vbCabecalho.Bairro, "Arial", 9, False
        PrnDireita 1.2, 15, vbCabecalho.Cidade, "Arial", 9, False
        PrnDireita 1.6, 3, vbCabecalho.Estado, "Arial", 9, False
        PrnDireita 1.6, 4, vbCabecalho.Cep, "Arial", 9, False
        PrnDireita 1.6, 6.5, vbCabecalho.Cgc, "Arial", 9, False
        PrnDireita 1.6, 10.5, vbCabecalho.Inscricao, "Arial", 9, False
        PrnDireita 1.6, 15, vbCabecalho.Telefone, "Arial", 9, False
        PrnLinha 2, 1, 2.03, 20, vbSolid
        PrnDireita 2.5, 1, "Relatório de débitos a receber", "Arial", 10, True
        PrnDireita 3, 1, "Data", "Arial", 10, True
        PrnDireita 3, 3, "Identific.", "Arial", 10, True
        PrnDireita 3, 4.3, "Nome", "Arial", 10, True
        PrnDireita 3, 14, "Vencimento", "Arial", 10, True
        PrnEsquerda 3, 20, "Valor", "Arial", 10, True
        PrnLinha 3.4, 1, 3.42, 20, vbSolid
Nova:
        iAtual = 3.6
        Do While Not .EOF And iAtual < 25.5
          If iPage >= Page1 And iPage <= Page2 Then
            PrnDireita iAtual, 1, Format$(!Data, "dd/mm/yyyy"), "Arial", 9, False
            PrnDireita iAtual, 3, !CDSC, "Arial", 9, False
            PrnDireita iAtual, 4.3, !Nome, "Arial", 9, False
            PrnDireita iAtual, 14, Format$(!VCTO, "dd/mm/yyyy"), "Arial", 9, False
            PrnEsquerda iAtual, 20, Format$(!Valr, "#,##0.00"), "Arial", 9, False
          End If
          iAtual = iAtual + 0.4
          iSoma = iSoma + !Valr
          iSub = iSub + !Valr
          iRef = iRef + 1
          .MoveNext
          DoEvents
        Loop
        If Not .EOF Then
          If iPage >= Page1 And iPage <= Page2 Then
            PrnLinha 26, 1, 26.02, 20
            PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
            PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            Printer.NewPage
          End If
          iPage = iPage + 1
        Else
Final:
          If iPage <= Page2 Then
            If iPage = (Max + 1) Then
              PrnDireita 25.5, 10, "Total a receber", "Arial", 10, False
              PrnEsquerda 25.5, 20, Format$(iSoma, "#,##0.00"), "Arial", 10, True
              PrnLinha 26, 1, 26.02, 20
              PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
              PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            Else
              PrnLinha 26, 1, 26.02, 20
              PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
              PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            End If
          End If
        End If
      Loop
    End If
  End With
         
  Printer.EndDoc

End Sub


'==============================================
'Boletos quitados
'==============================================
Public Sub BoletosQuitados()
  
  Dim iAtual  As Single
  Dim iPage   As Integer
  Dim iSoma   As Double
  Dim iSub    As Double
  Dim iRef    As Integer
  Dim sIdent  As String
  Dim sSql    As String
  
  iPage = 1
  iSoma = 0
  iSub = 0
  iRef = 1
  sIdent = ""
  
  sSql = "SELECT DISTINCT Boletos.Data, Boletos.VCTO, Boletos.VALR, Boletos.CDSC, Boletos.NOME, Mensalidades.Situacao " & _
        "FROM Boletos LEFT JOIN Mensalidades ON Boletos.EXTR = Mensalidades.Boleto " & _
        "WHERE (((Mensalidades.Situacao)='S')) ORDER BY Boletos.Nome;"
  
  Set rsSelecao = dbfDados.OpenRecordset(sSql, dbOpenDynaset)
  DoEvents
  
  With rsSelecao
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Caption = "Listagem de débitos"
      Iniciar
      Do While Not .EOF
        DoEvents
        'Cabeçalho
        If Dir(sFormataCaminho(App.Path) & "coeducar.jpg") <> "" Then
          PrintDesenho iPage - 1, 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "coeducar.jpg"
        End If
        PrintDireita iPage - 1, 0.6, 3, vbCabecalho.Nome, "Arial", 14, True
        PrintDireita iPage - 1, 1.2, 3, vbCabecalho.Endereco, "Arial", 9, False
        PrintDireita iPage - 1, 1.2, 10, vbCabecalho.Bairro, "Arial", 9, False
        PrintDireita iPage - 1, 1.2, 15, vbCabecalho.Cidade, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 3, vbCabecalho.Estado, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 4, vbCabecalho.Cep, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 6.5, vbCabecalho.Cgc, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 10.5, vbCabecalho.Inscricao, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 15, vbCabecalho.Telefone, "Arial", 9, False
        PrintLinha iPage - 1, 2, 1, 2.03, 20, vbSolid
        PrintDireita iPage - 1, 2.5, 1, "Relatório de bloquetos quitados", "Arial", 10, True
        PrintDireita iPage - 1, 3, 1, "Data", "Arial", 10, True
        PrintDireita iPage - 1, 3, 3, "Ident.", "Arial", 10, True
        PrintDireita iPage - 1, 3, 4.3, "Nome", "Arial", 10, True
        PrintDireita iPage - 1, 3, 14, "Vencimento", "Arial", 10, True
        PrintEsquerda iPage - 1, 3, 20, "Valor", "Arial", 10, True
        PrintLinha iPage - 1, 3.4, 1, 3.42, 20, vbSolid
        iAtual = 3.6
        Do While Not .EOF And iAtual < 25.5
          PrintDireita iPage - 1, iAtual, 1, Format$(!Data, "dd/mm/yyyy"), "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 3, !CDSC & "", "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 4.3, !Nome & "", "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 14, Format$(!VCTO, "dd/mm/yyyy"), "Arial", 9, False
          PrintEsquerda iPage - 1, iAtual, 20, Format$(!Valr, "#,##0.00"), "Arial", 9, False
          iAtual = iAtual + 0.4
          iSoma = iSoma + (0 & !Valr)
          iSub = iSub + (0 & !Valr)
          iRef = iRef + 1
          .MoveNext
          DoEvents
        Loop
        If Not .EOF Then
          PrintLinha iPage - 1, 26, 1, 26.02, 20
          PrintDireita iPage - 1, 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
          PrintEsquerda iPage - 1, 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
          NovaPagina
          iPage = iPage + 1
        Else
          PrintDireita iPage - 1, 25.5, 10, "Total recebido", "Arial", 10, False
          PrintEsquerda iPage - 1, 25.5, 20, Format$(iSoma, "#,##0.00"), "Arial", 10, True
          PrintLinha iPage - 1, 26, 1, 26.02, 20
          PrintDireita iPage - 1, 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
          PrintEsquerda iPage - 1, 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
        End If
      Loop
    End If
  End With
         
  Relatorios.FimDeImpressao
  
  Set rsSelecao = Nothing
  ChamaLpt = 2
  
End Sub

Private Sub BoletosQuitadosLpt(ByVal Page1 As Single, ByVal Page2 As Single)
  
  Dim iAtual  As Single
  Dim iPage   As Integer
  Dim iSoma   As Double
  Dim iSub    As Double
  Dim iRef    As Integer
  Dim sIdent  As String
  Dim sSql    As String
  
  iPage = 1
  iSoma = 0
  iSub = 0
  iRef = 1
  sIdent = ""
  
  sSql = "SELECT DISTINCT Boletos.VCTO, Boletos.VALR, Boletos.CDSC, Boletos.NOME, Boletos.DATA, Boletos.BOLE, Mensalidades.Situacao " & _
        "FROM Boletos RIGHT JOIN Mensalidades ON Boletos.CDSC = Mensalidades.Identificacao " & _
        "Where (((Mensalidades.Situacao) = 'S')) ORDER BY Boletos.NOME;"

  Set rsSelecao = dbfDados.OpenRecordset(sSql, dbOpenDynaset)
  DoEvents
  
  With rsSelecao
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Printer.ScaleMode = vbCentimeters
      Do While Not .EOF
        DoEvents
        If iPage < Page1 Then
          GoTo Nova
        End If
        If iPage > Page2 Then
          GoTo Final
        End If
        'Cabeçalho
        If Dir(sFormataCaminho(App.Path) & "coeducar.jpg") <> "" Then
          PrnDesenho 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "coeducar.jpg"
        End If
        PrnDireita 0.6, 3, vbCabecalho.Nome, "Arial", 14, True
        PrnDireita 1.2, 3, vbCabecalho.Endereco, "Arial", 9, False
        PrnDireita 1.2, 10, vbCabecalho.Bairro, "Arial", 9, False
        PrnDireita 1.2, 15, vbCabecalho.Cidade, "Arial", 9, False
        PrnDireita 1.6, 3, vbCabecalho.Estado, "Arial", 9, False
        PrnDireita 1.6, 4, vbCabecalho.Cep, "Arial", 9, False
        PrnDireita 1.6, 6.5, vbCabecalho.Cgc, "Arial", 9, False
        PrnDireita 1.6, 10.5, vbCabecalho.Inscricao, "Arial", 9, False
        PrnDireita 1.6, 15, vbCabecalho.Telefone, "Arial", 9, False
        PrnLinha 2, 1, 2.03, 20, vbSolid
        PrnDireita 2.5, 1, "Relatório de bloquetos quitados", "Arial", 10, True
        PrnDireita 3, 1, "Data", "Arial", 10, True
        PrnDireita 3, 3, "Identific.", "Arial", 10, True
        PrnDireita 3, 4.3, "Nome", "Arial", 10, True
        PrnDireita 3, 14, "Vencimento", "Arial", 10, True
        PrnEsquerda 3, 20, "Valor", "Arial", 10, True
        PrnLinha 3.4, 1, 3.42, 20, vbSolid
Nova:
        iAtual = 3.6
        Do While Not .EOF And iAtual < 25.5
          If iPage >= Page1 And iPage <= Page2 Then
            PrnDireita iAtual, 1, Format$(!Data, "dd/mm/yyyy"), "Arial", 9, False
            PrnDireita iAtual, 3, !CDSC, "Arial", 9, False
            PrnDireita iAtual, 4.3, !Nome, "Arial", 9, False
            PrnDireita iAtual, 14, Format$(!VCTO, "dd/mm/yyyy"), "Arial", 9, False
            PrnEsquerda iAtual, 20, Format$(!Valr, "#,##0.00"), "Arial", 9, False
          End If
          iAtual = iAtual + 0.4
          iSoma = iSoma + !Valr
          iSub = iSub + !Valr
          iRef = iRef + 1
          .MoveNext
          DoEvents
        Loop
        If Not .EOF Then
          If iPage >= Page1 And iPage <= Page2 Then
            PrnLinha 26, 1, 26.02, 20
            PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
            PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            Printer.NewPage
          End If
          iPage = iPage + 1
        Else
Final:
          If iPage <= Page2 Then
            If iPage = (Max + 1) Then
              PrnDireita 25.5, 10, "Total recebido", "Arial", 10, False
              PrnEsquerda 25.5, 20, Format$(iSoma, "#,##0.00"), "Arial", 10, True
              PrnLinha 26, 1, 26.02, 20
              PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
              PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            Else
              PrnLinha 26, 1, 26.02, 20
              PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
              PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            End If
          End If
        End If
      Loop
    End If
  End With
         
  Printer.EndDoc

End Sub


'=============================================================
'Relatório de pias/responsáveis
'=============================================================
Public Sub PaisResp()
  
  Dim iAtual  As Single
  Dim iPage   As Integer
  Dim sSql    As String
  
  iPage = 1
  
  With tbPais
    .Index = "nomeid"
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Caption = "Listagem de Pais/Responsáveis"
      Iniciar
      Do While Not .EOF
        DoEvents
        'Cabeçalho
        If Dir(sFormataCaminho(App.Path) & "coeducar.jpg") <> "" Then
          PrintDesenho iPage - 1, 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "coeducar.jpg"
        End If
        PrintDireita iPage - 1, 0.6, 3, vbCabecalho.Nome, "Arial", 14, True
        PrintDireita iPage - 1, 1.2, 3, vbCabecalho.Endereco, "Arial", 9, False
        PrintDireita iPage - 1, 1.2, 10, vbCabecalho.Bairro, "Arial", 9, False
        PrintDireita iPage - 1, 1.2, 15, vbCabecalho.Cidade, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 3, vbCabecalho.Estado, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 4, vbCabecalho.Cep, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 6.5, vbCabecalho.Cgc, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 10.5, vbCabecalho.Inscricao, "Arial", 9, False
        PrintDireita iPage - 1, 1.6, 15, vbCabecalho.Telefone, "Arial", 9, False
        PrintLinha iPage - 1, 2.02, 1, 2.04, 20, vbSolid
        PrintDireita iPage - 1, 2.5, 1, "Relatório de Pais/Responsáveis", "Arial", 10, True
        PrintDireita iPage - 1, 3, 1, "Ident.", "Arial", 10, True
        PrintDireita iPage - 1, 3, 3, "Nome", "Arial", 10, True
        PrintDireita iPage - 1, 3, 14, "Fone", "Arial", 10, True
        PrintEsquerda iPage - 1, 3, 18, "Venc.", "Arial", 10, True
        PrintEsquerda iPage - 1, 3, 20, "Valor", "Arial", 10, True
        PrintLinha iPage - 1, 3.5, 1, 3.52, 20, vbSolid
        iAtual = 3.6
        Do While Not .EOF And iAtual < 25.5
          sSql = "Select * From Alunos Where Identificacao = '" & !Identificacao & "' Order By Nome;"
          Set rsSelecao = dbfDados.OpenRecordset(sSql, dbOpenDynaset)
          DoEvents
          PrintDireita iPage - 1, iAtual, 1, !Identificacao & "", "Arial", 9, False
          PrintDireita iPage - 1, iAtual, 3, !Nome & "", "Arial", 9, True
          PrintDireita iPage - 1, iAtual, 14, !Residencial & "", "Arial", 9, False
          PrintEsquerda iPage - 1, iAtual, 18, Format$(!Vencimento, "#0"), "Arial", 9, False
          PrintEsquerda iPage - 1, iAtual, 20, Format$(!Mensal, "#,##0.00"), "Arial", 9, False
          iAtual = iAtual + 0.4
          If rsSelecao.RecordCount > 0 Then
            rsSelecao.MoveLast
            rsSelecao.MoveFirst
            Do While Not rsSelecao.EOF
              PrintDireita iPage - 1, iAtual, 2, rsSelecao!Nome & "", "Arial Narrow", 9, False, True
              PrintDireita iPage - 1, iAtual, 16, Format(rsSelecao!Nascimento, "dd/mm/yyyy"), "Arial Narrow", 9, False, True
              iAtual = iAtual + 0.4
              rsSelecao.MoveNext
              DoEvents
            Loop
          End If
          .MoveNext
          DoEvents
        Loop
        If Not .EOF Then
          PrintLinha iPage - 1, 26, 1, 26.02, 20
          PrintDireita iPage - 1, 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
          PrintEsquerda iPage - 1, 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
          NovaPagina
          iPage = iPage + 1
        Else
          PrintLinha iPage - 1, 26, 1, 26.02, 20
          PrintDireita iPage - 1, 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
          PrintEsquerda iPage - 1, 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
        End If
      Loop
    End If
  End With
         
  FimDeImpressao
  
  Set rsSelecao = Nothing
  
  ChamaLpt = 3
  
End Sub

Private Sub PaisRespLpt(ByVal Page1 As Single, ByVal Page2 As Single)
  
  Dim iAtual  As Single
  Dim iPage   As Integer
  Dim sSql    As String
  
  iPage = 1
  
  With tbPais
    .Index = "nomeid"
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Printer.ScaleMode = vbCentimeters
      Do While Not .EOF
        DoEvents
        If iPage < Page1 Then
          GoTo Nova
        End If
        If iPage > Page2 Then
          GoTo Final
        End If
        'Cabeçalho
        If Dir(sFormataCaminho(App.Path) & "coeducar.jpg") <> "" Then
          PrnDesenho 0.6, 1, 1.53, 1.3, sFormataCaminho(App.Path) & "coeducar.jpg"
        End If
        PrnDireita 0.6, 3, vbCabecalho.Nome, "Arial", 14, True
        PrnDireita 1.2, 3, vbCabecalho.Endereco, "Arial", 9, False
        PrnDireita 1.2, 10, vbCabecalho.Bairro, "Arial", 9, False
        PrnDireita 1.2, 15, vbCabecalho.Cidade, "Arial", 9, False
        PrnDireita 1.6, 3, vbCabecalho.Estado, "Arial", 9, False
        PrnDireita 1.6, 4, vbCabecalho.Cep, "Arial", 9, False
        PrnDireita 1.6, 6.5, vbCabecalho.Cgc, "Arial", 9, False
        PrnDireita 1.6, 10.5, vbCabecalho.Inscricao, "Arial", 9, False
        PrnDireita 1.6, 15, vbCabecalho.Telefone, "Arial", 9, False
        PrnLinha 2.02, 1, 2.04, 20, vbSolid
        PrnDireita 2.5, 1, "Relatório de Pais/Responsáveis", "Arial", 10, True
        PrnDireita 3, 1, "Ident.", "Arial", 10, True
        PrnDireita 3, 3, "Nome", "Arial", 10, True
        PrnDireita 3, 14, "Fone", "Arial", 10, True
        PrnEsquerda 3, 18, "Venc.", "Arial", 10, True
        PrnEsquerda 3, 20, "Valor", "Arial", 10, True
        PrnLinha 3.5, 1, 3.52, 20, vbSolid
Nova:
        iAtual = 3.6
        Do While Not .EOF And iAtual < 25.5
          If iPage >= Page1 And iPage <= Page2 Then
            PrnDireita iAtual, 1, !Identificacao & "", "Arial", 9, False
            PrnDireita iAtual, 3, !Nome & "", "Arial", 9, True
            PrnDireita iAtual, 14, !Residencial & "", "Arial", 9, False
            PrnEsquerda iAtual, 18, Format$(!Vencimento, "#0"), "Arial", 9, False
            PrnEsquerda iAtual, 20, Format$(!Mensal, "#,##0.00"), "Arial", 9, False
          End If
          iAtual = iAtual + 0.4
          sSql = "Select * From Alunos Where Identificacao = '" & !Identificacao & "' Order By Nome;"
          Set rsSelecao = dbfDados.OpenRecordset(sSql, dbOpenDynaset)
          DoEvents
          If rsSelecao.RecordCount > 0 Then
            rsSelecao.MoveLast
            rsSelecao.MoveFirst
            Do While Not rsSelecao.EOF
              If iPage >= Page1 And iPage <= Page2 Then
                PrnDireita iAtual, 2, rsSelecao!Nome & "", "Arial Narrow", 9, False, True
                PrnDireita iAtual, 16, Format(rsSelecao!Nascimento, "dd/mm/yyyy"), "Arial Narrow", 9, False, True
              End If
              iAtual = iAtual + 0.4
              rsSelecao.MoveNext
              DoEvents
            Loop
          End If
          Set rsSelecao = Nothing
          .MoveNext
          DoEvents
        Loop
        If Not .EOF Then
          If iPage >= Page1 And iPage <= Page2 Then
            PrnLinha 26, 1, 26.02, 20
            PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
            PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
            Printer.NewPage
          End If
          iPage = iPage + 1
        Else
Final:
          If iPage <= Page2 Then
            PrnLinha 26, 1, 26.02, 20
            PrnDireita 26.2, 1, "J & F Software's Ltda (31) 3891-4298", "Arial Narrow", 7, False
            PrnEsquerda 26.2, 20, "Página: " & iPage, "Arial Narrow", 7, False
          End If
        End If
      Loop
    End If
  End With
         
  Printer.EndDoc
  
  Set rsSelecao = Nothing

End Sub

