VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MostraRel 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6960
   Icon            =   "Mostrarel.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6960
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   870
      Left            =   4755
      TabIndex        =   2
      Top             =   2115
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1535
      _Version        =   393217
      TextRTF         =   $"Mostrarel.frx":0442
   End
   Begin VB.ListBox VarPages 
      Height          =   1815
      Left            =   2700
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   4980
      Top             =   1245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Relatorio 
      Height          =   4275
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   7541
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Mostrarel.frx":04C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MenuInprimir 
      Caption         =   "&Imprimir"
   End
   Begin VB.Menu MenuPrimeira 
      Caption         =   "P&rimeira"
   End
   Begin VB.Menu MenuAnterior 
      Caption         =   "&Anterior"
   End
   Begin VB.Menu MenuProxima 
      Caption         =   "&Próxima"
   End
   Begin VB.Menu MenuUltima 
      Caption         =   "Ulti&ma"
   End
   Begin VB.Menu MenuFechar 
      Caption         =   "&Fechar"
   End
End
Attribute VB_Name = "MostraRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPage As Integer
Dim NewRel As RichTextBox
Public sCap As String

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Refresh
  KeyPreview = True
  iPage = 0
  Set NewRel = RichTextBox1
  Proc_ListaArquivos (App.Path)
  DoEvents
  With Relatorio
    If VarPages.ListCount > 0 Then
      If VarPages.List(iPage) <> "" Then
        .LoadFile (sFormataCaminho(App.Path) & VarPages.List(iPage))
      Else
        .FileName = ""
      End If
    Else
      .FileName = ""
    End If
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
    .Font.Bold = False
    .Refresh
  End With
  ChecaPrint
  MudaCaption
End Sub

Private Sub Form_Resize()
  With Relatorio
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
    .Refresh
  End With
End Sub

Private Sub MenuAnterior_Click()
  If iPage > 0 Then
    iPage = iPage - 1
  End If
  With Relatorio
    If VarPages.ListCount > 0 Then
      If VarPages.List(iPage) <> "" Then
        .LoadFile (sFormataCaminho(App.Path) & VarPages.List(iPage))
      Else
        .FileName = ""
      End If
    Else
      .FileName = ""
    End If
  End With
  MudaCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Kill sFormataCaminho(App.Path) & "rel0010.*"
End Sub

Private Sub MenuFechar_Click()
  Unload Me
End Sub

Private Sub MenuInprimir_Click()
On Error GoTo Errado
  
  Dim iCounter As Integer
  Dim X        As Printer
  Dim iPg1     As Integer
  Dim iPg2     As Integer
  
  With Dialog1
    .DialogTitle = "Selecione a impressora"
    .CancelError = True
    If VarPages.ListCount > 1 Then
      .FromPage = 1
      .ToPage = 1
      .Max = CInt(VarPages.ListCount - 1)
      .Flags = cdlPDDisablePrintToFile Or cdlPDReturnDC Or cdlPDAllPages Or cdlPDUseDevModeCopies Or cdlPDNoSelection
    Else
      .Flags = cdlPDDisablePrintToFile Or cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDUseDevModeCopies Or cdlPDNoSelection
    End If
    .ShowPrinter
    For Each X In Printers
       If X.hDC = .hDC Then
          Set Printer = X
          Exit For
       End If
    Next
    If Not VerificaFlags(.Flags) Then
      Progre.Label1.Caption = "Imprimindo página:"
      Progre.Show
      Progre.Refresh
      For iCounter = 0 To VarPages.ListCount - 1
        Progre.Label1.Caption = "Imprimindo página: " & Str(iCounter + 1)
        Progre.Label1.Refresh
        DoEvents
        If VarPages.List(iCounter) <> "" Then
          NewRel.LoadFile (sFormataCaminho(App.Path) & VarPages.List(iCounter))
        End If
        NewRel.SelPrint (Printer.hDC)
        Printer.EndDoc
        DoEvents
      Next iCounter
      Unload Progre
    Else
      iPg1 = .FromPage - 1
      iPg2 = .ToPage - 1
      If iPg2 > VarPages.ListCount - 1 Then
        iPg2 = VarPages.ListCount - 1
      End If
      Progre.Label1.Caption = "Imprimindo página:"
      Progre.Show
      Progre.Refresh
      For iCounter = iPg1 To iPg2
        Progre.Label1.Caption = "Imprimindo página: " & Str(iCounter + 1)
        Progre.Label1.Refresh
        DoEvents
        If VarPages.List(iCounter) <> "" Then
          NewRel.LoadFile (sFormataCaminho(App.Path) & VarPages.List(iCounter))
        End If
        NewRel.SelPrint (Printer.hDC)
        Printer.EndDoc
        DoEvents
      Next iCounter
      Unload Progre
    End If
  End With


Fim:
  Exit Sub

Errado:
  If Err.Number = 32755 Then
    Resume Fim
  Else
    MsgBox "Erro No. " & Str(Err.Number) & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
    Resume Fim
  End If
  
End Sub

Public Sub Proc_ListaArquivos(ByVal sPasta As String)
'FAZ LISTA DE ARQUIVOS DO DIRETÓRIO ESCOLHIDO


'declara variável
Dim sArquivo As String
sPasta = sFormataCaminho(sPasta)
'pega a primeira entrada...
sArquivo = Dir(sPasta & "rel0010.*", vbArchive)

'começa o Loop enquanto nomes forem encontrados...
Do While sArquivo <> ""
    ' Ignora o diretório...
    If sArquivo <> "." And sArquivo <> ".." Then
        'verifica se é um arquivo
        If (GetAttr(sPasta & sArquivo) And vbArchive) = vbArchive Then
            'acrescenta este arquivo à lista...
            VarPages.AddItem sArquivo
        End If
    End If
    'captura o próximo nome...
    sArquivo = Dir
Loop

End Sub

Private Sub MenuPrimeira_Click()
  iPage = 0
  With Relatorio
    If VarPages.ListCount > 0 Then
      If VarPages.List(iPage) <> "" Then
        .LoadFile (sFormataCaminho(App.Path) & VarPages.List(iPage))
      Else
        .FileName = ""
      End If
    Else
      .FileName = ""
    End If
  End With
  MudaCaption
End Sub

Private Sub MenuProxima_Click()
  If iPage < VarPages.ListCount - 1 Then
    iPage = iPage + 1
  End If
  With Relatorio
    If VarPages.ListCount > 0 Then
      If VarPages.List(iPage) <> "" Then
        .LoadFile (sFormataCaminho(App.Path) & VarPages.List(iPage))
      Else
        .FileName = ""
      End If
    Else
      .FileName = ""
    End If
  End With
  MudaCaption
End Sub

Private Sub MudaCaption()
  Caption = sCap & "  Pagina: " & (iPage + 1) & " de " & VarPages.ListCount
End Sub

Private Sub MenuUltima_Click()
  iPage = VarPages.ListCount - 1
  With Relatorio
    If VarPages.ListCount > 0 Then
      If VarPages.List(iPage) <> "" Then
        .LoadFile (sFormataCaminho(App.Path) & VarPages.List(iPage))
      Else
        .FileName = ""
      End If
    Else
      .FileName = ""
    End If
  End With
  MudaCaption
End Sub

Private Sub ChecaPrint()
On Error Resume Next
  Dim sPrint As String
  
  sPrint = Printer.DeviceName
  
  If sPrint = "" Then
    MenuInprimir.Enabled = False
  Else
    MenuInprimir.Enabled = True
  End If
  
End Sub

Private Function VerificaFlags(ByVal nflags As Long) As Boolean
  Dim bRet As Boolean
  
  bRet = False
  
  If nflags >= cdlPDNoPageNums Then
    nflags = nflags - cdlPDNoPageNums
  End If
  
  If nflags >= cdlPDDisablePrintToFile Then
    nflags = nflags - cdlPDDisablePrintToFile
  End If
  
  If nflags >= cdlPDUseDevModeCopies Then
    nflags = nflags - cdlPDUseDevModeCopies
  End If
  
  If nflags >= cdlPDHelpButton Then
    nflags = nflags - cdlPDHelpButton
  End If
  
  If nflags >= cdlPDReturnDefault Then
    nflags = nflags - cdlPDReturnDefault
  End If
  
  If nflags >= cdlPDReturnDC Then
    nflags = nflags - cdlPDReturnDC
  End If
  
  If nflags >= cdlPDNoWarning Then
    nflags = nflags - cdlPDNoWarning
  End If
  
  If nflags >= cdlPDPrintSetup Then
    nflags = nflags - cdlPDPrintSetup
  End If
  
  If nflags >= cdlPDPrintToFile Then
    nflags = nflags - cdlPDPrintToFile
  End If
  
  If nflags >= cdlPDCollate Then
    nflags = nflags - cdlPDCollate
  End If
  
  If nflags >= cdlPDNoPageNums Then
    nflags = nflags - cdlPDNoPageNums
  End If
  
  If nflags >= cdlPDNoSelection Then
    nflags = nflags - cdlPDNoSelection
  End If
  
  If nflags >= cdlPDPageNums Then
    nflags = nflags - cdlPDPageNums
    bRet = True
  End If
  
  If nflags >= cdlPDSelection Then
    nflags = nflags - cdlPDSelection
  End If
  
  If nflags >= cdlPDAllPages Then
    nflags = nflags - cdlPDAllPages
  End If
  
  VerificaFlags = bRet
End Function
