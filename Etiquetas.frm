VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Etiquetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de etiquetas"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   ControlBox      =   0   'False
   Icon            =   "Etiquetas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OpCidade 
      Caption         =   "Para &sócio fora da cidade sede"
      Height          =   195
      Left            =   3435
      TabIndex        =   22
      Top             =   2790
      Width           =   2520
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   210
      Left            =   45
      TabIndex        =   20
      Top             =   4035
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Fechar"
      Height          =   795
      Left            =   4995
      Picture         =   "Etiquetas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1185
      Width           =   1260
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   795
      Left            =   4995
      Picture         =   "Etiquetas.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   390
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   165
      TabIndex        =   15
      Top             =   0
      Width           =   4410
      Begin VB.OptionButton OpDuas 
         Alignment       =   1  'Right Justify
         Caption         =   "ou '2' por formulário"
         Height          =   195
         Left            =   1830
         TabIndex        =   17
         Top             =   210
         Width           =   1755
      End
      Begin VB.OptionButton OpUma 
         Alignment       =   1  'Right Justify
         Caption         =   "Etiquetas com '1'"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox cpGeralF 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1650
      Width           =   2025
   End
   Begin VB.TextBox cpGeralI 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1335
      Width           =   2025
   End
   Begin VB.ComboBox cpFim 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3465
      Width           =   4110
   End
   Begin VB.ComboBox cpInicio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3105
      Width           =   4110
   End
   Begin VB.OptionButton OpBloquetos 
      Caption         =   "Para os &bloquetos bancários"
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   2775
      Width           =   2325
   End
   Begin VB.TextBox cpMes 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Top             =   2340
      Width           =   525
   End
   Begin VB.OptionButton OpAniversario 
      Caption         =   "Para os Aniversariantes do Mês"
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   2055
      Width           =   1425
   End
   Begin VB.ComboBox cpTipo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   2985
   End
   Begin VB.OptionButton OpGeral 
      Caption         =   "Geral..."
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   630
      Value           =   -1  'True
      Width           =   840
   End
   Begin VB.Label Porcento 
      AutoSize        =   -1  'True
      Caption         =   "Progresso"
      Height          =   195
      Left            =   45
      TabIndex        =   21
      Top             =   3810
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Terminar em"
      Height          =   195
      Left            =   405
      TabIndex        =   12
      Top             =   1710
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Iniciar em"
      Height          =   195
      Left            =   405
      TabIndex        =   11
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Terminar em"
      Height          =   195
      Left            =   405
      TabIndex        =   10
      Top             =   3540
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Iniciar em"
      Height          =   195
      Left            =   405
      TabIndex        =   9
      Top             =   3165
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mês de"
      Height          =   195
      Left            =   405
      TabIndex        =   5
      Top             =   2385
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo do sócio"
      Height          =   195
      Left            =   405
      TabIndex        =   2
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private SelCad As Recordset
Private PrnSim As Boolean

Private Sub PrnEtiquetas()
On Error GoTo Errado
  Dim nFree As Integer
  Dim i As Integer
  If cpInicio.ListCount = 0 Then
    MsgBox "Você deve primeiro gerar os lançamentos.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  If cpInicio.Text = "" Then
    MsgBox "Por favor selecione o nome de onde devo iniciar.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  If cpFim.Text = "" Then
    MsgBox "Por favor selecione o nome onde devo terminar.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With tbBoletos
    .Index = "NOMEID"
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        If !Nome >= cpInicio.Text And !Nome <= cpFim.Text Then
          For i = 0 To 15
            DoEvents
          Next i
          If PrnSim = False Then
            If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
              GoTo Fim
            Else
              PrnSim = True
            End If
          End If
          Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); !NumeroDaCota; Chr(27) & Chr(70)
          Print #nFree, !Nome
          Print #nFree, !Endereco
          Print #nFree, !Bairro
          Print #nFree, !Cidade; Spc(2); !Estado; Spc(2); !Cep
          Print #nFree, ""
        End If
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
        End If
        DoEvents
        .MoveNext
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub

Private Sub PrnEtiqDupla()
On Error GoTo Errado
  Dim nFree As Integer
  Dim Info(1 To 7, 1 To 2) As String
  Dim i As Integer
  If cpInicio.ListCount = 0 Then
    MsgBox "Você deve primeiro gerar os lançamentos.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  If cpInicio.Text = "" Then
    MsgBox "Por favor selecione o nome de onde devo iniciar.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  If cpFim.Text = "" Then
    MsgBox "Por favor selecione o nome onde devo terminar.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With tbBoletos
    .Index = "NOMEID"
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0); Chr(18)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        If !Nome >= cpInicio.Text And !Nome <= cpFim.Text Then
          For i = 0 To 15
            DoEvents
          Next i
          If PrnSim = False Then
            If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
              GoTo Fim
            Else
              PrnSim = True
            End If
          End If
          If Not .EOF Then
            Info(1, 1) = !NumeroDaCota
            Info(2, 1) = !Nome
            Info(3, 1) = !Endereco
            Info(4, 1) = !Bairro
            Info(5, 1) = !Cidade
            Info(6, 1) = !Estado
            Info(7, 1) = !Cep
          End If
          .MoveNext
          If Not .EOF Then
            Info(1, 2) = !NumeroDaCota
            Info(2, 2) = !Nome
            Info(3, 2) = !Endereco
            Info(4, 2) = !Bairro
            Info(5, 2) = !Cidade
            Info(6, 2) = !Estado
            Info(7, 2) = !Cep
          End If
          Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); Info(1, 1); Tab(68); Info(1, 2); Chr(27) & Chr(70)
          Print #nFree, Info(2, 1); Tab(43); Info(2, 2)
          Print #nFree, Info(3, 1); Tab(43); Info(3, 2)
          Print #nFree, Info(4, 1); Tab(43); Info(4, 2)
          Print #nFree, Info(5, 1); Spc(2); Info(6, 1); Spc(2); Info(7, 1); Tab(43); Info(5, 2); Spc(2); Info(6, 2); Spc(2); Info(7, 2)
          Print #nFree, ""
        End If
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
          .MoveNext
        End If
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox Err.Number & vbCrLf & vbCrLf & Err.Description
  Resume Fim

End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  Dim nTipo   As Long
  Dim SelForm As String
  Dim sLetra  As String
  PrnSim = True
  If OpGeral.Value Then
    If cpTipo.Text = "TODOS" Then
      If cpGeralI.Text = "" Then
        MsgBox "Preencha o campo iniciar em.", vbCritical + vbOKOnly, Caption
      Else
        If Len(cpGeralI.Text) = 1 Then
          If cpGeralF.Text = "" Then
            MsgBox "Preencha o campo terminar em.", vbCritical + vbOKOnly, Caption
          Else
            sLetra = cpGeralF.Text & cpGeralF.Text
            SelForm = "Select * from Cadstro Where Nome Like '[" & cpGeralI.Text & "-" & sLetra & "]*' Order By Nome"
          End If
        Else
          SelForm = "Select * From Cadastro Where Nome Like '" & cpGeralI.Text & "*' Order By Nome"
        End If
      End If
    Else
      nTipo = AchaTipo(cpTipo.Text)
      If cpGeralI.Text = "" Then
        MsgBox "Preencha o campo iniciar em.", vbCritical + vbOKOnly, Caption
      Else
        If Len(cpGeralI.Text) = 1 Then
          If cpGeralF.Text = "" Then
            MsgBox "Preencha o campo terminar em.", vbCritical + vbOKOnly, Caption
          Else
            sLetra = cpGeralF.Text & cpGeralF.Text
            SelForm = "Select * from Cadstro Where Nome Like '[" & cpGeralI.Text & "-" & sLetra & "]*' And Tipo = " & nTipo & " Order By Nome"
          End If
        Else
          SelForm = "Select * From Cadastro Where Nome Like '" & cpGeralI.Text & "*' And Tipo = " & nTipo & " Order By Nome"
        End If
      End If
    End If
    Set SelCad = db.OpenRecordset(SelForm, dbOpenDynaset)
    DoEvents
    If OpUma.Value Then
      PrnGeral SelCad
    Else
      PrnGeralDupla SelCad
    End If
  ElseIf OpAniversario.Value Then
    If OpUma.Value Then
      PrnNascimento
    Else
      PrnNascDupla
    End If
  ElseIf OpBloquetos.Value Then
    If OpUma.Value Then
      PrnEtiquetas
    Else
      PrnEtiqDupla
    End If
  ElseIf OpCidade.Value Then
    If OpUma.Value Then
      PrnCidadeUma
    Else
      PrnCidadeDuas
    End If
  End If
End Sub

Private Sub cpFim_Click()
  cmdPrint.SetFocus
End Sub

Private Sub cpGeralF_GotFocus()
  cpGeralF.SelStart = 0
  cpGeralF.SelLength = Len(cpGeralF.Text)
End Sub

Private Sub cpGeralF_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpGeralI_Change()
  If Len(cpGeralI.Text) > 1 Then
    cpGeralF.Locked = True
  Else
    cpGeralF.Locked = False
  End If
End Sub

Private Sub cpGeralI_GotFocus()
  cpGeralI.SelStart = 0
  cpGeralI.SelLength = Len(cpGeralI.Text)
End Sub

Private Sub cpGeralI_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpGeralF.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpInicio_Click()
  Dim II As Integer
  For II = 0 To cpFim.ListCount - 1
    If cpFim.List(II) = cpInicio.Text Then
      cpFim.ListIndex = II
      Exit For
    End If
  Next II
End Sub

Private Sub cpMes_GotFocus()
  cpMes.SelStart = 0
  cpMes.SelLength = Len(cpMes.Text)
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpTipo_Click()
  cpGeralI.SetFocus
End Sub

Private Sub Form_Activate()
  Dim II As Integer
  With cpTipo
    For II = 0 To .ListCount - 1
      If .List(II) = "TODOS" Then
        .ListIndex = II
        Exit For
      End If
    Next II
  End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    If PrnSim = False Then
      KeyAscii = 0
      Unload Me
    Else
      PrnSim = False
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  KeyPreview = True
  Refresh
  DoEvents
'  With tbTipo
'    If .RecordCount > 0 Then
'      .MoveFirst
'      While Not .EOF
'        cpTipo.AddItem (!Descricao)
'        .MoveNext
'      Wend
'    End If
'  End With
  cpTipo.AddItem ("TODOS")
  With tbBoletos
    If .RecordCount > 0 Then
      .MoveFirst
      While Not .EOF
        cpInicio.AddItem (!Nome)
        cpFim.AddItem (!Nome)
        .MoveNext
      Wend
    End If
  End With
End Sub

Private Sub OpAniversario_Click()
  cpMes.SetFocus
End Sub

Private Sub OpBloquetos_Click()
  cpInicio.SetFocus
End Sub

Private Sub OpGeral_Click()
  cpTipo.SetFocus
End Sub

Private Function AchaTipo(ByVal sTipo As String) As Long
'  With daTipo
'    .Index = "descricaoID"
'    .Seek "=", sTipo
'    If Not .NoMatch Then
'      AchaTipo = !Codigo
'    Else
'      AchaTipo = 0
'    End If
'  End With
End Function

Private Sub PrnNascimento()
On Error GoTo Errado
  Dim nFree As Integer
  Dim i As Integer
  If cpMes.Text = "" Then
    MsgBox "Você deve primeiro informar o mês do aniversário.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With tbAssociados
    .Index = "pernome"
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        If Month(!DataDeNascimento) = Val(cpMes.Text) Then
          For i = 0 To 15
            DoEvents
          Next i
          If PrnSim = False Then
            If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
              GoTo Fim
            Else
              PrnSim = True
            End If
          End If
          Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); !NumeroDaCota; Chr(27) & Chr(70)
          Print #nFree, !Nome
          Print #nFree, !Endereco
          Print #nFree, !Bairro
          Print #nFree, !Cidade; Spc(2); !Estado; Spc(2); !Cep
          Print #nFree, ""
        End If
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
        End If
        .MoveNext
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub

Private Sub PrnNascDupla()
On Error GoTo Errado
  Dim nFree As Integer
  Dim Info(1 To 7, 1 To 2) As String
  Dim i As Integer
  If cpMes.Text = "" Then
    MsgBox "Você deve primeiro informar o mês do aniversário.", vbCritical + vbOKOnly, Caption
    GoTo Fim
  End If
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With tbAssociados
    .Index = "pernome"
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0); Chr(18)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        If Month(!DataDeNascimento) = Val(cpMes.Text) Then
          For i = 0 To 15
            DoEvents
          Next i
          If PrnSim = False Then
            If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
              GoTo Fim
            Else
              PrnSim = True
            End If
          End If
          If Not .EOF Then
            Info(1, 1) = !NumeroDaCota
            Info(2, 1) = !Nome
            Info(3, 1) = !Endereco
            Info(4, 1) = !Bairro
            Info(5, 1) = !Cidade
            Info(6, 1) = !Estado
            Info(7, 1) = !Cep
          End If
          .MoveNext
          If Not .EOF Then
            Info(1, 2) = !NumeroDaCota
            Info(2, 2) = !Nome
            Info(3, 2) = !Endereco
            Info(4, 2) = !Bairro
            Info(5, 2) = !Cidade
            Info(6, 2) = !Estado
            Info(7, 2) = !Cep
          End If
          Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); Info(1, 1); Tab(68); Info(1, 2); Chr(27) & Chr(70)
          Print #nFree, Info(2, 1); Tab(43); Info(2, 2)
          Print #nFree, Info(3, 1); Tab(43); Info(3, 2)
          Print #nFree, Info(4, 1); Tab(43); Info(4, 2)
          Print #nFree, Info(5, 1); Spc(2); Info(6, 1); Spc(2); Info(7, 1); Tab(43); Info(5, 2); Spc(2); Info(6, 2); Spc(2); Info(7, 2)
          Print #nFree, ""
        End If
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
          .MoveNext
        End If
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub

Private Sub PrnGeral(ByRef rs_dados As Recordset)
On Error GoTo Errado
  Dim nFree As Integer
  Dim i As Integer
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With rs_dados
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        For i = 0 To 15
          DoEvents
        Next i
        If PrnSim = False Then
          If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
            GoTo Fim
          Else
            PrnSim = True
          End If
        End If
        Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); !NumeroDaCota; Chr(27) & Chr(70)
        Print #nFree, !Nome
        Print #nFree, !Endereco
        Print #nFree, !Bairro
        Print #nFree, !Cidade; Spc(2); !Estado; Spc(2); !Cep
        Print #nFree, ""
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
        End If
        .MoveNext
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub

Private Sub PrnGeralDupla(ByRef rs_dados As Recordset)
On Error GoTo Errado
  Dim nFree As Integer
  Dim Info(1 To 7, 1 To 2) As String
  Dim i As Integer
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With rs_dados
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0); Chr(18)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        For i = 0 To 15
          DoEvents
        Next i
        If PrnSim = False Then
          If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
            GoTo Fim
          Else
            PrnSim = True
          End If
        End If
        If Not .EOF Then
          Info(1, 1) = !NumeroDaCota
          Info(2, 1) = !Nome
          Info(3, 1) = !Endereco
          Info(4, 1) = !Bairro
          Info(5, 1) = !Cidade
          Info(6, 1) = !Estado
          Info(7, 1) = !Cep
        End If
        .MoveNext
        If Not .EOF Then
          Info(1, 2) = !NumeroDaCota
          Info(2, 2) = !Nome
          Info(3, 2) = !Endereco
          Info(4, 2) = !Bairro
          Info(5, 2) = !Cidade
          Info(6, 2) = !Estado
          Info(7, 2) = !Cep
        End If
        Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); Info(1, 1); Tab(68); Info(1, 2); Chr(27) & Chr(70)
        Print #nFree, Info(2, 1); Tab(43); Info(2, 2)
        Print #nFree, Info(3, 1); Tab(43); Info(3, 2)
        Print #nFree, Info(4, 1); Tab(43); Info(4, 2)
        Print #nFree, Info(5, 1); Spc(2); Info(6, 1); Spc(2); Info(7, 1); Tab(43); Info(5, 2); Spc(2); Info(6, 2); Spc(2); Info(7, 2)
        Print #nFree, ""
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
          .MoveNext
        End If
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub

''''''''''
Private Sub PrnCidadeUma()
On Error GoTo Errado
  Dim nFree As Integer
  Dim i As Integer
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With tbAssociados
    .Index = "Pernome"
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        If Not IsNull(!Postagem) And !Postagem > 0 Then
          For i = 0 To 15
            DoEvents
          Next i
          If PrnSim = False Then
            If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
              GoTo Fim
            Else
              PrnSim = True
            End If
          End If
          Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); !NumeroDaCota; Chr(27) & Chr(70)
          Print #nFree, !Nome
          Print #nFree, !Endereco
          Print #nFree, !Bairro
          Print #nFree, !Cidade; Spc(2); !Estado; Spc(2); !Cep
          Print #nFree, ""
        End If
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
        End If
        .MoveNext
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub

Private Sub PrnCidadeDuas()
On Error GoTo Errado
  Dim nFree As Integer
  Dim Info(1 To 7, 1 To 2) As String
  Dim i As Integer
  Porcento.Caption = "Imprimindo... 0%"
  Porcento.Refresh
  nFree = FreeFile
  With tbAssociados
    .Index = "PerNome"
    If .RecordCount > 0 Then
      Open Printer.Port For Output As #nFree
      Print #nFree, Chr(27) & Chr(77); Chr(27) & Chr(120) & Chr(0); Chr(18)
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      Print #nFree, ""
      .MoveFirst
      Do While Not .EOF
        If Not IsNull(!Postagem) And !Postagem > 0 Then
          For i = 0 To 15
            DoEvents
          Next i
          If PrnSim = False Then
            If MsgBox("Cancelar a impressão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes Then
              GoTo Fim
            Else
              PrnSim = True
            End If
          End If
          If Not .EOF Then
            Info(1, 1) = !NumeroDaCota
            Info(2, 1) = !Nome
            Info(3, 1) = !Endereco
            Info(4, 1) = !Bairro
            Info(5, 1) = !Cidade
            Info(6, 1) = !Estado
            Info(7, 1) = !Cep
          End If
Dinovo:
          .MoveNext
          If Not .EOF Then
            If (Left(!Cidade, 6) = "VIÇOSA") Or (Left(!Cidade, 6) = "VICOSA") Then
              GoTo Dinovo
            Else
              If Not .EOF Then
                Info(1, 2) = !NumeroDaCota
                Info(2, 2) = !Nome
                Info(3, 2) = !Endereco
                Info(4, 2) = !Bairro
                Info(5, 2) = !Cidade
                Info(6, 2) = !Estado
                Info(7, 2) = !Cep
              End If
            End If
          End If
          Print #nFree, Chr(18); Tab(25); Chr(27) & Chr(69); Info(1, 1); Tab(68); Info(1, 2); Chr(27) & Chr(70)
          Print #nFree, Info(2, 1); Tab(43); Info(2, 2)
          Print #nFree, Info(3, 1); Tab(43); Info(3, 2)
          Print #nFree, Info(4, 1); Tab(43); Info(4, 2)
          Print #nFree, Info(5, 1); Spc(2); Info(6, 1); Spc(2); Info(7, 1); Tab(43); Info(5, 2); Spc(2); Info(6, 2); Spc(2); Info(7, 2)
          Print #nFree, ""
        End If
        If Not .EOF Then
          Porcento.Caption = "Imprimindo... " & Format(.PercentPosition, "#0\%")
          Porcento.Refresh
          Barra.Value = .PercentPosition
          .MoveNext
        End If
      Loop
    End If
  End With

Fim:
  Barra.Value = 0
  Porcento.Caption = "Progresso"
  Porcento.Refresh
  Close #nFree
  Exit Sub

Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Caption
  Resume Fim

End Sub
