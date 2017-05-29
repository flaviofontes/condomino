VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form EnviarEmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eviar e-mail"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   Icon            =   "EnviarEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Anexo 
      Left            =   480
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLocalizaInq 
      Caption         =   "..."
      Height          =   315
      Left            =   1950
      TabIndex        =   11
      Top             =   480
      Width           =   435
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   315
      Left            =   6360
      TabIndex        =   9
      Top             =   3900
      Width           =   1035
   End
   Begin VB.CommandButton cmdAnexo 
      Caption         =   "..."
      Height          =   315
      Left            =   6900
      TabIndex        =   8
      Top             =   3420
      Width           =   495
   End
   Begin VB.TextBox cpAnexo 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3420
      Width           =   5775
   End
   Begin VB.TextBox cpMensagem 
      Height          =   2295
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   960
      Width           =   6195
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2430
      TabIndex        =   1
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   795
      _ExtentX        =   1402
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
   Begin rdActiveText.ActiveText cpNomeInq 
      Height          =   315
      Left            =   2430
      TabIndex        =   12
      Top             =   480
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
   Begin rdActiveText.ActiveText cpCodigoInq 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   480
      Width           =   795
      _ExtentX        =   1402
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
      MaxLength       =   9
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label Label5 
      Height          =   795
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   5955
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Para"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Anexo"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   195
      Width           =   855
   End
End
Attribute VB_Name = "EnviarEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancela As Boolean
Dim tbCondominio As Recordset
Dim tbAssociados As Recordset

Private Sub cmdAnexo_Click()
On Error GoTo Erro
  With Anexo
    .CancelError = True
    .DialogTitle = "Anexo"
    .Filter = "Todos os arquivos|*.*"
    .InitDir = sFormataCaminho(PastaSistema(5))
    .ShowOpen
    If .filename <> "" Then
      cpAnexo.Text = .filename
    End If
  End With

Fim:
  Exit Sub

Erro:
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  Else
    cpAnexo.Text = ""
  End If
  Resume Fim
End Sub

Private Sub cmdEnviar_Click()
On Error GoTo Erro

  Dim sFile As String
  Dim i As Integer
  Dim sNome As String
  Dim sEmail As String
  Dim tPags As Long
  Dim nEnviados As Integer
  Dim rs As Recordset
  
  If Trim(cpMensagem.Text) = "" Then
    MsgBox "O campo mensagem é obrigatório.", vbInformation + vbOKOnly, "Aviso"
    cpMensagem.SetFocus
    Exit Sub
  End If
  
  If Val(cpCodigo.Text) = 0 And Val(cpCodigoInq.Text) = 0 Then
    MsgBox "O condomínio é necessário.", vbInformation + vbOKOnly, "Aviso"
    cpMensagem.SetFocus
    Exit Sub
  End If
  
  Cancela = False
  
  If Val(cpCodigoInq.Text) > 0 Then
    Set rs = db.OpenRecordset("select codigo, nome, email from associados where codigo = " & cpCodigoInq.Text & ";", dbOpenDynaset)
  Else
    Set rs = db.OpenRecordset("select codigo, nome, email from associados where condominio = " & cpCodigo.Text & " order by nome;", dbOpenDynaset)
  End If
  
  Progre.Label1.Caption = "Enviando e-mail... Tecle ESC para cancelar."
  Progre.Show
  Call SetTopMostWindow(Progre.hWnd, True)
  Progre.Refresh
    
  
  nEnviados = 0
  
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If Cancela = True Then
          GoTo 10
        End If
        sNome = AcertaLetras(NomeCompleto(rs!Codigo))
        sEmail = rs!email
        Progre.Label1.Caption = "Enviando e-mail... Tecle ESC para cancelar." & vbCrLf & sNome
        Progre.Refresh
        
        If Trim(sEmail) <> "" Then
          Progre.Label1.Caption = "Enviando e-mail... Tecle ESC para cancelar." & vbCrLf & sNome & vbCrLf & sEmail
          Progre.Refresh
          
          sFile = cpAnexo.Text
          
          If Cancela = True Then
            GoTo 10
          End If
          
          If EnviarOutrosEmail(sEmail, sFile, cpMensagem.Text) Then
            nEnviados = nEnviados + 1
          End If
          
        End If
        DoEvents
        rs.MoveNext
      Loop
    End If
  End With
  
10  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  MsgBox nEnviados & " boleto(s) enviado(s) com sucesso!", vbExclamation + vbOKOnly, "E-mail"
  
Fim:
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  Exit Sub

Erro:
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
  If Err.Number <> 32755 Then
    MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  End If
  Resume Fim

End Sub

Private Sub cmdLocalizaInq_Click()
  AchaInquilino 1
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

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodigo.Text
        If Not .NoMatch Then
          cpNome.Text = !Nome
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
      End If
    End If
  End If
End Sub

Private Sub AchaInquilino(ByVal Tipo As Integer)
  If Val(cpCodigoInq.Text) > 0 And Tipo = 0 Then
    With tbAssociados
      .Index = "codigoid"
      .Seek "=", cpCodigoInq.Text
      If Not .NoMatch Then
        cpNomeInq.Text = !Tipo & " " & !Apartamento & " " & !Proprietario
      Else
        MsgBox "Inquilino não encontrado.", vbInformation + vbOKOnly, "Localizar"
      End If
    End With
  Else
    RetCodigo = 0
    lAssociado.Show 1
    If RetCodigo > 0 Then
      With tbAssociados
        .Index = "codigoid"
        .Seek "=", RetCodigo
        If Not .NoMatch Then
          cpCodigoInq.Text = !Codigo
          cpNomeInq.Text = !Tipo & " " & !Apartamento & " " & !Proprietario
        End If
      End With
    End If
  End If
End Sub

Private Sub cpCodigoInq_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    AchaInquilino 0
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    Cancela = True
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Label5.Caption = "Se for selecionado o condomínio a mensagem e o anexo serão enviados para todos os inquilinos. Se for escolhido um inquilino serão enviados somente para ele." & vbCrLf & "A preferência é o inquilino, independente de se escolher um condomínio."
  Cancela = False
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
  Set tbAssociados = Nothing
End Sub
