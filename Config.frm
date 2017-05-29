VERSION 5.00
Begin VB.Form Configuracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banco de dados do sistema"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ControlBox      =   0   'False
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPadrao 
      Caption         =   "&Padrão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4245
      TabIndex        =   5
      Top             =   1095
      Width           =   1400
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "C&ancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2400
      TabIndex        =   4
      Top             =   1095
      Width           =   1400
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   585
      TabIndex        =   3
      Top             =   1095
      Width           =   1400
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5625
      TabIndex        =   2
      Top             =   450
      Width           =   615
   End
   Begin VB.TextBox TxtLocalDados 
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
      Left            =   105
      TabIndex        =   0
      Top             =   450
      Width           =   5415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Dados"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Base de Dados"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "Configuracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdConfirma_Click()
On Error GoTo Errado

  Dim RetDados As Boolean
  
  Resp = MsgBox("Esta alteração só será realizada apos reiniciar o programa! Prosseguir?", vbQuestion + vbYesNo, "Aviso")
  If Resp = vbYes Then
    If TxtLocalDados.Text = "" Then
      Resp = MsgBox("O local do banco de dados não é válido!", vbCritical + vbYesNo, "Aviso")
    Else
      If TxtLocalDados.Text <> Parametros.Dados Then
        EscreveIni "Config", "DADOS", TxtLocalDados.Text, Inifile
      End If
    End If
    Unload Me
  End If

Fim:
  Exit Sub
  
Errado:
  MsgBox "Erro no. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim
  
End Sub

Private Sub cmdLocalizar_Click()
    
  Dim sPath   As String
  Dim sDados  As String
  
  sPath = MinhaDll.sProcuraPorDiretorio("Local do banco de dados.", Me)
  
  If Len(sPath) > 0 Then
    sPath = sFormataCaminho(sPath)
    If Dir$(sPath & "dados.mdb") = "" Then
      MsgBox "O local informado não possui um banco de dados válido!", vbExclamation + vbOKOnly, Caption
    Else
      TxtLocalDados.Text = sPath & "dados.mdb"
    End If
  End If

End Sub

Private Sub cmdPadrao_Click()
  TxtLocalDados.Text = sFormataCaminho(App.Path) & "dados.mdb"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Refresh
  TxtLocalDados.Text = Parametros.Dados
  KeyPreview = True
End Sub
