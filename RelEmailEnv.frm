VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelEmailEnv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de e-mails enviados"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "RelEmailEnv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpFim 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
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
   Begin rdActiveText.ActiveText cpInicio 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   480
      Width           =   1155
      _ExtentX        =   2037
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
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   6060
      TabIndex        =   4
      Top             =   480
      Width           =   1140
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   8
      Top             =   60
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
      TabIndex        =   0
      Top             =   60
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   540
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No período de"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1035
   End
End
Attribute VB_Name = "RelEmailEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbCondominio As Recordset

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
  Dim sOrdem As String
  Dim filtro As String
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(cpInicio.Text) Then
    MsgBox "A data inicial não foi informada ou não é válida.", vbInformation + vbOKOnly, "Aviso"
    cpInicio.SetFocus
    Exit Sub
  End If
    
  If Not IsDate(cpFim.Text) Then
    MsgBox "A data final não foi informada ou não é válida.", vbInformation + vbOKOnly, "Aviso"
    cpFim.SetFocus
    Exit Sub
  End If
    
  If CDate(cpInicio.Text) > CDate(cpFim.Text) Then
    MsgBox "A data final não pode ser menor que a data inicial.", vbInformation + vbOKOnly, "Aviso"
    cpFim.SetFocus
    Exit Sub
  End If
  
  filtro = "{EMAIL_ENVIADOS.DATA} >= Date(" & Year(cpInicio.Text) & ", " & Month(cpInicio.Text) & ", " & Day(cpInicio.Text) & ") "
  filtro = filtro & " AND {EMAIL_ENVIADOS.DATA} <= Date(" & Year(cpFim.Text) & ", " & Month(cpFim.Text) & ", " & Day(cpFim.Text) & ") "
  filtro = filtro & " AND {EMAIL_ENVIADOS.ID_CONDOMINIO} = " & cpCodigo.Text
  sOrdem = "E-mails enviados entre " & cpInicio.Text & " e " & cpFim.Text & " do condomínio " & cpNome.Text
  RelatoriosRPT.Carregar "", Parametros.dados, filtro, sOrdem, sFormataCaminho(App.Path) & "email_enviados.rpt"
  
End Sub

Private Sub cpCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpInicio.SetFocus
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
          cpInicio.SetFocus
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
        cpInicio.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpFim_GotFocus()
  If IsDate(cpInicio.Text) Then
    cpFim.Text = UltimoDia(Month(cpInicio.Text) & "/" & Year(cpInicio.Text))
  End If
End Sub

Private Sub cpFim_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdPrint.Enabled Then
      cmdPrint.SetFocus
    End If
  End If
End Sub

Private Sub cpInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFim.SetFocus
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
  Refresh
  KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
