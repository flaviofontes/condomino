VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form Descontos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento de descontos"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   Icon            =   "Descontos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   13
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpHistorico 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   1020
      Width           =   8175
      _ExtentX        =   14420
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
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   600
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
      MaxLength       =   7
      TextMask        =   9
      RawText         =   9
      Mask            =   "##/####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   180
      Width           =   6375
      _ExtentX        =   11245
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
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.CommandButton cmdLocInquilino 
      Caption         =   "..."
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   180
      Width           =   675
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   180
      Width           =   1035
      _ExtentX        =   1826
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mês/ano"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inquilino"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "Descontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tbDescontos As Recordset
Dim tbAssociados As Recordset
Dim deleta As Boolean

Private Sub cmdExcluir_Click()
  
  If Not deleta Then
    MsgBox "Nada para excluir.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If tbDescontos!Chave & "" <> "" Then
    Resp = MsgBox("Este desconto foi gerado de forma automática pelo sistema. Excluir assim mesmo?", vbQuestion + vbYesNo, "Excluir")
    If Resp = vbNo Then
      Exit Sub
    End If
  End If
  
  Resp = MsgBox("Confirma a exclusão deste desconto?", vbQuestion + vbYesNo, "Excluir")
  If Resp = vbYes Then
    tbDescontos.Delete
    Limpar
  End If
  
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  LocDesconto.Show 1
  deleta = False
  If RetCodigo > 0 Then
    With tbDescontos
      .Index = "id_desconto"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpMes.Text = !mes
        cpHistorico.Text = !Historico
        cpValor.Text = !valor
        cpCodigo.Text = !id_inquilino
        AchaInquilino cpCodigo.Text
        deleta = True
      End If
    End With
  End If
End Sub

Private Sub cmdLocInquilino_Click()
  AchaInquilino 0
End Sub

Private Sub cmdSalvar_Click()
  If Val(cpCodigo.Text) < 1 Then
    MsgBox "Informe o inquilino.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  If cpValor.Text <= 0# Then
    MsgBox "O valor informado não é válido.", vbInformation + vbOKOnly, "Aviso"
    cpValor.SetFocus
    Exit Sub
  End If
  If cpHistorico.Text = "" Then
    MsgBox "Preencha o histórico.", vbInformation + vbOKOnly, "Aviso"
    cpHistorico.SetFocus
    Exit Sub
  End If
  If Not IsDate("25/" & cpMes.Text) Then
    MsgBox "Mês/ano informado não é válido.", vbInformation + vbOKOnly, "Aviso"
    cpHistorico.SetFocus
    Exit Sub
  Else
    Dim s As Integer
    s = DateDiff("m", Date, CDate(Day(Date) & "/" & cpMes.Text))
    If s < 0 Then
      Resp = MsgBox("Mês/ano informado é " & (s * -1) & " meses inferior ao mês/ano atual. Gravar assim mesmo?", vbInformation + vbYesNo, "Aviso")
      If Resp = vbNo Then
        cpHistorico.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  With tbDescontos
    If Not deleta Then
      .AddNew
    Else
      .Edit
    End If
    !mes = cpMes.Text
    !Historico = cpHistorico.Text
    !valor = cpValor.Text
    !id_inquilino = cpCodigo.Text
    !id_condominio = CondominioInquilino(cpCodigo.Text)
    .Update
  End With
  
  MsgBox "Dados gravados com sucesso.", vbInformation + vbOKOnly, "Aviso"
  Limpar
End Sub

Private Sub Limpar()
  cpMes.Text = ""
  cpHistorico.Text = ""
  cpValor.Text = 0
  cpCodigo.Text = 0
  cpNome.Text = ""
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    AchaInquilino cpCodigo.Text
    cpMes.SetFocus
  End If
End Sub

Private Sub cpHistorico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  End If
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpHistorico.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSalvar.SetFocus
  End If
End Sub

Private Sub AchaInquilino(ByVal nCod As Long)
  If nCod > 0 Then
    With tbAssociados
      .Index = "codigoid"
      .Seek "=", nCod
      If Not .NoMatch Then
        cpCodigo.Text = nCod
        cpNome.Text = !Proprietario
        cpMes.SetFocus
      Else
        MsgBox "Código não encontrado!", vbInformation + vbOKOnly, "Aviso"
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
          cpCodigo.Text = RetCodigo
          cpNome.Text = !Proprietario
          cpMes.SetFocus
        End If
      End With
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  Set tbDescontos = db.OpenRecordset("descontos", dbOpenTable)
  deleta = False
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbDescontos = Nothing
  Set tbAssociados = Nothing
End Sub
