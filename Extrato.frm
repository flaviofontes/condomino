VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form Extrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   Icon            =   "Extrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpApartir 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      ToolTipText     =   "Se informar uma data o relatório será filtrado."
      Top             =   600
      Width           =   1275
      _ExtentX        =   2249
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
   Begin VB.CheckBox chInfo 
      Caption         =   "Informações complementares"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
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
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   6840
      TabIndex        =   5
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Com vencimento apartir de"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inquilino"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   585
   End
End
Attribute VB_Name = "Extrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbAssociados As Recordset
Dim selecao As String

Private Sub cmdLocalizar_Click()
  AchaInquilino 1
End Sub

Private Sub cmdPrint_Click()
  
  If IsDate(cpApartir.Text) Then
    selecao = " and ({BOLETOS.CANCELADO} = 'N' and {BOLETOS.VCTO} >= date(" & Year(cpApartir.Text) & "," & Month(cpApartir.Text) & "," & Day(cpApartir.Text) & "))"
  Else
    selecao = " and ({BOLETOS.CANCELADO} = 'N')"
  End If
  
  If chInfo.Value = 0 Then
    RelatoriosRPT.Carregar "", Parametros.dados, "{BOLETOS.CDSC} = " & cpCodigo.Text & selecao, "Extrato de inquilino: " & NomeCondominio(tbAssociados!Condominio), sFormataCaminho(App.Path) & "EXTRATO.rpt"
  Else
    RelatoriosRPT.Carregar "", Parametros.dados, "{BOLETOS.CDSC} = " & cpCodigo.Text & selecao, "Extrato de inquilino: " & NomeCondominio(tbAssociados!Condominio), sFormataCaminho(App.Path) & "extratocompl.rpt"
  End If
End Sub

Private Sub cpApartir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    AchaInquilino 0
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  KeyPreview = True
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  Refresh
End Sub

Private Sub AchaInquilino(ByVal Tipo As Integer)
  If Val(cpCodigo.Text) > 0 And Tipo = 0 Then
    With tbAssociados
      .Index = "codigoid"
      .Seek "=", cpCodigo.Text
      If Not .NoMatch Then
        cpNome.Text = !Tipo & " " & !Apartamento & " " & !Proprietario
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
          cpCodigo.Text = !Codigo
          cpNome.Text = !Tipo & " " & !Apartamento & " " & !Proprietario
        End If
      End With
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbAssociados = Nothing
End Sub
