VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form InfoBoleto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informações do boletos"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "InfoBoleto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cpStatus 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   5595
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "InfoBoleto.frx":000C
      Height          =   1875
      Left            =   120
      OleObjectBlob   =   "InfoBoleto.frx":0020
      TabIndex        =   12
      Top             =   2340
      Width           =   7155
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\VBOXSVR\programacao\Porto real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdLocaliza 
      Caption         =   "..."
      Height          =   315
      Left            =   4020
      TabIndex        =   10
      Top             =   480
      Width           =   555
   End
   Begin VB.TextBox cpBoleto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1680
      MaxLength       =   17
      TabIndex        =   0
      Top             =   480
      Width           =   2235
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1665
      TabIndex        =   4
      Top             =   1560
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   12
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpVencimento 
      Height          =   315
      Left            =   1665
      TabIndex        =   3
      Top             =   1200
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      FontName        =   "Arial"
      FontSize        =   9,75
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      Top             =   825
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextCase        =   1
      RawText         =   0
      FontName        =   "Arial"
      FontSize        =   9,75
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   1665
      TabIndex        =   1
      Top             =   825
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextMask        =   3
      RawText         =   3
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpNomeCond 
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   1140
      TabIndex        =   13
      Top             =   1980
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   735
      TabIndex        =   7
      Top             =   1260
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sacado"
      Height          =   195
      Left            =   1020
      TabIndex        =   6
      Top             =   900
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do boleto"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Top             =   555
      Width           =   1260
   End
End
Attribute VB_Name = "InfoBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vCodigo As Long
Dim i As Integer
Dim rsConta As Recordset


Private Sub AchaBoleto()
  Set rsConta = db.OpenRecordset("select * from Boletos Where bole = '" & cpBoleto.Text & "' and cond = " & RetCodigo & ";", dbOpenDynaset)
  With rsConta
    If .RecordCount > 0 Then
      cpCodigo.Text = !cdsc
      cpNome.Text = !Nome
      cpVencimento.Text = Format(!vcto, "dd/mm/yyyy")
      cpValor.Text = Format$(!corrigido, "#0.00")
      cpNomeCond.Text = NomeCondominio(!cond)
      vCodigo = !cond
      If Not IsNull(!idStatus) Then
        cpStatus.Text = RetornaStatus(!idStatus)
      End If
      Data1.RecordSource = "select * from pagamentos where id_boleto = " & !id & ";"
      Data1.Refresh
    Else
      MsgBox "Número de boleto não encontrado.", vbInformation + vbOKOnly, "Aviso"
      cpBoleto.SetFocus
    End If
  End With
End Sub

Private Sub cmdLocaliza_Click()
  RetNome = ""
  LocalizaBoleto.Situacao = "S"
  LocalizaBoleto.Show 1
  If RetNome <> "" Then
    cpBoleto.Text = RetNome
    AchaBoleto
  End If
End Sub

Private Sub cpBoleto_GotFocus()
  cpBoleto.SelStart = 0
  cpBoleto.SelLength = Len(cpBoleto.Text)
End Sub

Private Sub cpBoleto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    AchaBoleto
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
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
  Refresh
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
End Sub

Private Sub Limpar()
  cpCodigo.Text = ""
  cpNome.Text = ""
  cpBoleto.Text = "0"
  cpValor.Text = ""
  cpVencimento.Text = ""
End Sub
