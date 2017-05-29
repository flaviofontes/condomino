VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form Extorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estorno de pagamento"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "Extorno.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   3780
      TabIndex        =   12
      Top             =   180
      Width           =   435
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirmar"
      Height          =   405
      Left            =   5520
      TabIndex        =   11
      Top             =   1500
      Width           =   1485
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1425
      TabIndex        =   5
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
      TextMask        =   4
      RawText         =   4
      FontName        =   "Arial"
      FontSize        =   9,75
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpVencimento 
      Height          =   315
      Left            =   1425
      TabIndex        =   4
      Top             =   1215
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
   Begin rdActiveText.ActiveText cpCondominio 
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   870
      Width           =   5595
      _ExtentX        =   9869
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
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   2
      Top             =   525
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
      Left            =   1425
      TabIndex        =   1
      Top             =   525
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
   Begin rdActiveText.ActiveText cpNumero 
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   180
      Width           =   2295
      _ExtentX        =   4048
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
      MaxLength       =   17
      TextMask        =   9
      RawText         =   9
      Mask            =   "#################"
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   495
      TabIndex        =   9
      Top             =   1245
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   930
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sacado"
      Height          =   195
      Left            =   780
      TabIndex        =   7
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do boleto"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   255
      Width           =   1260
   End
End
Attribute VB_Name = "Extorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pode As Boolean
Dim nCond As Long
Dim id_boleto As Long
Dim rs As Recordset

Private Sub AchaBoleto()
  If nCond > 0 Then
    Set rs = db.OpenRecordset("Select * from Boletos where cond = " & nCond & " and bole = '" & cpNumero.Text & "';", dbOpenDynaset)
  Else
    Set rs = db.OpenRecordset("Select * from Boletos where bole = '" & cpNumero.Text & "';", dbOpenDynaset)
  End If
OutraVez:
  With rs
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      If .RecordCount = 1 Then
        If !pago = "N" Then
          MsgBox "Boleto não quitado!", vbInformation + vbOKOnly, "Aviso"
          Pode = False
        Else
          cpCondominio.Text = NomeCondominio(!cond)
          cpCodigo.Text = !cdsc
          cpNome.Text = NomeCompleto(!cdsc)
          cpVencimento.Text = Format(!vcto, "dd/mm/yyyy")
          cpValor.Text = Format(!MENS, "#,##0.00")
          cpCodigo.Text = !cdsc
          id_boleto = !id
          Pode = True
          cmdConfirma.SetFocus
        End If
      Else
        BoletosAchados.Caption = "Mais de um boleto para: " & cpNumero.Text
        Call BoletosAchados.Carregar(rs)
        If RetCodigo > 0 Then
          Set rs = db.OpenRecordset("select * from Boletos Where id = " & RetCodigo & ";", dbOpenDynaset)
          GoTo OutraVez
        End If
      End If
    Else
      MsgBox "Número de boleto não encontrado para o condominio selecionado.", vbInformation + vbOKOnly, "Aviso"
      cpNumero.SetFocus
      Pode = False
    End If
  End With
  Set rs = Nothing
End Sub

Private Sub cmdConfirma_Click()
  If Pode Then
    Resp = MsgBox("Confirma o extorno do pagamento deste boleto?", vbQuestion + vbYesNo, "Extorno")
    If Resp = vbYes Then
      db.Execute "Update Boletos Set Pago = 'N', idstatus = 2 Where id = " & id_boleto & ";"
      db.Execute "delete from pagamentos where id_boleto = " & id_boleto & ";"
      Limpar
    End If
  Else
    MsgBox "Boleto não encontrado ou não quitado, não pode estornar pagamento.", vbInformation + vbOKOnly, "Aviso"
  End If
End Sub

Private Sub cmdLocalizar_Click()
  RetNome = ""
  LocalizaBoleto.Situacao = "S"
  LocalizaBoleto.Show 1
  If RetNome <> "" Then
    cpNumero.Text = RetNome
    nCond = RetCodigo
    AchaBoleto
    cmdConfirma.SetFocus
  End If
End Sub

Private Sub cpNumero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    AchaBoleto
  Else
    KeyAscii = vNumero(KeyAscii)
    If KeyAscii > 0 Then
      nCond = 0
      Limpar False
    End If
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
  Pode = False
  KeyPreview = True
  Refresh
End Sub

Private Sub Limpar(Optional Tudo As Boolean = True)
  cpCodigo.Text = ""
  cpCondominio.Text = ""
  cpNome.Text = ""
  If Tudo Then
    cpNumero.Text = ""
  End If
  cpValor.Text = ""
  cpVencimento.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub
