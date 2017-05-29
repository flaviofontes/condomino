VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form Quitar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quitação de boletos"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "Quitar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpDtPgto 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton cmdLocaliza 
      Caption         =   "..."
      Height          =   315
      Left            =   4020
      TabIndex        =   12
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
   Begin VB.CommandButton cmdConfirma 
      BackColor       =   &H0000C000&
      Caption         =   "&Confirmar"
      Height          =   405
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1860
      Width           =   1485
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpVencimento 
      Height          =   315
      Left            =   1665
      TabIndex        =   3
      Top             =   1215
      Width           =   1350
      _ExtentX        =   2381
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextMask        =   3
      RawText         =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpNomeCond 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
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
      Caption         =   "Data do pagamento"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1980
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   720
      TabIndex        =   11
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   735
      TabIndex        =   8
      Top             =   1305
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inquilino"
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do boleto"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   6
      Top             =   555
      Width           =   1260
   End
End
Attribute VB_Name = "Quitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vCodigo As Long
Dim i As Integer
Dim rsConta As Recordset

Private Sub cmdConfirma_Click()
  If Not rsConta Is Nothing Then
    Resp = MsgBox("Confirma o pagamento deste boleto?", vbQuestion + vbYesNo, "Quitar")
    If Resp = vbYes Then
      If Not IsDate(cpDtPgto.Text) Then
        MsgBox "Data de pagamento não informada ou inválida.", vbInformation + vbOKOnly, "Aviso"
        cpDtPgto.SetFocus
        Exit Sub
      End If
      With rsConta
        If .RecordCount > 0 Then
          .MoveFirst
          If !acumulado = "N" Then
            If cpValor.Text < !corrigido Then
              Resp = MsgBox("O valor pago é menor que o valor do boleto. Prosseguir?", vbQuestion + vbYesNo, "Quitar")
              If Resp = vbNo Then
                Exit Sub
              End If
            End If
            .Edit
            !pago = "S"
            !dtpgto = cpDtPgto.Text
            !vlpago = cpValor.Text
            If !idStatus = 1 Or !idStatus = 2 Or !idStatus = 4 Or !idStatus = 6 Then
              !idStatus = 9
            Else
              !idStatus = 5
            End If
            .Update
          Else
            db.Execute "Update Boletos Set Pago = 'S', dtpgto = #" & Format$(cpDtPgto.Text, "MM/dd/yyyy") _
                & "#, vlpago = " & Replace(cpValor.Text, ",", ".") & ", idstatus = 9 Where pago = 'N' and VCTO <= #" & Format$(cpVencimento.Text, "mm/dd/yyyy") & "# And CDSC = " & cpCodigo.Text & ";"
          End If
          Pagamento !id, CDate(cpDtPgto.Text), CDbl(cpValor.Text), "ESCRITÓRIO", 0
        End If
      End With
      Set rsConta = Nothing
      Limpar
    End If
  End If
End Sub

Private Sub AchaBoleto()
  If RetCodigo > 0 Then
    Set rsConta = db.OpenRecordset("select * from Boletos Where bole = '" & cpBoleto.Text & "' and cond = " & RetCodigo & ";", dbOpenDynaset)
    With rsConta
      If .RecordCount > 0 Then
        .MoveFirst
        If !pago = "S" Then
          MsgBox "Boleto já quitado!", vbInformation + vbOKOnly, "Aviso"
          Set rsConta = Nothing
        Else
          cpCodigo.Text = !cdsc
          cpNome.Text = NomeCompleto(!cdsc)
          cpVencimento.Text = Format(!vcto, "dd/mm/yyyy")
          cpValor.Text = Format$(!corrigido, "#0.00")
          cpDtPgto.Text = Date
          cpNomeCond.Text = NomeCondominio(!cond)
          vCodigo = !cond
          cmdConfirma.SetFocus
        End If
      Else
        MsgBox "Número de boleto não encontrado.", vbInformation + vbOKOnly, "Aviso"
        cpBoleto.SetFocus
      End If
    End With
  Else
    Set rsConta = db.OpenRecordset("select * from Boletos Where bole = '" & cpBoleto.Text & "';", dbOpenDynaset)
OutraVez:
    With rsConta
      If .RecordCount > 0 Then
        .MoveLast
        If .RecordCount = 1 Then
          .MoveFirst
          If !pago = "S" Then
            MsgBox "Boleto já quitado!", vbInformation + vbOKOnly, "Aviso"
            Set rsConta = Nothing
          Else
            cpCodigo.Text = !cdsc
            cpNome.Text = NomeCompleto(!cdsc)
            cpVencimento.Text = Format(!vcto, "dd/mm/yyyy")
            cpValor.Text = Format$(!corrigido, "#0.00")
            cpDtPgto.Text = Date
            cpNomeCond.Text = NomeCondominio(!cond)
            vCodigo = !cond
            cmdConfirma.SetFocus
          End If
        Else
          BoletosAchados.Caption = "Mais de um boleto para: " & cpBoleto.Text
          Call BoletosAchados.Carregar(rsConta)
          If RetCodigo > 0 Then
            Set rsConta = db.OpenRecordset("select * from Boletos Where id = " & RetCodigo & ";", dbOpenDynaset)
            GoTo OutraVez
          End If
        End If
      Else
        MsgBox "Número de boleto não encontrado.", vbInformation + vbOKOnly, "Aviso"
        cpBoleto.SetFocus
      End If
    End With
  End If
End Sub

Private Sub cmdLocaliza_Click()
  RetNome = ""
  LocalizaBoleto.Situacao = "N"
  LocalizaBoleto.Show 1
  If RetNome <> "" Then
    cpBoleto.Text = RetNome
    AchaBoleto
    cmdConfirma.SetFocus
  End If
End Sub

Private Sub cpBoleto_Change()
  RetCodigo = 0
End Sub

Private Sub cpBoleto_GotFocus()
  RetCodigo = 0
  cpBoleto.SelStart = 0
  cpBoleto.SelLength = Len(cpBoleto.Text)
End Sub

Private Sub cpBoleto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    AchaBoleto
    cpDtPgto.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpDtPgto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdConfirma.SetFocus
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDtPgto.SetFocus
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
End Sub

Private Sub Limpar()
  cpCodigo.Text = ""
  cpNome.Text = ""
  cpBoleto.Text = "0"
  cpValor.Text = ""
  cpVencimento.Text = ""
  cpDtPgto.Text = Date
End Sub
