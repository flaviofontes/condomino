VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form CancelaBoleto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelamento de boleto"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "CancelaBoleto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   2220
      TabIndex        =   12
      Top             =   180
      Width           =   435
   End
   Begin VB.TextBox cpBoleto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1440
      MaxLength       =   17
      TabIndex        =   0
      Top             =   600
      Width           =   2235
   End
   Begin VB.CommandButton cmdLocaliza 
      Caption         =   "..."
      Height          =   315
      Left            =   3780
      TabIndex        =   11
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirmar"
      Height          =   405
      Left            =   5520
      TabIndex        =   9
      Top             =   1680
      Width           =   1485
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1425
      TabIndex        =   4
      Top             =   1740
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
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpVencimento 
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   1395
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
      Left            =   2370
      TabIndex        =   2
      Top             =   1005
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
      Left            =   1425
      TabIndex        =   1
      Top             =   1005
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
      Left            =   2700
      TabIndex        =   13
      Top             =   180
      Width           =   4335
      _ExtentX        =   7646
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
   Begin rdActiveText.ActiveText cpCodCond 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   180
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   495
      TabIndex        =   7
      Top             =   1485
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sacado"
      Height          =   195
      Left            =   780
      TabIndex        =   6
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do boleto"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   675
      Width           =   1260
   End
End
Attribute VB_Name = "CancelaBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelBol As Recordset
Dim PodeExcluir As Boolean
Dim i As Integer
Dim id_boleto As Long
Dim tbCondominio As Recordset

Private Sub cmdConfirma_Click()
  If PodeExcluir Then
    Resp = MsgBox("Confirma o cancelamento deste boleto?", vbQuestion + vbYesNo, "Cancelar")
    If Resp = vbYes Then
      DBEngine.BeginTrans
      If IsNull(SelBol!idStatus) Then
        db.Execute "delete from boletodetalhe where id_boleto = " & id_boleto & ";"
        db.Execute "Delete From Boletos Where id = " & id_boleto & ";"
        Limpar
      ElseIf (SelBol!idStatus = 1 Or SelBol!idStatus = 8 Or SelBol!idStatus = 5) Then
        db.Execute "delete from boletodetalhe where id_boleto = " & id_boleto & ";"
        db.Execute "Delete From Boletos Where id = " & id_boleto & ";"
        Limpar
      ElseIf (SelBol!idStatus = 2 Or SelBol!pago = "N") Then
        db.Execute "update Boletos set idstatus = 9, cancelado = 'S' Where id = " & id_boleto & ";"
        Limpar
      Else
        MsgBox "Boleto com status '" & RetornaStatus(SelBol!idStatus) & "' não permite cancelamento.", vbInformation + vbOKOnly, "Aviso"
      End If
      If Err.Number = 0 Then
        DBEngine.CommitTrans
      Else
        DBEngine.Rollback
      End If
    End If
  Else
    MsgBox "Boleto não encontrado ou não pode ser cancelado", vbCritical + vbOKOnly, "Cancelar"
  End If
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpNomeCond.Text = !Nome
        cpCodCond.Text = !Codigo
      End If
    End With
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
    cmdConfirma.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
    If KeyAscii > 0 Then
      Limpar False
    End If
  End If
End Sub

Private Sub cmdLocaliza_Click()
  RetNome = ""
  LocalizaBoleto.Situacao = "N"
  LocalizaBoleto.Show 1
  If RetNome <> "" Then
    cpBoleto.Text = RetNome
    cpCodCond.Text = RetCodigo
    cpNomeCond.Text = NomeCondominio(RetCodigo)
    AchaBoleto
    cmdConfirma.SetFocus
  End If
End Sub

Private Sub AchaBoleto()
  id_boleto = -1
  If RetCodigo > 0 Then
    Set SelBol = db.OpenRecordset("select * from Boletos Where bole = '" & cpBoleto.Text & "' and cond = " & RetCodigo & " and pago = 'N';", dbOpenDynaset)
    With SelBol
      If .RecordCount > 0 Then
        .MoveFirst
        cpCodigo.Text = !cdsc
        cpNome.Text = NomeCompleto(!cdsc)
        cpVencimento.Text = Format(!vcto, "dd/mm/yyyy")
        cpValor.Text = Format$(!corrigido, "#0.00")
        cpNomeCond.Text = NomeCondominio(!cond)
        id_boleto = !id
        PodeExcluir = True
        cmdConfirma.SetFocus
      Else
        MsgBox "Nenhum boleto em aberto foi encontrado para o número '" & cpBoleto.Text & "'." & vbCrLf & "Se o número estiver correto vá em 'Estorno de pagamento'.", vbInformation + vbOKOnly, "Aviso"
        cpBoleto.SetFocus
      End If
    End With
  Else
    Set SelBol = db.OpenRecordset("select * from Boletos Where bole = '" & cpBoleto.Text & "' and pago = 'N';", dbOpenDynaset)
OutraVez:
    With SelBol
      If .RecordCount > 0 Then
        .MoveLast
        If .RecordCount = 1 Then
          .MoveFirst
          cpCodigo.Text = !cdsc
          cpNome.Text = NomeCompleto(!cdsc)
          cpVencimento.Text = Format(!vcto, "dd/mm/yyyy")
          cpValor.Text = Format$(!corrigido, "#0.00")
          cpNomeCond.Text = NomeCondominio(!cond)
          id_boleto = !id
          cmdConfirma.SetFocus
          PodeExcluir = True
        Else
          BoletosAchados.Caption = "Mais de um boleto para: " & cpBoleto.Text
          Call BoletosAchados.Carregar(SelBol)
          If RetCodigo > 0 Then
            Set SelBol = db.OpenRecordset("select * from Boletos Where id = " & RetCodigo & ";", dbOpenDynaset)
            GoTo OutraVez
          End If
        End If
      Else
        MsgBox "Nenhum boleto em aberto foi encontrado para o número '" & cpBoleto.Text & "'." & vbCrLf & "Se o número estiver correto vá em 'Estorno de pagamento'.", vbInformation + vbOKOnly, "Aviso"
        cpBoleto.SetFocus
        PodeExcluir = False
      End If
    End With
  End If
End Sub

Private Sub cpCodCond_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodigo.Text
        If Not .NoMatch Then
          cpNome.Text = !Nome
          cpBoleto.SetFocus
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
        cpBoleto.SetFocus
      End If
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
  KeyPreview = True
  PodeExcluir = False
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
End Sub

Private Sub Limpar(Optional Tudo As Boolean = True)
  cpCodigo.Text = ""
  cpNome.Text = ""
  If Tudo Then
    cpBoleto.Text = ""
  Else
    cpCodCond.Text = ""
    cpNomeCond.Text = ""
  End If
  cpValor.Text = ""
  cpVencimento.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
