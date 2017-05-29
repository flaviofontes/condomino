VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelAsBoleto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relat�rio de boletos gerados"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "RelAsBoleto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.CheckBox chAvulsos 
      Caption         =   "Avulsos"
      Height          =   195
      Left            =   2460
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   540
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   5775
      TabIndex        =   4
      Top             =   570
      Width           =   1410
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   6
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
      Left            =   1110
      TabIndex        =   0
      Top             =   120
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condom�nio"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "M�s"
      Height          =   195
      Left            =   660
      TabIndex        =   5
      Top             =   600
      Width           =   300
   End
End
Attribute VB_Name = "RelAsBoleto"
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
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condominio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If Trim(cpMes.Text) <> "" Then
    If Not IsDate("25/" & cpMes.Text) Then
      MsgBox "O m�s/ano informado n�o � v�lido.", vbInformation + vbOKOnly, "Aviso"
      cpMes.SetFocus
      Exit Sub
    End If
    If chAvulsos.Value = 0 Then
      RelatoriosRPT.Carregar "", Parametros.dados, "{BOLETOS.IDSTATUS} <> 9 and {BOLETOS.TRAN} = '" & cpMes.Text & "'", "Relat�rio de boletos do m�s " & cpMes.Text, sFormataCaminho(App.Path) & "asboletos.rpt", "{BOLETOS.COND} = " & cpCodigo.Text
    Else
      RelatoriosRPT.Carregar "", Parametros.dados, "{BOLETOS.IDSTATUS} <> 9 and left({BOLETOS.TRAN},2) = 'AV' and Month({BOLETOS.DATA}) = " & Left(cpMes.Text, 2) & " and Year({BOLETOS.DATA}) = " & Right(cpMes.Text, 4), "Relat�rio de boletos do m�s " & cpMes.Text, sFormataCaminho(App.Path) & "asboletos.rpt", "{BOLETOS.COND} = " & cpCodigo.Text
    End If
  Else
    RelatoriosRPT.Carregar "{BOLETOS.IDSTATUS} <> 9 and {BOLETOS.VCTO};crAscendingOrder;BOLETOS|", Parametros.dados, "{BOLETOS.PAGO} = 'N'", "Relat�rio de boletos do m�s " & cpMes.Text, sFormataCaminho(App.Path) & "asboletos.rpt", "{BOLETOS.COND} = " & cpCodigo.Text
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
          cpMes.SetFocus
        Else
          MsgBox "O c�digo informado n�o existe!", vbInformation + vbOKOnly, "Aviso"
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
        cpMes.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpMes_GotFocus()
  cpMes.SelStart = 0
  cpMes.SelLength = Len(cpMes.Text)
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  KeyPreview = True
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
