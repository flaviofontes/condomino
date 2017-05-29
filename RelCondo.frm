VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelCondo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de condomínios"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "RelCondo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   5955
      TabIndex        =   2
      Top             =   540
      Width           =   1170
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
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
      Left            =   1020
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
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "RelCondo"
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
  Dim sSql As String
  Dim rs As Recordset
  Dim sFrac As Double
  'cmdPrint.Enabled = False
  
  Set rs = db.OpenRecordset("SELECT Sum(FRACAO.FRACAO) AS SomaDeFRACAO " _
    & "FROM (CONDOMINIO INNER JOIN ASSOCIADOS ON CONDOMINIO.CODIGO = ASSOCIADOS.CONDOMINIO) INNER JOIN FRACAO ON (ASSOCIADOS.CODIGO = FRACAO.ID_ASSOCIADO) AND (ASSOCIADOS.CODIGO = FRACAO.ID_ASSOCIADO) " _
    & "GROUP BY CONDOMINIO.CODIGO HAVING (((CONDOMINIO.CODIGO)=" & cpCodigo.Text & "));", dbOpenDynaset)
  
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    sFrac = rs!somadefracao
  Else
    sFrac = 0
  End If
  Set rs = Nothing
  sSql = "{EDIFICIO.CODIGO} = " & cpCodigo.Text
  Call RelatoriosRPT.Carregar("{ASSOCIADOS.PROPRIETARIO};crAscendingOrder;CONDOMINIO|", Parametros.dados, sSql, "Relatório de condomínios.", sFormataCaminho(App.Path) & "condominio.rpt", , , "fracao|" & sFrac)
  'cmdPrint.Enabled = True
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
          cmdPrint.SetFocus
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
        cmdPrint.SetFocus
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
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
  KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
