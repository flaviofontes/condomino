VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form AtualizaAtrazados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Correção de boletos ja vencidos"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "AtualizaAtrazados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocCond 
      Caption         =   "..."
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.CheckBox ChTodos 
      Caption         =   "Todos os condomínios"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Corrigir"
      Height          =   390
      Left            =   5700
      TabIndex        =   3
      Top             =   570
      Width           =   1410
   End
   Begin rdActiveText.ActiveText cpNomeCond 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
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
   Begin rdActiveText.ActiveText cpCodCond 
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
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "AtualizaAtrazados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbCondominio As Recordset

Private Sub ChTodos_Click()
  If ChTodos.Value = 1 Then
    cpCodCond.Locked = True
    cmdLocCond.Enabled = False
  Else
    cpCodCond.Locked = False
    cmdLocCond.Enabled = True
  End If
End Sub

Private Sub cmdLocCond_Click()
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
    cmdPrint.SetFocus
  End If
End Sub

Private Sub cmdPrint_Click()
  Resp = MsgBox("Deseja corrigir agora os lançamentos.", vbQuestion + vbOKCancel, "Corrigir")
  If Resp = vbOK Then
    Progresso.Caption = "Corrigindo"
    Progresso.Show vbModeless
    Call SetTopMostWindow(Progresso.hWnd, True)
    DoEvents
    If ChTodos.Value = 1 Then
      With tbCondominio
        .Index = "codigoid"
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            Progresso.Caption = "Corrigindo"
            Progresso.Refresh
            Call CalculaJuros(!Codigo, Progresso)
            .MoveNext
          Loop
        End If
      End With
    Else
      If Val(cpCodCond.Text) > 0 Then
        With tbCondominio
          .Index = "codigoid"
          .Seek "=", Val(cpCodCond.Text)
          If Not .NoMatch Then
            Progresso.Caption = "Corrigindo"
            Progresso.Refresh
            Call CalculaJuros(!Codigo, Progresso)
          End If
        End With
      End If
    End If
    DoEvents
    Call SetTopMostWindow(Progresso.hWnd, False)
    Unload Progresso
    MsgBox "Todos os valores foram corrigidos com sucesso!", vbInformation + vbOKOnly, "Corrigir"
  End If
End Sub

Private Sub cpCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
  End If
End Sub

Private Sub cpCodCond_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodCond.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodCond.Text
        If Not .NoMatch Then
          cpNomeCond.Text = !Nome
          cmdPrint.SetFocus
        Else
          MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
          cpCodCond.SetFocus
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
            cpNomeCond.Text = !Nome
            cpCodCond.Text = !Codigo
          End If
        End With
        cmdPrint.SetFocus
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
  KeyPreview = True
  Principal.balao.ShowBalloonTip "Verifique a data do computador, pois esta é utilizada para os calculos.", beInformation, "Atualizar atrazados", 5000
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
