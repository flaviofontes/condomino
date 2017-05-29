VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelMovimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de boletos baixados"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "RelMovimento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpFim 
      Height          =   315
      Left            =   2220
      TabIndex        =   2
      Top             =   660
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
   Begin VB.ComboBox cpCondominio 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   420
      Left            =   4260
      TabIndex        =   3
      Top             =   540
      Width           =   1260
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   660
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Até"
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   720
      Width           =   345
   End
End
Attribute VB_Name = "RelMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
  Dim sFormula As String
  If cpMes.Text = "" Or cpFim.Text = "" Then
    MsgBox "Preencha os dois campos.", vbCritical + vbOKOnly, "Aviso"
  Else
    If cpCondominio.Text = "" Then
      MsgBox "Selecione o condomínio.", vbInformation + vbOKOnly, "Condominio"
    ElseIf cpCondominio.Text <> "Todos" Then
      sFormula = "{boletos.cond} = " & cpCondominio.ItemData(cpCondominio.ListIndex) & _
                 "and {boletos.cancelado} = 'N' " & _
                 " and {boletos.dtpgto} >= date(" & Year(cpMes.Text) & ", " & _
                 Month(cpMes.Text) & ", " & Day(cpMes.Text) & ") and {boletos.dtpgto} <= date(" & _
                 Year(cpFim.Text) & ", " & Month(cpFim.Text) & ", " & Day(cpFim.Text) & ")"
    End If
  End If
  Relatorio 0, Parametros.Dados, sFormataCaminho(App.Path) & "bolbaixados.rpt", sFormula, "", , "Orçamento Mês Base " & cpMes.Text
End Sub

Private Sub cpFim_GotFocus()
  If cpFim.Text = "" Then
    cpFim.Text = cpMes.Text
  End If
End Sub

Private Sub cpFim_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrint.SetFocus
  End If
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
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
  Refresh
  KeyPreview = True
  With tbCondominio
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cpCondominio.AddItem (!Nome & "")
        cpCondominio.ItemData(cpCondominio.NewIndex) = !Codigo
        .MoveNext
      Loop
      cpCondominio.ListIndex = 0
    End If
  End With
End Sub
