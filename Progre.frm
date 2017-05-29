VERSION 5.00
Begin VB.Form Progre 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Progre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   615
      Left            =   15
      TabIndex        =   0
      Top             =   150
      Width           =   4860
   End
End
Attribute VB_Name = "Progre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    If InStr(Label1.Caption, "Tecle ESC") > 0 Then
      If RelatoriosRPT.Cancela = True Then
        Unload Me
      Else
        Label1.Caption = "Cancelando. Aguarde..."
        Label1.Refresh
        RelatoriosRPT.Cancela = True
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  KeyPreview = True
End Sub
