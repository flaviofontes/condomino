VERSION 5.00
Begin VB.Form Supervisor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Senha de supervisor"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "Supervisor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar a operação."
      Top             =   570
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   300
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Confirma."
      Top             =   570
      Width           =   1290
   End
   Begin VB.TextBox TxtSenha 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1350
      MaxLength       =   50
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   105
      Width           =   2970
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe a senha"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1140
   End
End
Attribute VB_Name = "Supervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
  sSuper = ""
  Unload Me
End Sub

Private Sub cmdOk_Click()
  sSuper = TxtSenha.Text
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    If KeyCode = 27 Then
      sSuper = ""
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
  Left = (Screen.Width - Me.Width) / 2
  Top = (Screen.Height - Me.Height) / 2
  Refresh
  sSuper = ""
  KeyPreview = True
End Sub

Private Sub TxtSenha_Change()
  cmdOk.Enabled = IIf(Len(TxtSenha.Text) > 0, True, False)
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdOk.Enabled Then
      cmdOk.SetFocus
    End If
  End If
End Sub
