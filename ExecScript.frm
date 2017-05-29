VERSION 5.00
Begin VB.Form ExecScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Executando script"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "ExecScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5340
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Executando Script de atualização. Por favor aguarde..."
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   3885
   End
End
Attribute VB_Name = "ExecScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VBScript As MSScriptControl.ScriptControl

Dim sDados As String
Dim mSoundex As Soundex

Property Get Onde() As String
  Onde = sDados
End Property

Private Sub Form_Load()
  Set mSoundex = New Soundex
  Set VBScript = New MSScriptControl.ScriptControl
  sDados = Parametros.dados
  'set Language
  VBScript.Language = "VBScript"
  'Add object
  VBScript.AddObject "ExecScript", ExecScript
  VBScript.AddObject "mSoundex", mSoundex
  'allowed to display user-interface elements.
  VBScript.AllowUI = True
  Refresh
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set VBScript = Nothing
End Sub

Private Sub Timer1_Timer()
On Error GoTo Errado
  Timer1.Enabled = False
  If Not achaScript Then
    gravaScript
    VBScript.AddCode Text1.Text
  End If
  Unload Me
  Exit Sub
Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  Err.Clear
  Unload Me
End Sub

Private Function achaScript() As Boolean
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("scripts", dbOpenTable)
  achaScript = False
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If !texto = Text1.Text Then
          achaScript = True
          Exit Do
        End If
        .MoveNext
      Loop
    End If
  End With
  
  Set rs = Nothing
End Function

Private Sub gravaScript()
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("scripts", dbOpenTable)
  With rs
    .AddNew
    !texto = Text1.Text
    .Update
  End With
  
  Set rs = Nothing
End Sub

