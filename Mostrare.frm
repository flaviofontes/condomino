VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Mostrarel 
   Caption         =   "Relatório"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   Icon            =   "Mostrare.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   7935
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5685
      Top             =   1185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RelVideo 
      Height          =   3870
      Left            =   960
      TabIndex        =   0
      Top             =   585
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6826
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Mostrare.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Mostrarel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Errado
  Dim Resp  As VbMsgBoxResult
  Dim nFree As Integer
  Dim Strl  As String
  Dim II    As Integer
  If Shift = 2 Then
    If KeyCode = 80 Then
      Resp = MsgBox("Ligue a impressora e clique em Ok.", vbInformation + vbOKCancel, Titulo)
      If Resp = vbOK Then
        If Left(RelVideo.Text, 1) = Chr(27) Then
          Shell App.Path + "\prnrel.bat " & vbFilePrint, vbHide
        Else
          nFree = FreeFile
          Open vbFilePrint For Input As #nFree
          Printer.PrintQuality = -1
          Printer.ScaleMode = 4
          Printer.Font = RelVideo.Font
          While Not EOF(nFree)
            Line Input #nFree, Strl
            If Left(Strl, 1) = Chr(12) Then
              Printer.NewPage
            Else
              Printer.Print Strl
            End If
            DoEvents
          Wend
          Printer.EndDoc
          Close #nFree
          For II = 0 To 15
            DoEvents
          Next II
        End If
      End If
    End If
  End If

Fim:
  Exit Sub

Errado:
  Resume Fim
End Sub

Private Sub Form_Load()
  Refresh
  KeyPreview = True
  With RelVideo
    If Dir(vbFilePrint) <> "" Then
      .FileName = vbFilePrint
    Else
      .Text = ""
    End If
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
  End With
End Sub

Private Sub Form_Resize()
  With RelVideo
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Dir(vbFilePrint) <> "" Then
    Kill vbFilePrint
  End If
End Sub
