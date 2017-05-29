VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form TipoCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Cliente"
   ClientHeight    =   2820
   ClientLeft      =   2010
   ClientTop       =   1545
   ClientWidth     =   4965
   Icon            =   "TipoCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TipoCliente.frx":000C
      Height          =   2430
      Left            =   15
      OleObjectBlob   =   "TipoCliente.frx":003A
      TabIndex        =   0
      Top             =   60
      Width           =   4905
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   495
      Top             =   1545
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\progvb60\bahia\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "TipoDoCliente"
      Top             =   2085
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "(Selecione o registro: DEL=Excuir)    ESC=Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   2550
      Width           =   4755
   End
End
Attribute VB_Name = "TipoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewCod As Long

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
  If Corrente.OpK003 = 0 Then
    MsgBox S_Permissao, vbInformation + vbOKOnly, Titulo
    Cancel = True
  Else
    Resp = MsgBox("Você Confirma a Exclusão Deste Tipo?", vbQuestion + vbYesNo, Titulo)
    If Resp = vbNo Then Cancel = True
  End If
End Sub

Private Sub DBGrid1_Error(ByVal DataError As Integer, Response As Integer)
  'MsgBox Error$(DataError), vbCritical + vbOKOnly, titulo
  Response = 0
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  Select Case DBGrid1.Col
    Case 1
      Select Case KeyAscii
        Case 8
          KeyAscii = KeyAscii
        Case Is >= 97
          KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
          KeyAscii = KeyAscii
      End Select
    Case 2
      Select Case KeyAscii
        Case 8
          KeyAscii = KeyAscii
        Case 48 To 57
          KeyAscii = KeyAscii
        Case 44, 46
          If InStr(DBGrid1.Text, ",") = 0 Then
            KeyAscii = 44
          Else
            KeyAscii = 0
          End If
        Case Else
          KeyAscii = 0
      End Select
  End Select
End Sub

Private Sub DBGrid1_OnAddNew()
  With daTipo
    If .RecordCount > 0 Then
      .MoveLast
      NewCod = !Codigo + 1
    Else
      NewCod = 1
    End If
  End With
  DBGrid1.Col = 0
  DBGrid1.Text = Str(NewCod)
  DBGrid1.Col = 1
  DBGrid1.EditActive = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 27
      Unload Me
  End Select
End Sub

Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Refresh
  DoEvents
  Data1.DatabaseName = Trim(Parametros.Dados) & "\dados.mdb"
  Me.KeyPreview = True
  If Corrente.OpK001 = 0 Then DBGrid1.AllowAddNew = False
  If Corrente.OpK002 = 0 Then DBGrid1.AllowUpdate = False
  Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Data1.Recordset.Index = "TPCodigo"
End Sub
