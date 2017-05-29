VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form BoletosOld 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de Boletos Bancários"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   Icon            =   "BoletosOld.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cpSocio 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   525
      Width           =   4410
   End
   Begin VB.TextBox cpCota 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      MaxLength       =   8
      TabIndex        =   1
      Top             =   525
      Width           =   750
   End
   Begin MSMask.MaskEdBox cpVenc 
      Height          =   330
      Left            =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   795
      Left            =   1642
      Picture         =   "BoletosOld.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   975
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Fechar"
      Height          =   795
      Left            =   3322
      Picture         =   "BoletosOld.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   975
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Do sócio"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   585
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vencidos em"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "BoletosOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vbCodigo As Long

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Errado
  Dim SelBoleto As String
  
  If cpVenc.ClipText = "" And cpCota.Text = "" Then
    MsgBox "Por favor informe o vencimento ou o sócio.", vbCritical + vbOKOnly, Titulo
    GoTo Fim
  End If
  
  If cpCota.Text <> "" Then
    SelBoleto = "{Boletos.Cota}= '" & cpCota.Text & "'"
  Else
    SelBoleto = "{Boletos.VCTO}=Date(" & Year(cpVenc.Text) & "," & Month(cpVenc.Text) & "," & Day(cpVenc.Text) & ")"
  End If
  cmdCancelar.Enabled = False
  cmdPrint.Enabled = False
  Call Relatorio(0, Trim(Parametros.Dados) & "\dados.mdb", App.Path & "\oldboletos.rpt", SelBoleto, , "Bloquetos de cobrança bancária")

Fim:
  cmdCancelar.Enabled = True
  cmdPrint.Enabled = True
  Exit Sub

Errado:
  MsgBox "Erro No. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, Titulo
  Resume Fim

End Sub

Private Sub cpCota_GotFocus()
  cpCota.SelStart = 0
  cpCota.SelLength = Len(cpCota.Text)
End Sub

Private Sub cpCota_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Len(cpCota.Text) > 0 Then
      With daClientes
        .Index = "PerCota"
        .Seek "=", cpCota
        If Not .NoMatch Then
          cpSocio.Text = !Nome
          cmdPrint.SetFocus
        Else
          Beep
          MsgBox "Esta cota não está cadastrada! Verifique por favor.", vbInformation + vbOKOnly, Titulo
        End If
      End With
    Else
      RetCodigo = 0
      LocClientes.Show 1
      If RetCodigo > 0 Then
        With daClientes
          .Index = "PerCodigo"
          .Seek "=", RetCodigo
          If Not .NoMatch Then
            cpSocio.Text = !Nome
            cpCota.Text = !NumeroDaCota
            cmdPrint.SetFocus
          End If
        End With
      End If
    End If
  End If
End Sub

Private Sub cpVenc_GotFocus()
  cpVenc.SelStart = 0
  cpVenc.SelLength = Len(cpVenc.Text)
End Sub

Private Sub cpVenc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0: cmdPrint.SetFocus
End Sub

Private Sub cpVenc_LostFocus()
  If cpVenc.ClipText <> "" Then
    If TestaData(cpVenc.Text) = False Then
      cpVenc.SetFocus
    Else
      cpCota.Text = ""
      cpSocio.Text = ""
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
  Top = 0
  Left = 0
  Refresh
  DoEvents
  KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.KeyPreview = False
End Sub
