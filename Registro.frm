VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Registro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro do Sistema"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Registro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirma"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5535
      TabIndex        =   19
      Top             =   105
      Width           =   930
   End
   Begin VB.TextBox cpRegistro 
      Height          =   315
      Left            =   1710
      MaxLength       =   30
      TabIndex        =   18
      Top             =   2790
      Width           =   4020
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Sair"
      Height          =   345
      Left            =   5535
      TabIndex        =   16
      Top             =   450
      Width           =   930
   End
   Begin MSMask.MaskEdBox cpInscricao 
      Height          =   315
      Left            =   1305
      TabIndex        =   7
      Top             =   2010
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox cpCnpj 
      Height          =   315
      Left            =   1305
      TabIndex        =   6
      Top             =   1695
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   18
      Mask            =   "99.999.999/9999-99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox cpCep 
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Top             =   1365
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      Mask            =   "99999-999"
      PromptChar      =   "_"
   End
   Begin VB.TextBox cpEstado 
      Height          =   315
      Left            =   1305
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1365
      Width           =   465
   End
   Begin VB.TextBox cpCidade 
      Height          =   315
      Left            =   1305
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1050
      Width           =   2955
   End
   Begin VB.TextBox cpBairro 
      Height          =   315
      Left            =   1305
      MaxLength       =   30
      TabIndex        =   2
      Top             =   735
      Width           =   2955
   End
   Begin VB.TextBox cpEndereco 
      Height          =   315
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   1
      Top             =   420
      Width           =   4050
   End
   Begin VB.TextBox cpEmpresa 
      Height          =   315
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   0
      Top             =   105
      Width           =   4050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Código de Registro"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   2850
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Insc. Est."
      Height          =   195
      Left            =   540
      TabIndex        =   15
      Top             =   2100
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ"
      Height          =   195
      Left            =   780
      TabIndex        =   14
      Top             =   1785
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Left            =   1950
      TabIndex        =   13
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   705
      TabIndex        =   12
      Top             =   1425
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   705
      TabIndex        =   11
      Top             =   1110
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Left            =   795
      TabIndex        =   10
      Top             =   810
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Endereço"
      Height          =   195
      Left            =   525
      TabIndex        =   9
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Razão Social"
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   180
      Width           =   945
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary
Dim tbParametros As Recordset

Private Sub cmdCancela_Click()
  End
End Sub

Private Sub cmdConfirma_Click()
  Dim i         As Integer
  Dim gRegistro As String
  Dim DatFile   As String
  Dim Parte1    As String
  Dim Parte2    As String
  Dim Parte3    As String
  For i = 1 To Len(cpEmpresa.Text)
    Parte1 = Parte1 & Hex(Asc(Mid(cpEmpresa.Text, i, 1)))
  Next i
  For i = 1 To Len(cpCnpj.Text)
    Parte2 = Parte2 & Hex(Asc(Mid(cpCnpj.Text, i, 1)))
  Next i
  For i = 1 To Len(cpInscricao.Text)
    Parte3 = Parte3 & Hex(Asc(Mid(cpInscricao.Text, i, 1)))
  Next i
  gRegistro = Left(Parte1, 10) & Left(Parte2, 10) & Left(Parte3, 10)
  If gRegistro = cpRegistro.Text Then
    DatFile = App.Path & "\supersoft.dat"
    With tbParametros
      If .RecordCount > 0 Then
        .MoveFirst
        .Edit
      Else
        .AddNew
      End If
      !Empresa = cpEmpresa.Text
      !endereco = cpEndereco.Text
      !Cidade = cpCidade.Text
      !bairro = cpBairro.Text
      !cep = cpCep.Text
      !estado = cpEstado.Text
      !Cnpj = cpCnpj.Text
      !Inscricao = cpInscricao.Text
      !Uso = 255
      !Registro = cpRegistro.Text
      .Update
    End With
    WritePrivateProfileString "Info", "EP", CStr(Codifica(cpEmpresa.Text)), DatFile
    WritePrivateProfileString "Info", "ED", CStr(Codifica(cpEndereco.Text)), DatFile
    WritePrivateProfileString "Info", "EB", CStr(Codifica(cpBairro.Text)), DatFile
    WritePrivateProfileString "Info", "EC", CStr(Codifica(cpCidade.Text)), DatFile
    WritePrivateProfileString "Info", "EE", CStr(Codifica(cpCep.Text)), DatFile
    WritePrivateProfileString "Info", "EU", CStr(Codifica(cpEstado.Text)), DatFile
    WritePrivateProfileString "Info", "EG", CStr(Codifica(cpCnpj.Text)), DatFile
    WritePrivateProfileString "Info", "EI", CStr(Codifica(cpInscricao.Text)), DatFile
    MsgBox "Sistema registrado com sucesso!", vbExclamation + vbOKOnly, Titulo
    Unload Me
  Else
    MsgBox "O código de registro informado não é válido!", vbCritical + vbOKOnly, Titulo
  End If
End Sub

Private Sub cpBairro_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpCidade.SetFocus
    Case Is >= 97
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
      KeyAscii = KeyAscii
  End Select
End Sub

Private Sub cpCep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCnpj.SetFocus
  End If
End Sub

Private Sub cpCidade_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpEstado.SetFocus
    Case Is >= 97
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
      KeyAscii = KeyAscii
  End Select
End Sub

Private Sub cpCnpj_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpInscricao.SetFocus
  End If
End Sub

Private Sub cpEmpresa_Change()
  If Len(cpEmpresa.Text) > 0 Then
    cmdConfirma.Enabled = True
  Else
    cmdConfirma.Enabled = False
  End If
End Sub

Private Sub cpEmpresa_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpEndereco.SetFocus
    Case Is >= 97
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
      KeyAscii = KeyAscii
  End Select
End Sub

Private Sub cpEndereco_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpBairro.SetFocus
    Case Is >= 97
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
      KeyAscii = KeyAscii
  End Select
End Sub

Private Sub cpEstado_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpCep.SetFocus
    Case Is >= 97
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
      KeyAscii = KeyAscii
  End Select
End Sub

Private Sub cpInscricao_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cpRegistro.SetFocus
    Case Is >= 97
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
      KeyAscii = KeyAscii
  End Select
End Sub

Private Sub cpRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cmdConfirma.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set tbParametros = db.OpenRecordset("Parametros", dbOpenTable)
  Me.Refresh
  DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbParametros = Nothing
End Sub
