VERSION 5.00
Object = "{CCF682E3-14F1-11D7-80E2-10E08C102140}#1.0#0"; "ZProgBar.ocx"
Begin VB.Form Backup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup dos dados"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   HelpContextID   =   4260
   Icon            =   "backup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProsseguir 
      Caption         =   "&Prosseguir"
      Height          =   795
      Left            =   2220
      Picture         =   "backup.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1980
      Top             =   840
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   1620
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1500
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ZealProgressBar.ProgressBar Barra 
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   1860
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0
   End
   Begin VB.CommandButton cmdLocaliza 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3990
      TabIndex        =   5
      Top             =   1035
      Width           =   690
   End
   Begin VB.TextBox cpLocal 
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   3930
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   795
      Left            =   3540
      Picture         =   "backup.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   1140
   End
   Begin VB.OptionButton OpRestaurar 
      Caption         =   "&Restaurar"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   990
   End
   Begin VB.OptionButton OpBackup 
      Caption         =   "&Backup"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   4635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo/Local do backup"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   810
      Width           =   1815
   End
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NomeFile(2) As String
Dim lblProgresso As String
Dim CAB As New cCAB

Public SemMsg As Boolean

Private Sub FazerBackup()
  Dim sDiretorio As String
  Dim sRet As String
  Dim sTemp As String
  
  sTemp = cpLocal.Text
  
  sTemp = sFormataCaminho(sTemp) & "supersoft " & Format$(Now, "ddMMyy HHMM") & ".bkp"
  
  If sysFiles.FileExists(sTemp) Then
    sysFiles.DeleteFile sTemp, True
  End If
    
  If cpLocal.Text <> "" Then
    Me.MousePointer = 11
    sDiretorio = Parametros.Dados
    NomeFile(0) = sDiretorio
    NomeFile(1) = sFormataCaminho(App.Path) & "supersoft.dat"
    Compacta sTemp, NomeFile
    If Not SemMsg Then
      If sysFiles.FileExists(sTemp) Then
        MsgBox "Arquivo '" & sTemp & "' gerado com sucesso!", vbExclamation + vbOKOnly, "Backup"
      Else
        MsgBox "Erro durante Backup!", vbExclamation + vbOKOnly, "Backup"
      End If
    Else
      Unload Me
    End If
    Me.MousePointer = 0
  Else
    MsgBox "Informe o arquivo/local do backup.", vbExclamation + vbOKOnly, "Backup"
  End If
End Sub


Private Sub RestaurarBackup()

  'declara variável
  Dim sDiretorio As String
  Dim sDirDados As String
  Dim sFile As String
  Dim sPathDesc As String

  sPathDesc = sFormataCaminho(App.Path) + "temp\"
  If sysFiles.FolderExists(sPathDesc) Then
    Call sysFiles.DeleteFile(sPathDesc & "*.*", True)
  Else
    Call sysFiles.CreateFolder(sPathDesc)
  End If
  
  If cpLocal.Text = Empty Then
    MsgBox "Localize o arquivo de backup.", vbInformation + vbOKOnly, "Backup"
    Exit Sub
  End If
  
  sFile = cpLocal.Text
  If sysFiles.FileExists(sFile) Then
    DesCompacta sFile, "*.*", sPathDesc, False
    If sysFiles.FileExists(sPathDesc + "dados.mdb") Then
      If sysFiles.FileExists(Parametros.Dados) Then
        Exclui_Arquivo Parametros.Dados, True
        Exclui_Arquivo sFormataCaminho(App.Path) + "supersoft.dat", True
      End If
      If CAB.Copiar(sPathDesc + "dados.mdb", Parametros.Dados) = Sucesso Then
        Call sysFiles.CopyFile(sPathDesc + "supersoft.dat", sFormataCaminho(App.Path))
        Call sysFiles.DeleteFile(sPathDesc + "*.*", True)
        MsgBox "Restauração do backup realizada com sucesso!", vbExclamation + vbOKOnly, "Backup"
      Else
        MsgBox "Erro durante restauração do backup!", vbExclamation + vbOKOnly, "Backup"
      End If
    End If
  End If
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdLocaliza_Click()
  Dim sDir As String
  
  sDir = MinhaDll.sProcuraPorDiretorio("Local para o backup?", Me, OpRestaurar.Value)
  
  If sDir <> "" Then
    cpLocal.Text = sDir
  Else
    cpLocal.Text = ""
  End If
End Sub

Private Sub cmdProsseguir_Click()
  cmdProsseguir.Enabled = False
  cmdCancelar.Enabled = False
  If OpBackup.Value Then
    CloseDataBase
    FazerBackup
    Call AbrirArquivos(Parametros.Dados, 5)
    db.Execute "update versaodb set ultimo_backup = #" & Format$(Date, "MM/dd/yyyy") & "#;"
  Else
    CloseDataBase
    RestaurarBackup
    Call AbrirArquivos(Parametros.Dados, 5)
  End If
  cmdProsseguir.Enabled = True
  cmdCancelar.Enabled = True
End Sub

Private Sub Form_Load()
  InicializaZip Me, txtZip
  DoEvents
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SemMsg = False
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If SemMsg Then
    FazerBackup
  End If
End Sub

Private Sub txtZip_Change()
  'Indicação de progresso da compactação/descompactação por arquivo
  '----------------------------------------------------------------
  Label2.Caption = TipoAção(Val(GetAction(txtZip))) & " " & GetFileName(txtZip) '& " -> " '& GetPercentComplete(txtZip) & "%"
  'Tipo de ação que esta sendo feita no momento
  Barra.BarText = "%" '& TipoAção(Val(GetAction(txtZip)))
  If Barra.Value <> Int(Val(GetPercentComplete(txtZip))) Then
    Barra.Value = Int(Val(GetPercentComplete(txtZip)))
    DoEvents
  End If
  'Força a atualização da tela
End Sub
