VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Backup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup dos Dados"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   ControlBox      =   0   'False
   Icon            =   "backupDados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar Barra 
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   2070
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox cpDrive 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   795
      Width           =   2445
   End
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "&Cancelar"
      Height          =   795
      Left            =   3900
      Picture         =   "backupDados.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton BtnProsseguir 
      Caption         =   "&Prosseguir"
      Height          =   795
      Left            =   2715
      Picture         =   "backupDados.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1140
   End
   Begin VB.OptionButton OpRestaurar 
      Caption         =   "&Restaurar"
      Height          =   195
      Left            =   2385
      TabIndex        =   2
      Top             =   285
      Width           =   1170
   End
   Begin VB.OptionButton OpBackup 
      Caption         =   "&Backup"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   255
      Value           =   -1  'True
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Percento Concluido:"
      Height          =   195
      Left            =   30
      TabIndex        =   7
      Top             =   1830
      Width           =   1440
   End
   Begin VB.Label Texto 
      AutoSize        =   -1  'True
      Caption         =   "Realizar o backup em"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   870
      Width           =   1545
   End
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EmProcesso As Boolean

Private Sub BtnCancelar_Click()
  If EmProcesso = False Then
    Unload Me
  End If
End Sub

Private Sub BtnProsseguir_Click()
On Error GoTo Errado
  Dim ResultCod As Integer

  BtnProsseguir.Enabled = False
  BtnCancelar.Enabled = False
  cpDrive.Enabled = False

  Call CloseDataBase

  If OpBackup.Value Then
    With ZipFiles
      .ClearDisks = True
      .FilesToProcess = Trim(Parametros.Dados) & "\DADOS.MDB"
      .ZipFileName = cpDrive.Text & "\backup.j_f"
      .MultidiskMode = True
      ResultCod = .Add(0)
      If ResultCod <> 0 Then
        MsgBox "Erro: " & ResultCod, vbCritical + vbOKOnly, Titulo
      Else
        MsgBox "Beckup completado com sucesso!", vbExclamation + vbOKOnly, Titulo
      End If
    End With
  Else
    Resp = MsgBox("Insira o último disquete do backup no drive: " & cpDrive.Text & ".", vbQuestion + vbOKCancel, Titulo)
    If Resp = vbOK Then
      With ZipFiles
        .Overwrite = 2
        .ExtractDirectory = False
        .UsePaths = False
        .MultidiskMode = True
        .ExtractDirectory = Trim(Parametros.Dados)
        .ZipFileName = cpDrive.Text & "\backup.j_f"
        .FilesToProcess = "*.*"
        DoEvents
        ResultCod = .Extract(0)

        If ResultCod <> 0 Then
          MsgBox "Erro: " & ResultCod, vbCritical + vbOKOnly, Titulo
        Else
          MsgBox "Restauração completada com sucesso!", vbExclamation + vbOKOnly, Titulo
        End If
      End With
    End If
  End If

  If daClientes Is Nothing Then Call AbrirArquivos(Trim(Parametros.Dados) & "\dados.mdb", 2)

Fim:
  BtnProsseguir.Enabled = True
  BtnCancelar.Enabled = True
  cpDrive.Enabled = True
  OpBackup.Enabled = True
  OpRestaurar.Enabled = True
  Barra.Value = 0
  Label1.Caption = "Percento Concluido:"
  Label1.Refresh
  Exit Sub

Errado:
  Call MsgBox(Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical + vbRetryCancel, Titulo)
  Resume Fim

End Sub

Private Sub Form_Load()
  Dim nDrive As Drive
  Me.Refresh
  DoEvents
  
  
  Me.KeyPreview = True
  EmProcesso = False
  For Each nDrive In SysFiles.Drives
    cpDrive.AddItem (nDrive.DriveLetter & ":")
  Next
  If cpDrive.ListCount > 0 Then cpDrive.ListIndex = 0
End Sub

Private Sub OpBackup_Click()
  Texto.Caption = "Ralizar o backup em"
  Texto.Refresh
End Sub

Private Sub OpRestaurar_Click()
  Texto.Caption = "Restaurar de"
  Texto.Refresh
End Sub

Private Sub ZipFiles_GlobalStatus(ByVal TotalFiles As Long, ByVal ProcessedFiles As Long, ByVal CompletionFiles As Integer, ByVal TotalBytes As Long, ByVal ProcessedBytes As Long, ByVal CompletionBytes As Integer)
  Label1.Caption = "Percento Concluido: " & Format(ProcessedBytes / TotalBytes * 100, "#0\%")
  Label1.Refresh
  Barra.Value = ProcessedBytes / TotalBytes * 100
End Sub

'Private Sub ZipFiles_DiskNotEmpty(xAction As XceedZipLibCtl.xcdNotEmptyAction)
'  Resp = MsgBox("Este disco contem arquivos. Deletar?", vbQuestion + vbYesNoCancel, Titulo)
'  If Resp = vbYes Then
'    xAction = xnaErase
'  ElseIf Resp = vbNo Then
'    xAction = xnaAskAnother
'  ElseIf Resp = vbCancel Then
'    xAction = xnaAbort
'  End If
'End Sub

'Private Sub ZipFiles_GlobalStatus(ByVal lFilesTotal As Long, ByVal lFilesProcessed As Long, ByVal lFilesSkipped As Long, ByVal nFilesPercent As Integer, ByVal lBytesTotal As Long, ByVal lBytesProcessed As Long, ByVal lBytesSkipped As Long, ByVal nBytesPercent As Integer, ByVal lBytesOutput As Long, ByVal nCompressionRatio As Integer)
'  Label1.Caption = "Percento Concluido: " & Format(nBytesPercent, "#0\%")
'  Label1.Refresh
'  Barra.Value = nBytesPercent
'End Sub


Private Sub ZipFiles_Newdisk(ByVal DiskNumber As Integer)
  If DiskNumber = 0 Then DiskNumber = 1
  Resp = MsgBox("Insira o disco número: " & DiskNumber & " no drive: " & cpDrive.Text & ".", vbQuestion + vbYesNo, Titulo)
  If Resp = vbYes Then
    'nada
  Else
    ZipFiles.Abort = True
  End If
End Sub

