VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form AbrirRetorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de retornos"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   ControlBox      =   0   'False
   Icon            =   "AbrirRetorno.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5940
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   315
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5940
      Width           =   1095
   End
   Begin rdActiveText.ActiveText cpData 
      Height          =   315
      Left            =   1860
      TabIndex        =   6
      Top             =   540
      Width           =   1395
      _ExtentX        =   2461
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
   Begin ComctlLib.ListView listaRetornos 
      Height          =   4815
      Left            =   60
      TabIndex        =   4
      Top             =   960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Condomínio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Data/Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sequencial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Arquivo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "Listar"
      Height          =   315
      Left            =   8580
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "..."
      Height          =   315
      Left            =   7860
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox cpDiretorio 
      Height          =   315
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "*** CONVÊNIO NÃO ENCONTRADO"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5940
      Width           =   2640
   End
   Begin VB.Label lbProgresso 
      BackColor       =   &H00FF0000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   9255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data mínima para listar"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Diretório de retornos"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1425
   End
End
Attribute VB_Name = "AbrirRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim nKey As Long
Dim iIndexFolder As Long
Dim Parar As Boolean

Private Sub cmdAbrir_Click()
  Dim sDir As String
  
  sDir = MinhaDll.sProcuraPorDiretorio("Diretório de retorno?", Me, False, False)
  
  If sDir <> "" Then
    cpDiretorio.Text = sDir
  End If
End Sub

Private Function ProcurarArquivos(sPathBegin As String, iIndexFolder As Long) As Boolean
On Error GoTo Errado
  Dim Fso     As Object
  Dim oRep    As Object
  Dim oSubRep As Object
  Dim oFolder As Folder 'Object
  Dim oFiles  As Object

  If Parar Then
    cmdListar.Caption = "Parar"
    cmdListar.Refresh
    GoTo Sair
  End If
  
  Set Fso = CreateObject("Scripting.FileSystemObject")
  Set oRep = Fso.GetFolder(sPathBegin)
  
  If iIndexFolder = 0 Then
    If oRep.Attributes <> 22 Then
      For Each oFiles In oRep.Files
        lbProgresso.Caption = AcertaLabel("Processando: " & oFiles.Path, "\")
        lbProgresso.Refresh
        If Right(oFiles.Path, 4) = ".ret" Then
          AcrescentaArquivo oFiles.Path
        End If
        DoEvents
        If Parar Then
          cmdListar.Caption = "Parar"
          cmdListar.Refresh
          Exit For
        End If
      Next oFiles
    End If
  End If
  
  For Each oFolder In oRep.SubFolders
    iIndexFolder = iIndexFolder + 1
    lbProgresso.Caption = AcertaLabel("Processando: " & oFolder.Path, "\")
    lbProgresso.Refresh
    If (InStr(oFolder.Path, "secur") = 0) And InStr(oFolder.Path, "segur") = 0 Then
      If oFolder.Attributes <> 22 Then
        For Each oFiles In oFolder.Files
          lbProgresso.Caption = AcertaLabel("Processando: " & oFiles.Path, "\")
          lbProgresso.Refresh
          If Right(oFiles.Path, 4) = ".ret" Then
            AcrescentaArquivo oFiles.Path
          End If
          DoEvents
          If Parar Then
            cmdListar.Caption = "Parar"
            cmdListar.Refresh
            Exit For
          End If
        Next oFiles
      End If
      DoEvents
      If Parar Then
        cmdListar.Caption = "Parar"
        cmdListar.Refresh
        Exit For
      End If
    End If
  Next oFolder
  
  DoEvents
  If Parar Then
    cmdListar.Caption = "Parar"
    cmdListar.Refresh
    Exit Function
  End If

  For Each oSubRep In oRep.SubFolders
     ProcurarArquivos oSubRep.Path, iIndexFolder
  Next oSubRep
  Set Fso = Nothing
  
Sair:
  Exit Function
Errado:
  MsgBox "Erro n. " & Err.Number & " - " & iIndexFolder & ":" & nKey & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Function

Private Sub AcrescentaArquivo(sArquivo As String)
  Dim txtDados() As String
  Dim sFile As Integer
  Dim sDetalhe As String
  Dim sCampo As String
  Dim nItem As ListItem
  Dim sData As String
  Dim sCond As String
  
  On Error Resume Next
  sFile = FreeFile
  
  Open sArquivo For Input As sFile
    txtDados = Split(Input(LOF(sFile), sFile), vbCrLf)
  Close sFile
  
  sDetalhe = txtDados(0)
  If Mid$(sDetalhe, 1, 3) <> "104" Then
    Exit Sub
  End If
  
  Erase txtDados
  
  sCampo = Mid$(sDetalhe, 144, 8)
  sCampo = Left$(sCampo, 2) & "/" & Mid$(sCampo, 3, 2) & "/" & Right$(sCampo, 4)
  sData = sCampo
  
  If CDate(sData) < CDate(cpData.Text) Then
    Exit Sub
  End If
  
  sCampo = Mid$(sDetalhe, 152, 6)
  sCampo = Left$(sCampo, 2) & ":" & Mid$(sCampo, 3, 2) & ":" & Right$(sCampo, 2)
  sData = sData & " " & sCampo
  sCampo = Mid$(sDetalhe, 158, 6)
  
  sCond = Mid$(sDetalhe, 59, 6)
  sCond = GetCondominioConvenio(sCond, sDetalhe)
  
  Set nItem = listaRetornos.ListItems.Add(, "a" & nKey)
  nItem.Text = sCond
  nItem.SubItems(1) = sData
  nItem.SubItems(2) = sCampo
  nItem.SubItems(3) = sArquivo
    
  nKey = nKey + 1
End Sub

Private Sub cmdCancelar_Click()
  RetornoCEF.cpArquivo.Text = ""
  RetornoCEF.KeyRemove = -1
  Unload Me
End Sub

Private Sub cmdListar_Click()
  
  If cmdListar.Caption = "Parar" Then
    Parar = True
    Exit Sub
  End If
  
  If Trim(cpDiretorio.Text) = "" Then
    MsgBox "Informe o diretório dos retornos.", vbCritical + vbOKOnly, "Aviso"
    cmdAbrir.SetFocus
    Exit Sub
  End If
    
  If Not sysFiles.FolderExists(cpDiretorio.Text) Then
    MsgBox "Diretório '" & cpDiretorio.Text & "' não existe.", vbCritical + vbOKOnly, "Aviso"
    cmdAbrir.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(cpData.Text) Then
    MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
    cpData.SetFocus
    Exit Sub
  End If
  
  If CDate(cpData.Text) > Date Then
    MsgBox "A data informada é maior que a data atual.", vbCritical + vbOKOnly, "Aviso"
    cpData.SetFocus
    Exit Sub
  End If
  
  lbProgresso.Visible = True
  lbProgresso.Caption = "Localizando arquivos de retorno..."
  lbProgresso.Refresh
  listaRetornos.Visible = False
  nKey = 1
  Parar = False
  listaRetornos.ListItems.Clear
  cmdListar.Caption = "Parar"
  cmdListar.Refresh
  ProcurarArquivos sFormataCaminho(cpDiretorio.Text), 0
  cmdListar.Caption = "Listar"
  cmdListar.Refresh
  Call AutoAjusteListView(listaRetornos, 0)
  lbProgresso.Visible = False
  listaRetornos.Visible = True
  Refresh
End Sub

Public Function GetCondominioConvenio(ByVal sConvenio As String, sLinha As String) As String
  Dim rs As Recordset
  Dim sRet As String
  Set rs = db.OpenRecordset("select * from condominio where right(conta,6) = '" & sConvenio & "';", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      If !titularboleto = 1 Then
        sRet = !Nome
      Else
        sRet = !razaoboleto
      End If
    Else
      If Trim(sLinha) <> "" Then
        sRet = Mid$(sLinha, 73, 30) & "***"
      Else
        sRet = "CONVÊNIO " & sConvenio & " NÃO ENCONTRADO"
      End If
    End If
  End With
  rs.Close
  Set rs = Nothing
  GetCondominioConvenio = sRet
End Function

Private Sub cmdOk_Click()
  With listaRetornos
    For nKey = 1 To .ListItems.Count
      If .ListItems(nKey).Selected Then
        RetornoCEF.cpArquivo.Text = LCase(.ListItems(nKey).SubItems(3))
        RetornoCEF.KeyRemove = nKey
        Exit For
      End If
    Next nKey
  End With
  Me.Hide
End Sub

Private Sub cpData_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdListar.SetFocus
  End If
End Sub

Private Sub Form_Load()
  cpData.Text = DateAdd("d", -3, Date)
  cpDiretorio.Text = MinhaDll.Le_dados_ini("Retorno", "Local", Inifile)
  Refresh
  KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MinhaDll.Grava_dados_ini "Retorno", "Local", cpDiretorio.Text, Inifile
End Sub
