VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{CCF682E3-14F1-11D7-80E2-10E08C102140}#1.0#0"; "zprogbar.ocx"
Begin VB.Form CanBol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelamento de boletos"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "CanBol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin ZealProgressBar.ProgressBar Barra 
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
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
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Processar"
      Height          =   405
      Left            =   5760
      TabIndex        =   3
      Top             =   540
      Width           =   1425
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   540
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   7
      TextMask        =   9
      RawText         =   9
      Mask            =   "##/####"
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   7
      Top             =   60
      Width           =   4815
      _ExtentX        =   8493
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
      MaxLength       =   50
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Alignment       =   1
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
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Boletos do mês"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "CanBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelBol As Recordset
Dim tbCondominio As Recordset

Private Sub cmdConfirma_Click()
On Error GoTo Errado
  Dim nCod As Long
  Dim status As Boolean
  
  If Not IsDate("25/" & cpMes.Text) Then
    MsgBox "O mês/ano informado não é válido.", vbInformation + vbOKOnly, "Aviso"
    cpMes.SetFocus
    Exit Sub
  End If
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  Resp = MsgBox("Confirma o cancelamento dos boletos do mês '" & cpMes.Text & "'?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir")
  If Resp = vbYes Then
    nCod = cpCodigo.Text
    Set SelBol = db.OpenRecordset("Select * From Boletos Where Tran = '" & cpMes.Text & "' and cond = " & nCod & ";", dbOpenDynaset)
    With SelBol
      If .RecordCount > 0 Then
        .MoveLast
        .MoveFirst
        Do While Not .EOF
          If !pago = "S" Then
            MsgBox "Encontrei boletos do mês '" & cpMes.Text & "' já quitados. Não posso cancelar!", vbInformation + vbOKOnly, "Aviso"
            Exit Sub
          End If
          .MoveNext
        Loop
        DBEngine.BeginTrans
        .MoveFirst
        status = False
        Do While Not .EOF
          If (!idStatus = 2 Or !idStatus = 3 Or !idStatus = 4 Or !idStatus = 6 Or !idStatus = 7 Or !idStatus = 9) Then
            status = True
            Exit Do
          End If
          If Not .EOF Then
            Barra.Value = .PercentPosition
          End If
          .MoveNext
        Loop
        If status Then
          .MoveFirst
          Do While Not .EOF
            .Edit
            !idStatus = 9
            !CANCELADO = "S"
            .Update
            If Not .EOF Then
              Barra.Value = .PercentPosition
            End If
            .MoveNext
          Loop
        Else
          .MoveFirst
          Do While Not .EOF
            db.Execute "delete from boletodetalhe where id_boleto = " & !id & ";"
            If Not .EOF Then
              Barra.Value = .PercentPosition
            End If
            .MoveNext
          Loop
          db.Execute "Delete From Boletos Where TRAN = '" & cpMes.Text & "' and cond =" & nCod & ";"
          MsgBox "Boletos do mês '" & cpMes.Text & "' cancelados!", vbInformation + vbOKOnly, "Aviso"
        End If
        DBEngine.CommitTrans
      Else
        MsgBox "Não encontrei boletos do mês '" & cpMes.Text & "'.", vbInformation + vbOKOnly, "Aviso"
      End If
    End With
    Barra.Value = 100
  End If
Sair:
  Exit Sub
Errado:
  MsgBox "Erro " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  DBEngine.Rollback
  Resume Sair
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpNome.Text = !Nome
        cpCodigo.Text = !Codigo
      End If
    End With
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodigo.Text
        If Not .NoMatch Then
          cpNome.Text = !Nome
          cpMes.SetFocus
        Else
          MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
          cpCodigo.SetFocus
        End If
      End With
    Else
      RetCodigo = 0
      l_Condominio.Show 1
      If RetCodigo > 0 Then
        With tbCondominio
          .Index = "codigoid"
          .Seek "=", RetCodigo
          If Not .NoMatch Then
            cpNome.Text = !Nome
            cpCodigo.Text = !Codigo
          End If
        End With
        cpMes.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdConfirma.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
