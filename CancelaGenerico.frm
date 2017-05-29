VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{CCF682E3-14F1-11D7-80E2-10E08C102140}#1.0#0"; "ZProgBar.ocx"
Begin VB.Form CancelaGenerico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar boleto genérico"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   Icon            =   "CancelaGenerico.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   Begin ZealProgressBar.ProgressBar Barra 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   1140
      Width           =   7275
      _ExtentX        =   12832
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
      Value           =   100
   End
   Begin VB.CommandButton cmdEtiquetas 
      Caption         =   "Cancelar"
      Height          =   795
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Width           =   1260
   End
   Begin VB.ComboBox cpHistoricoPrint 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   660
      Width           =   7275
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   2115
      TabIndex        =   1
      Top             =   180
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2595
      TabIndex        =   4
      Top             =   180
      Width           =   5955
      _ExtentX        =   10504
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
      Left            =   1320
      TabIndex        =   0
      Top             =   180
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   255
      Width           =   855
   End
End
Attribute VB_Name = "CancelaGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim rsSel As Recordset
Dim tbCondominio As Recordset

Private Sub cmdEtiquetas_Click()
  If Val(cpCodigo.Text) = 0 Then
    MsgBox "Selecione um condomínio.", vbCritical + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If cpHistoricoPrint.Text = "" Then
    MsgBox "Selecione um histórico.", vbCritical + vbOKOnly, "Aviso"
    cpHistoricoPrint.SetFocus
    Exit Sub
  End If
  
  Me.MousePointer = 11
  cmdEtiquetas.Enabled = False
  
  Resp = MsgBox("Tem certeza que deseja cancelar estes boletos?", vbQuestion + vbYesNo, "Cancelar")
  If Resp = vbYes Then
    Set rsSel = db.OpenRecordset("select id from boletos where historico = '" & cpHistoricoPrint.Text & "' and cond = " & cpCodigo.Text & ";", dbOpenDynaset)
    With rsSel
      If .RecordCount > 0 Then
        .MoveLast
        Do While Not .BOF
          If (!idStatus = 1 Or !idStatus = 5 Or !idStatus = 8) Then
            db.Execute "delete from boletodetalhe where id_boleto = " & !id & ";"
            .Delete
          Else
            .Edit
            !idStatus = 9
            .Update
          End If
          .MovePrevious
          If Not .BOF Then
            Barra.Value = Int(.PercentPosition)
          End If
        Loop
      Else
        MsgBox "Nenhum boleto encontrado com os dados informados ou com status que permitem canelamento.", vbInformation + vbOKOnly, "Aviso"
      End If
    End With
  End If

  Barra.Value = 0
  Seleciona
  Me.MousePointer = 0
  cmdEtiquetas.Enabled = True
  
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
          Seleciona
          cpHistoricoPrint.SetFocus
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
        Seleciona
        cpHistoricoPrint.SetFocus
      End If
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
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
  KeyPreview = True
End Sub

Private Sub Seleciona()
  
  If Val(cpCodigo.Text) = 0 Then
    MsgBox "Selecione um condomínio.", vbCritical + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  cpHistoricoPrint.Clear
  Set rsSel = db.OpenRecordset("select distinct historico from boletos where left(tran,2) = 'GE' and cond = " & cpCodigo.Text & ";", dbOpenDynaset)
  With rsSel
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cpHistoricoPrint.AddItem (!Historico & "")
        .MoveNext
      Loop
      cpHistoricoPrint.ListIndex = 0
    End If
  End With
  Set rsSel = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
