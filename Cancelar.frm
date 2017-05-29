VERSION 5.00
Begin VB.Form Cancelamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelamento de Recibo"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "Cancelamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   780
      Left            =   4095
      Picture         =   "Cancelamento.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1845
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Recibo"
      Height          =   1155
      Left            =   30
      TabIndex        =   2
      Top             =   630
      Width           =   5190
      Begin VB.TextBox cpValor 
         Alignment       =   1  'Right Justify
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
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   975
      End
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
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   4020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   540
         TabIndex        =   7
         Top             =   735
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sócio"
         Height          =   195
         Left            =   495
         TabIndex        =   6
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.TextBox cpRecibo 
      Alignment       =   1  'Right Justify
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
      Left            =   1545
      MaxLength       =   8
      TabIndex        =   0
      Top             =   225
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do Recibo"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   285
      Width           =   1335
   End
End
Attribute VB_Name = "Cancelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelDetRec As Recordset
Dim vaCodigo  As Long

Private Sub cmdOk_Click()
On Error GoTo Errado
  
  If cpSocio.Text = "" Then
    MsgBox "Favor Selecionar um Recibo.", vbInformation + vbOKOnly, Titulo
    GoTo Fim
  End If
  
  cmdOk.Enabled = False
  Set SelDetRec = db.OpenRecordset("Select * From DetalheRecibo Where NumeroRecibo = " & cpRecibo, dbOpenDynaset)
  DoEvents
  daMensalidade.Index = "CTReferencia"
  
  With SelDetRec
    If .RecordCount > 0 Then
      .MoveLast
      DoEvents
      .MoveFirst
      Do While Not .EOF
        If !Tipo = "R" Then
'          daReceber.Seek "=", Val(!Codigo)
'          If Not daReceber.NoMatch Then
'            daReceber.Edit
'            daReceber!Quitado = "N"
'            daReceber!DataDoPagamento = Null
'            daReceber!FormaDoPagamento = ""
'            daReceber.Update
'          End If
        Else
          daMensalidade.Seek "=", !Codigo
          If Not daMensalidade.NoMatch Then
            Do While Not daMensalidade.EOF
              If daMensalidade!Referencia = !Codigo Then
                If daMensalidade!CodigoDoCliente = vaCodigo Then
                  If Left(!Descricao, 24) = "Pagamento Antecipado ate" Then
                    daMensalidade.Delete
                  Else
                    daMensalidade.Edit
                    daMensalidade!Situacao = Null
                    daMensalidade!DataPagamento = Null
                    daMensalidade.Update
                  End If
                  Exit Do
                End If
              End If
              daMensalidade.MoveNext
            Loop
          End If
        End If
        .MoveNext
      Loop
    End If
  End With
  
  With daRecibo
    .Edit
    !Cancelado = "SIM"
    .Update
  End With
  Limpar

Fim:
  cmdOk.Enabled = True
  Exit Sub

Errado:
  MsgBox Err.Number & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly, Titulo
  Resume Fim

End Sub

Private Sub cpRecibo_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      KeyAscii = KeyAscii
    Case 13
      KeyAscii = 0
      cmdOk.SetFocus
    Case 48 To 57
      KeyAscii = KeyAscii
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub cpRecibo_LostFocus()
  If Len(cpRecibo) > 0 Then
    With daRecibo
      .Index = "RecNumero"
      .Seek "=", Val(cpRecibo.Text)
      If Not .NoMatch Then
        cpSocio.Text = !NomeCliente
        cpValor.Text = Format(!Valor, "#,##0.00")
        vaCodigo = !CodigoCliente
      Else
        MsgBox "Este Recibo não Consta no Cadastro! Verifique Por Favor.", vbInformation + vbOKOnly, Titulo
      End If
    End With
  Else
    LocRecibo.Show 1
    If RetCodigo > 0 Then
      With daRecibo
        .Index = "RecNumero"
        .Seek "=", RetCodigo
        If Not .NoMatch Then
          cpSocio.Text = !NomeCliente
          cpValor.Text = Format(!Valor, "#,##0.00")
          cpRecibo.Text = !Numero
          vaCodigo = !CodigoCliente
        End If
      End With
    End If
  End If
End Sub

Private Sub Limpar()
  cpSocio.Text = ""
  cpValor.Text = ""
  cpRecibo.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then KeyAscii = 0: Unload Me
End Sub

Private Sub Form_Load()
  Me.Refresh
  DoEvents
  Me.KeyPreview = True
End Sub
