VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form ReAvulsa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpimir boleta avulsa"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "ReAvulsa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   7380
      TabIndex        =   2
      Top             =   4080
      Width           =   1275
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   4080
      Width           =   435
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ReAvulsa.frx":000C
      Height          =   3945
      Left            =   45
      OleObjectBlob   =   "ReAvulsa.frx":0020
      TabIndex        =   3
      Top             =   60
      Width           =   8745
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programas desenvolvidos\Predio\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1410
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   4
      Top             =   4080
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
      Left            =   1110
      TabIndex        =   0
      Top             =   4080
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   4155
      Width           =   855
   End
End
Attribute VB_Name = "ReAvulsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim tbCondominio As Recordset

Private Sub ImprimirBoleto()
  If IsNumeric(DBGrid1.Columns(1).Text) Then
    Data1.Recordset.Bookmark = DBGrid1.Bookmark
    With Data1.Recordset
      If !vcto < Date Then
        Resp = MsgBox("Este boleto já venceu, utilize reimpressão de boleto vencido. Imprimir assim mesmo?", vbQuestion + vbYesNo, "Aviso")
        If Resp = vbNo Then
          Exit Sub
        End If
      End If
    End With
    sSql = "Select * from boletos where cdsc = " & Data1.Recordset!cdsc & " and nosso = '" & Data1.Recordset!nosso & "' order by nome;"
    RelatoriosRPT.mnuSupExport.Visible = False
    RelatoriosRPT.mnuExportBoleto.Visible = True
    RelatoriosRPT.mnuEviarEmail.Visible = True
    RelatoriosRPT.Carregar "", Parametros.dados, "{boletos.cdsc} = " & Data1.Recordset!cdsc & " and {boletos.nosso} = '" _
        & Data1.Recordset!nosso & "'", "Boletos", sFormataCaminho(App.Path) & "avulso.rpt", , sSql
  End If
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

Private Sub cmdSelecionar_Click()
  
  If Val(cpCodigo.Text) = 0 Then
    Data1.RecordSource = "SELECT ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS snome, BLOCOS.NOME_BLOCO, * " _
      & "FROM BLOCOS RIGHT JOIN (ASSOCIADOS INNER JOIN Boletos ON ASSOCIADOS.CODIGO = Boletos.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO " _
      & "WHERE (((Left([TRAN],2))='AV') AND ((Boletos.PAGO)='N') AND ((Boletos.CANCELADO)='N')) " _
      & "ORDER BY ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO, Boletos.VCTO DESC;"
  Else
    Data1.RecordSource = "SELECT ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS snome, BLOCOS.NOME_BLOCO, * " _
      & "FROM BLOCOS RIGHT JOIN (ASSOCIADOS INNER JOIN Boletos ON ASSOCIADOS.CODIGO = Boletos.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO " _
      & "WHERE (((Left([TRAN],2))='AV') AND ((Boletos.PAGO)='N') AND ((Boletos.CANCELADO)='N') AND ((Boletos.COND)=" & Val(cpCodigo.Text) & ")) " _
      & "ORDER BY ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO, Boletos.VCTO DESC;"
  End If
  Data1.Refresh
  With Data1.Recordset
    If .RecordCount > 0 Then
      .MoveFirst
    End If
  End With
  
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
          cmdSelecionar.SetFocus
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
        cmdSelecionar.SetFocus
      End If
    End If
  End If
End Sub

Private Sub DBGrid1_DblClick()
  ImprimirBoleto
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    ImprimirBoleto
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
  KeyPreview = True
  Data1.DatabaseName = Parametros.dados
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
