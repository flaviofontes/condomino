VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBoletos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de boletos"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11250
   Icon            =   "frmBoletos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cpData 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5760
      Width           =   2595
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   9540
      TabIndex        =   10
      Top             =   4500
      Width           =   1635
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Visualizar"
      Height          =   375
      Left            =   9540
      TabIndex        =   9
      Top             =   5280
      Width           =   1635
   End
   Begin VB.ComboBox cpSituacao 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5340
      Width           =   2595
   End
   Begin VB.ComboBox cpTipo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4920
      Width           =   2595
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   4500
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8880
      Top             =   2580
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmBoletos.frx":000C
      Height          =   4335
      Left            =   60
      OleObjectBlob   =   "frmBoletos.frx":0020
      TabIndex        =   0
      Top             =   60
      Width           =   11115
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   2
      Top             =   4500
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
      Left            =   1050
      TabIndex        =   3
      Top             =   4500
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSComCtl2.DTPicker cpDataUsada 
      Height          =   315
      Left            =   3900
      TabIndex        =   13
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   57671683
      CurrentDate     =   42171
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5820
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Situação"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4980
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4575
      Width           =   855
   End
End
Attribute VB_Name = "frmBoletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbCondominio As Recordset

Private Sub cmdFiltrar_Click()
  AplicarFiltro
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

Private Sub cmdVisualizar_Click()
  Dim PerSind As Double
  Dim sSql As String
  With DBGrid1
    If .ApproxCount > 0 Then
      If Trim(.Columns(0).Text) <> "" Then
        Data1.Recordset.Bookmark = .Bookmark
        If IsNumeric(Left(Trim(.Columns(0).Text), 2)) Then
          sSql = "Select * from boletos where tran = '" & Data1.Recordset!tran & "' and cond = " & Data1.Recordset!cond & " order by bole;"
          PerSind = PersentualSindico(CodigoSindico(Data1.Recordset!cond))
          RelatoriosRPT.mnuSupExport.Visible = False
          RelatoriosRPT.mnuExportBoleto.Visible = True
          RelatoriosRPT.mnuOrdenar.Visible = True
          RelatoriosRPT.mnuEviarEmail.Visible = True
          RelatoriosRPT.Carregar "{boletos.bole};crAscendingOrder;boletos|", Parametros.dados, "{boletos.bole} = '" & Data1.Recordset!bole & "' and {boletos.cond} = " _
              & Data1.Recordset!cond, "Boletos", sFormataCaminho(App.Path) & "cobranca.rpt", , sSql, "persind|" & PerSind, , Data1.Recordset!cond, Data1.Recordset!tran
        ElseIf Left(Trim(.Columns(0).Text), 2) = "GE" Then
          sSql = "Select * from boletos where left(tran,2) = 'GE' and cond = " & Data1.Recordset!cond & " and bole = '" & Data1.Recordset!bole & "' order by bole;"
          PerSind = PersentualSindico(CodigoSindico(Data1.Recordset!cond))
          RelatoriosRPT.mnuSupExport.Visible = False
          RelatoriosRPT.mnuExportBoleto.Visible = True
          RelatoriosRPT.mnuOrdenar.Visible = True
          RelatoriosRPT.mnuEviarEmail.Visible = True
          RelatoriosRPT.Carregar "{boletos.bole};crAscendingOrder;boletos|", Parametros.dados, "{boletos.historico}= '" & Data1.Recordset!Historico & "' and Left({boletos.tran},2) = 'GE' and {boletos.cond} = " _
              & Data1.Recordset!cond & " AND {BOLETOS.BOLE}='" & Data1.Recordset!bole _
              & "'", "Boletos", sFormataCaminho(App.Path) & "generico.rpt", , sSql, "persind|" & PerSind, , Data1.Recordset!cond, "0000"
        Else
          sSql = "Select * from boletos where cdsc = " & Data1.Recordset!cdsc & " and nosso = '" & Data1.Recordset!nosso & "' order by nome;"
          RelatoriosRPT.mnuSupExport.Visible = False
          RelatoriosRPT.mnuExportBoleto.Visible = True
          RelatoriosRPT.mnuEviarEmail.Visible = True
          RelatoriosRPT.Carregar "", Parametros.dados, "{boletos.cdsc} = " & Data1.Recordset!cdsc & " and {boletos.nosso} = '" _
              & Data1.Recordset!nosso & "'", "Boletos", sFormataCaminho(App.Path) & "avulso.rpt", , sSql
        End If
      End If
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
      End If
    End If
  End If
End Sub

Private Sub AplicarFiltro()
  Dim Sql As String
  Dim st As String
  Dim ss As String
  Dim sc As String
  Dim dt As String
  Dim sWhere As String
  
  Select Case cpTipo.ListIndex
    Case 0
      st = " "
    Case 1
      st = " isnumeric(Left([BOLETOS].[TRAN],2)) "
    Case 2
      st = " Left([BOLETOS].[TRAN],2)='AV' "
    Case 3
      st = " Left([BOLETOS].[TRAN],2)='GE' "
    Case 4
      st = " Left([BOLETOS].[TRAN],2)='PR' "
  End Select
  
  Select Case cpSituacao.ListIndex
    Case 0
      ss = " "
    Case 1
      ss = " BOLETOS.PAGO='S' "
    Case 2
      ss = " BOLETOS.PAGO='N' "
  End Select
  
  Select Case cpData.ListIndex
    Case 0
      dt = " "
    Case 1
      dt = " AND [BOLETOS].[VCTO] <= #" & Format$(cpDataUsada.Value, "MM/dd/yyyy") & "# "
    Case 2
      dt = " AND [BOLETOS].[VCTO] >= #" & Format$(cpDataUsada.Value, "MM/dd/yyyy") & "# "
    Case 3
      dt = " AND [BOLETOS].[DTPGTO] <= #" & Format$(cpDataUsada.Value, "MM/dd/yyyy") & "# "
    Case 4
      dt = " AND [BOLETOS].[DTPGTO] >= #" & Format$(cpDataUsada.Value, "MM/dd/yyyy") & "# "
  End Select
  
  If Val(cpCodigo.Text) > 0 Then
    sc = " BOLETOS.COND = " & cpCodigo.Text & " "
  Else
    sc = " "
  End If
  
  If Trim(sc) <> "" And Trim(st) <> "" And Trim(ss) <> "" Then
    sWhere = "WHERE " + sc + " AND " + st + " AND " + ss
  ElseIf Trim(sc) <> "" And Trim(st) <> "" And Trim(ss) = "" Then
    sWhere = "WHERE " + sc + " AND " + st
  ElseIf Trim(sc) <> "" And Trim(st) = "" And Trim(ss) = "" Then
    sWhere = "WHERE " + sc
  ElseIf Trim(sc) <> "" And Trim(st) = "" And Trim(ss) <> "" Then
    sWhere = "WHERE " + sc + " AND " + ss
  ElseIf Trim(sc) = "" And Trim(st) <> "" And Trim(ss) <> "" Then
    sWhere = "WHERE " + st + " AND " + ss
  ElseIf Trim(sc) = "" And Trim(st) <> "" And Trim(ss) = "" Then
    sWhere = "WHERE " + st
  ElseIf Trim(sc) = "" And Trim(st) = "" And Trim(ss) <> "" Then
    sWhere = "WHERE " + ss
  ElseIf Trim(sc) = "" And Trim(st) = "" And Trim(ss) = "" Then
    sWhere = " "
  End If
  
  Sql = "SELECT BOLETOS.*, BLOCOS.NOME_BLOCO+' '+ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS NOMEC, CONDOMINIO.NOME " _
    + "FROM CONDOMINIO RIGHT JOIN (BLOCOS RIGHT JOIN (ASSOCIADOS RIGHT JOIN BOLETOS ON ASSOCIADOS.CODIGO = BOLETOS.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO) ON CONDOMINIO.CODIGO = BOLETOS.COND " _
    + sWhere & " AND BOLETOS.CANCELADO = 'N' " & dt _
    + "ORDER BY BOLETOS.COND, BOLETOS.VCTO;"
  
  Data1.RecordSource = Sql
  Data1.Refresh

End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Data1.DatabaseName = Parametros.dados
  cpTipo.AddItem ""
  cpTipo.AddItem "Normal"
  cpTipo.AddItem "Avulso"
  cpTipo.AddItem "Genérico"
  cpTipo.AddItem "Parcelamento"
  cpTipo.ListIndex = 0
  cpSituacao.AddItem ""
  cpSituacao.AddItem "Pago"
  cpSituacao.AddItem "Em aberto"
  cpSituacao.ListIndex = 0
  cpData.AddItem ("")
  cpData.AddItem ("Vencidas antes de")
  cpData.AddItem ("Vencidas depois de")
  cpData.AddItem ("Pagas antes de")
  cpData.AddItem ("Pagas depois de")
  cpDataUsada.Value = Date
  Refresh
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Call SetTopMostWindow(Progre.hWnd, True)
  Progre.Label1.Caption = "Carregando. Aguarde..."
  Progre.Show
  Progre.Refresh
  DoEvents
  Data1.RecordSource = "SELECT BOLETOS.*, BLOCOS.NOME_BLOCO+' '+ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO AS NOMEC, CONDOMINIO.NOME " _
    + "FROM CONDOMINIO RIGHT JOIN (BLOCOS RIGHT JOIN (ASSOCIADOS RIGHT JOIN BOLETOS ON ASSOCIADOS.CODIGO = BOLETOS.CDSC) ON BLOCOS.ID_BLOCO = ASSOCIADOS.ID_BLOCO) ON CONDOMINIO.CODIGO = BOLETOS.COND " _
    + "WHERE (((BOLETOS.CANCELADO)='N')) ORDER BY BOLETOS.COND, BOLETOS.VCTO;"
  Data1.Refresh
  Call SetTopMostWindow(Progre.hWnd, False)
  Unload Progre
End Sub
