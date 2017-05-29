VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form LocMensalidade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boletos em aberto"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   ControlBox      =   0   'False
   Icon            =   "LocMensalidade.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin MSComCtl2.DTPicker cpVencimento 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   6180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   16580611
      CurrentDate     =   41822
   End
   Begin VB.CheckBox chJuros 
      Caption         =   "Corrigido"
      Height          =   195
      Left            =   2580
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin ComctlLib.ListView lsDebitos 
      Height          =   2955
      Left            =   180
      TabIndex        =   3
      Top             =   3120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5212
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Boleto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vencimento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin rdActiveText.ActiveText cpTotal 
      Height          =   315
      Left            =   5700
      TabIndex        =   10
      Top             =   6180
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   4
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Programas desenvolvidos\Vicosa\Recanto\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2820
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LocMensalidade.frx":000C
      Height          =   1815
      Left            =   180
      OleObjectBlob   =   "LocMensalidade.frx":0020
      TabIndex        =   2
      Top             =   960
      Width           =   6975
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   13
      Top             =   480
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
      Top             =   480
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
      TabIndex        =   14
      Top             =   555
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   6240
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor total selecionado"
      Height          =   195
      Left            =   4020
      TabIndex        =   11
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Débitos "
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o condomínio, de duplo clique sobre o nome e depois esolha o(s) débito(s)."
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   120
      Width           =   6045
   End
End
Attribute VB_Name = "LocMensalidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nItem As ListItem
Dim rs As Recordset
Dim iCod As Long
Dim i As Integer
Dim dValor As Double
Dim tbCondominio As Recordset

Private Sub chJuros_Click()
  DBGrid1_DblClick
End Sub

Private Sub cmdCancelar_Click()
  BoletosAvulso.DebAnt = True
  Unload Me
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
        SelecionaInquilinos
      End If
    End With
  End If
End Sub

Private Sub cmdOk_Click()
  Dim sHistorico As String
  sHistorico = ""
  BoletosAvulso.DebAnt = True
  If cpTotal.Text > 0# Then
    For i = 1 To lsDebitos.ListItems.Count
      If lsDebitos.ListItems(i).Selected Then
        ReDim Preserve mRetDeb(i - 1)
        mRetDeb(i - 1) = Val(lsDebitos.ListItems(i).SubItems(3))
        sHistorico = sHistorico & lsDebitos.ListItems(i).Text & " de " & lsDebitos.ListItems(i).SubItems(1) & " "
      End If
    Next i
    With BoletosAvulso
      .cpVenc.Text = Format$(cpVencimento.Value, "dd/MM/yy")
      .cpVenc.Locked = True
      .DebAnt = True
      .cpValor.Text = cpTotal.Text
      .cpValor.Locked = True
      .cpCodigo.Text = Val(DBGrid1.Columns(0).Text)
      .cpCodigo.Locked = True
      .cpNome.Text = DBGrid1.Columns(1).Text
      .cpNome.Locked = True
      .cpHistorico.Text = "Boleto referente a " & Trim(sHistorico)
      .cpHistorico.Locked = True
      .cmdLocalizar.Enabled = False
    End With
  End If
  Unload Me
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
          DBGrid1.SetFocus
          SelecionaInquilinos
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
            SelecionaInquilinos
          End If
        End With
        DBGrid1.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpVencimento_Change()
  DBGrid1_DblClick
End Sub

Private Sub DBGrid1_DblClick()
  If DBGrid1.Columns(0).Text <> "" Then
    Me.MousePointer = 11
    DoEvents
    lsDebitos.ListItems.Clear
    iCod = DBGrid1.Columns(0).Text
    Set rs = db.OpenRecordset("Select * from boletos where cdsc = " & iCod & " and (PAGO <> 'S' or PAGO is null) order by vcto;", dbOpenDynaset)
    With rs
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          If chJuros.Value = 1 Then
            .Edit
            !corrigido = Reajustar(!cond, !valr, cpVencimento.Value, !vcto)
            .Update
          End If
          Set nItem = lsDebitos.ListItems.Add()
          nItem.Text = !tran & ""
          nItem.SubItems(1) = Format$(!vcto, "dd/MM/yyyy")
          If chJuros.Value = 1 Then
            nItem.SubItems(2) = Format$(!corrigido, "#,##0.00")
          Else
            nItem.SubItems(2) = Format$(!valr, "#,##0.00")
          End If
          nItem.SubItems(3) = !id & ""
          .MoveNext
        Loop
      End If
    End With
    Set rs = Nothing
    Me.MousePointer = 0
  End If
End Sub

Private Sub SelecionaInquilinos()
  Data1.RecordSource = "SELECT ASSOCIADOS.CODIGO, ASSOCIADOS.TIPO+' '+ASSOCIADOS.APARTAMENTO+' '+ASSOCIADOS.PROPRIETARIO AS MNOME from Associados " _
    & " where CONDOMINIO = " & cpCodigo.Text & " ORDER BY ASSOCIADOS.TIPO, ASSOCIADOS.APARTAMENTO;"
  Data1.Refresh
  DBGrid1.ReBind
  DBGrid1.ClearSelCols
End Sub

Private Sub Form_Load()
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Data1.DatabaseName = Parametros.dados
  cpVencimento.Value = Date
  KeyPreview = True
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
  Set rs = Nothing
End Sub

Private Sub lsDebitos_Click()
  dValor = 0
  If lsDebitos.ListItems.Count > 0 Then
    For i = 1 To lsDebitos.ListItems.Count
      If lsDebitos.ListItems(i).Selected Then
        dValor = dValor + CDbl(lsDebitos.ListItems(i).SubItems(2))
      End If
    Next i
  End If
  cpTotal.Text = dValor
End Sub

Public Function Reajustar(ByVal lCod As Long, _
                          ByVal oldValor As Double, _
                          ByVal NovoVenc As Date, _
                          ByVal oldVenc As Date) As Double
On Error GoTo Errado
  
  Dim dData     As Date
  Dim nDias     As Long
  Dim nJuros    As Single
  Dim valor     As String
  Dim Carencia  As Double
  Dim juros     As Double
  Dim Correcao  As Double
  Dim Multa     As Double
  Dim dbAnterior As Double
  
  Carencia = 30
  dData = NovoVenc
  dbAnterior = 0
  
  With tbCondominio
    .Index = "codigoid"
    .Seek "=", lCod
    If Not .NoMatch Then
      juros = !juros
      Multa = !Multa
      nJuros = juros / 30
    Else
      juros = Parametros.juros
      Multa = Parametros.Multa
      nJuros = juros / 30
    End If
  End With
  
  nDias = NovoVenc - oldVenc
  If nDias > 0 Then
    Correcao = Round(oldValor * (Multa / 100), 2)
    valor = Round(oldValor * (nJuros * nDias) / 100, 2)
    dbAnterior = oldValor + Correcao + valor
  Else
    dbAnterior = oldValor
  End If
  
Fim:
  Reajustar = dbAnterior
  Exit Function
  
Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Descrição: " & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim
  
End Function


