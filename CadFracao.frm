VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CadFracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fração"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "CadFracao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3060
      Top             =   2280
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CadFracao.frx":000C
      Height          =   2715
      Left            =   60
      OleObjectBlob   =   "CadFracao.frx":0020
      TabIndex        =   0
      Top             =   60
      Width           =   5895
   End
   Begin VB.Data Data1 
      Caption         =   "Fração"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2100
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "CadFracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lCod As Long

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If ColIndex = 2 Then
    DBGrid1.Columns(1).Text = Format$(Date, "dd/MM/yyyy")
    DBGrid1.Columns(3).Text = lCod
    If Trim(DBGrid1.Columns(4).Text & "") = "" Then
      DBGrid1.Columns(4).Text = "PRINCIPAL"
    End If
  End If
End Sub

Private Sub DBGrid1_BeforeInsert(Cancel As Integer)
  If DBGrid1.ApproxCount >= 10 Then
    MsgBox "Número máximo de frações por inquilino!", vbCritical + vbOKOnly, "Aviso"
    Cancel = True
  End If
End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
  If DBGrid1.AddNewMode > 0 Then
    If JaExisteFracao(lCod, DBGrid1.Columns(4).Text) Then
      MsgBox "Já existe uma fração com mesma descrição para este inquilino!", vbCritical + vbOKOnly, "Aviso"
      Cancel = True
    End If
  End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If DBGrid1.Col = 4 Then
    KeyAscii = vTexto(KeyAscii)
    If Len(DBGrid1.Columns(4).Text) >= 20 Then
      KeyAscii = 0
    End If
  End If
  If DBGrid1.Col = 2 Then
    If KeyAscii <> Asc(",") Then
      KeyAscii = vNumero(KeyAscii)
    End If
  End If
End Sub

Private Sub Form_Load()
  KeyPreview = True
  Data1.DatabaseName = Parametros.Dados
  Timer1.Enabled = True
  Refresh
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Data1.RecordSource = "select * from fracao where id_associado = " & lCod & " order by data_cadastro;"
  Data1.Refresh
End Sub

Private Function JaExisteFracao(ByVal nCod As Long, ByVal sNome As String) As Boolean
  Dim rs As Recordset
  Dim ret As Boolean
  
  ret = False
  
  Set rs = db.OpenRecordset("select * from fracao where id_associado = " & nCod & " and descricao = '" & sNome & "';", dbOpenDynaset)
  If rs.RecordCount > 0 Then
    ret = True
  End If
  Set rs = Nothing
  JaExisteFracao = ret
End Function
