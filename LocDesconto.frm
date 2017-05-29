VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form LocDesconto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de descontos"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "LocDesconto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   3360
      Width           =   435
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Aplicar filtro"
      Height          =   375
      Left            =   7380
      TabIndex        =   4
      Top             =   4140
      Width           =   1755
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
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
      MaxLength       =   13
      Text            =   "0,00"
      TextMask        =   4
      RawText         =   4
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   3780
      Width           =   915
      _ExtentX        =   1614
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
      MaxLength       =   7
      TextMask        =   9
      RawText         =   9
      Mask            =   "##/####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8520
      Top             =   1920
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LocDesconto.frx":000C
      Height          =   3195
      Left            =   60
      OleObjectBlob   =   "LocDesconto.frx":0020
      TabIndex        =   5
      Top             =   60
      Width           =   9135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1635
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   9
      Top             =   3360
      Width           =   6075
      _ExtentX        =   10716
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
      Top             =   3360
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   660
      TabIndex        =   8
      Top             =   4260
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mês/ano"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3420
      Width           =   855
   End
End
Attribute VB_Name = "LocDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbCondominio As Recordset

Private Sub cmdFiltrar_Click()
  Dim sSql As String
  If Val(cpCodigo.Text) = 0 And Not IsDate("25/" & cpMes.Text) And cpValor <= 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) > 0 And Not IsDate("25/" & cpMes.Text) And cpValor <= 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.ID_CONDOMINIO = " & cpCodigo.Text _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) > 0 And IsDate("25/" & cpMes.Text) And cpValor <= 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.ID_CONDOMINIO = " & cpCodigo.Text _
        & " and DESCONTOS.MES = '" & cpMes.Text & "' " _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) > 0 And IsDate("25/" & cpMes.Text) And cpValor > 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.ID_CONDOMINIO = " & cpCodigo.Text _
        & " and DESCONTOS.MES = '" & cpMes.Text & "' " _
        & "and DESCONTOS.VALOR = " & Replace(cpValor.Text, ",", ".") _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) = 0 And IsDate("25/" & cpMes.Text) And cpValor <= 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.MES = '" & cpMes.Text & "' " _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) = 0 And IsDate("25/" & cpMes.Text) And cpValor > 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.MES = '" & cpMes.Text & "' " _
        & "AND DESCONTOS.VALOR = " & Replace(cpValor.Text, ",", ".") _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) = 0 And Not IsDate("25/" & cpMes.Text) And cpValor > 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.VALOR = " & Replace(cpValor.Text, ",", ".") _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  ElseIf Val(cpCodigo.Text) > 0 And Not IsDate("25/" & cpMes.Text) And cpValor > 0# Then
    sSql = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
        & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
        & "WHERE DESCONTOS.ID_CONDOMINIO = " & cpCodigo.Text _
        & " AND DESCONTOS.VALOR = " & Replace(cpValor.Text, ",", ".") _
        & " ORDER BY [Associados].[Tipo], [Associados].[Apartamento];"
  End If
  Data1.RecordSource = sSql
  Data1.Refresh
  DBGrid1.ReBind
End Sub

Private Sub cpCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMes.SetFocus
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
    cpValor.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If IsDate("25/" & cpMes.Text) Then
    cpMes.Text = FormataMesAno(cpMes.Text)
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdFiltrar.SetFocus
  End If
End Sub

Private Sub DBGrid1_DblClick()
  If DBGrid1.Text <> "" Then
    RetCodigo = Data1.Recordset!id_desconto
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Data1.DatabaseName = Parametros.dados
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Data1.RecordSource = "SELECT [Associados].[Tipo]+'  '+[Associados].[Apartamento] AS nome, Associados.Proprietario, DESCONTOS.MES, DESCONTOS.HISTORICO, DESCONTOS.VALOR, DESCONTOS.ID_CONDOMINIO, DESCONTOS.ID_DESCONTO " _
      & "FROM Associados RIGHT JOIN DESCONTOS ON Associados.Codigo = DESCONTOS.ID_INQUILINO " _
      & "ORDER BY [Associados].[Tipo],[Associados].[Apartamento];"
  Data1.Refresh
  DBGrid1.ReBind
End Sub
