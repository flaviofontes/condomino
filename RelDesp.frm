VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelDesp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório da distribuição das despesas"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "RelDesp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   120
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
      Left            =   1140
      TabIndex        =   0
      Top             =   120
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
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   615
      Width           =   1050
      _ExtentX        =   1852
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Distribuir"
      Height          =   360
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   660
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Do mês"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   705
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   195
      Width           =   855
   End
End
Attribute VB_Name = "RelDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim TodosCond As Recordset
Private tbDespFixa  As Recordset
Dim Assoc As Recordset
Dim tbCondominio As Recordset
Dim tbAssociados As Recordset

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

Private Sub cmdPrint_Click()
  Dim ViraFundo As Long
  Dim DescFracao As String
  Dim rs As Recordset
  Dim FracoesNao As String
  Dim i As Integer
  Dim rs1 As Recordset
  Dim Fundos() As Fixas
  Dim Sql As String
  Dim totalDesp As Double
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If Not IsDate("25/" & cpMes.Text) Then
    MsgBox "O mês/ano informado não é válido.", vbInformation + vbOKOnly, "Aviso"
    cpMes.SetFocus
    Exit Sub
  End If
  
  cmdPrint.Enabled = False
  
  Set TodosCond = db.OpenRecordset("select * from condominio where codigo = " & cpCodigo.Text & ";", dbOpenDynaset)
  With TodosCond
    If .RecordCount > 0 Then
      .MoveFirst
      If cmdPrint.Caption = "Imprimir" Then
        PrintTela !Codigo, !geral
      Else
        ViraFundo = ViraFundoReserva(!Codigo)
        DescFracao = FracaoVira(ViraFundo)
        If ViraFundo > 0 Then
          db.Execute "delete from despesafixa where condominio = " & !Codigo & " and fracao = -1;"
          db.Execute "Delete From DespesaDistribuida where mes = '" & cpMes.Text & "' and id_condominio = " & !Codigo & ";"
          db.Execute "Delete From DespesaInquilino where mes = '" & cpMes.Text & "' and id_condominio = " & !Codigo & ";"
          
          'buscar nas despesas fixas se há algum por percentual sobre despesas
          Sql = "Select * From Despesas Where Condominio = " & !Codigo & " And Mes = '" & cpMes.Text & "' and desconto = 0 Order By Tipo;"
          Set rs1 = db.OpenRecordset(Sql, dbOpenDynaset)
          totalDesp = 0
          With rs1
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                If !Percentual <> 1 Then
                  totalDesp = totalDesp + !valor
                End If
                .MoveNext
              Loop
            End If
          End With
          
          Set rs1 = Nothing
          'fundo
          Set rs1 = db.OpenRecordset("select * from despesafixa where condominio = " & cpCodigo.Text _
              & " and fracao = 2;", dbOpenDynaset)
          i = 0
          With rs1
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                ReDim Preserve Fundos(i)
                Fundos(i).Descricao = !Historico
                Fundos(i).Percentual = !valor
                Fundos(i).valor = !valor * totalDesp / 100
                Fundos(i).Sindico = !Sindico
                Fundos(i).Fracao = !desc_fracao
                i = i + 1
                .MoveNext
              Loop
            End If
          End With
          Set rs1 = Nothing
          If VetorIniciado(Fundos) Then
            For i = 0 To UBound(Fundos)
              CalcularFundos TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, Fundos(i)
            Next i
          End If
          
          FracoesNao = ""
          Set rs = db.OpenRecordset("select distinct desc_fracao from despesas where Condominio = " & !Codigo & " And Mes = '" & cpMes.Text & "';", dbOpenDynaset)
          With rs
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                CalcularTaxa TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                CalcularTaxaFixa TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                If TodosCond!geral < 4 Then
                  CalcularDesconto TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                End If
                If InStr(!desc_fracao, DescFracao) <= 0 Then
                  FracoesNao = FracoesNao & !desc_fracao
                End If
                Erase Fundos
                .MoveNext
              Loop
            End If
          End With
          Set rs = Nothing
          DespesaFixaIndividual !Codigo, cpMes.Text, !Nome
          
          With tbDespFixa
            .AddNew
            !Condominio = TodosCond!Codigo
            !Nome = TodosCond!Nome
            !Historico = DescricaoVira(ViraFundo)
            !Data_cadastro = Date
            !valor = PegaValor(ViraFundo)
            !Fracao = -1
            !Proprietario = 0
            !Sindico = 0
            !por_associado = 0
            !desc_fracao = FracoesNao
            .Update
          End With
          
          Set rs = db.OpenRecordset("select distinct desc_fracao from despesas where desc_fracao <> '" & DescFracao & "' and Condominio = " & !Codigo & " And Mes = '" & cpMes.Text & "';", dbOpenDynaset)
          With rs
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                CalcularTaxaFixa TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                .MoveNext
              Loop
            End If
          End With
          
          PrintTela !Codigo, !geral
        Else
          db.Execute "Delete From DespesaDistribuida where mes = '" & cpMes.Text & "' and id_condominio = " & !Codigo & ";"
          db.Execute "Delete From DespesaInquilino where mes = '" & cpMes.Text & "' and id_condominio = " & !Codigo & ";"
          
          'buscar nas despesas fixas se há algum por percentual sobre despesas
          Sql = "Select * From Despesas Where Condominio = " & !Codigo & " And Mes = '" & cpMes.Text & "' and desconto = 0 Order By Tipo;"
          Set rs1 = db.OpenRecordset(Sql, dbOpenDynaset)
          totalDesp = 0
          With rs1
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                If !Percentual <> 1 Then
                  totalDesp = totalDesp + !valor
                End If
                .MoveNext
              Loop
            End If
          End With
          
          Set rs1 = Nothing
          Set rs1 = db.OpenRecordset("select * from despesafixa where condominio = " & cpCodigo.Text _
              & " and fracao = 2;", dbOpenDynaset)
          i = 0
          With rs1
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                ReDim Preserve Fundos(i)
                Fundos(i).Descricao = !Historico
                Fundos(i).Percentual = !valor
                Fundos(i).valor = !valor * totalDesp / 100
                Fundos(i).Sindico = !Sindico
                Fundos(i).Fracao = !desc_fracao
                i = i + 1
                .MoveNext
              Loop
            End If
          End With
            
          Set rs1 = Nothing

          If VetorIniciado(Fundos) Then
            For i = 0 To UBound(Fundos)
              CalcularFundos TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, Fundos(i)
            Next i
          End If
          
          Set rs = db.OpenRecordset("select distinct desc_fracao from despesas where Condominio = " & !Codigo & " And Mes = '" & cpMes.Text & "';", dbOpenDynaset)
          With rs
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                CalcularTaxa TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                CalcularTaxaFixa TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                If TodosCond!geral < 4 Then
                  CalcularDesconto TodosCond!Codigo, cpMes.Text, TodosCond!geral, 0, !desc_fracao
                End If
                Erase Fundos
                .MoveNext
              Loop
            End If
          End With
          Set rs = Nothing
          DespesaFixaIndividual !Codigo, cpMes.Text, !Nome
          PrintTela !Codigo, !geral
        End If
      End If
    End If
  End With
  cmdPrint.Enabled = True
End Sub

Private Sub cpCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMes.SetFocus
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
    cmdPrint.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Refresh
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Set tbDespFixa = db.OpenRecordset("despesafixa", dbOpenTable)
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
End Sub

Private Sub PrintTela(ByVal iCod As Long, iGeral As Long)
  Dim rsAuxiliar As Recordset
  Dim rsLeitura As Recordset
  Dim rsParticular As Recordset
  Dim relDespX As Recordset
  Dim vFora As Double
  Dim vProp As String
  Dim sTit As String
  Dim id_fora As Long
  
  cmdPrint.Enabled = False
  
  dbLocal.Execute "delete from leitura_auxiliar;"
  Set rsAuxiliar = dbLocal.OpenRecordset("leitura_auxiliar", dbOpenTable)
  ') AS SomaDeFRACAO
  Set relDespX = db.OpenRecordset("SELECT despesainquilino.MES, despesainquilino.DESCRICAO, despesainquilino.ID_CONDOMINIO, Sum(despesainquilino.VALOR_PROPRIETARIO) AS SomaDeVALOR_PROPRIETARIO, despesainquilino.ID_ASSOCIADO, despesainquilino.BOX, Sum(despesainquilino.VALOR) AS SomaDeVALOR, Sum(despesainquilino.FRACAO) AS SomaDeFRACAO " _
      & "from despesainquilino GROUP BY despesainquilino.MES, despesainquilino.DESCRICAO, despesainquilino.ID_CONDOMINIO, despesainquilino.ID_ASSOCIADO, despesainquilino.BOX " _
      & "HAVING (((despesainquilino.MES)='" & cpMes.Text & "') AND ((despesainquilino.ID_CONDOMINIO)=" & iCod & ")) ORDER BY despesainquilino.DESCRICAO;", dbOpenDynaset)
  
  id_fora = 0
  With relDespX
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        'leituras
        rsAuxiliar.AddNew
        If Trim(!Descricao & "") = "" Then
          Set rsLeitura = db.OpenRecordset("SELECT LEITURA_INDIVIDUAL.* FROM CONTAS_LEITURA LEFT " _
              & "JOIN LEITURA_INDIVIDUAL ON CONTAS_LEITURA.ID_LEITURA = LEITURA_INDIVIDUAL.ID_LEITURA " _
              & "WHERE (((CONTAS_LEITURA.CODIGO_CONDOMINIO)=" & !id_condominio & "" _
              & ") AND ((LEITURA_INDIVIDUAL.ID_ASSOCIADO)=" & !id_associado & "" _
              & ") AND ((CONTAS_LEITURA.MES_LEITURA)= '" & cpMes.Text & "'));", dbOpenDynaset)
  
          If rsLeitura.RecordCount > 0 Then
            rsLeitura.MoveFirst
            rsAuxiliar!ant_individual = rsLeitura!ant_individual
            rsAuxiliar!atual_individual = rsLeitura!atual_individual
            rsAuxiliar!gasto_individual = rsLeitura!gasto_individual
            rsAuxiliar!valor_individual = rsLeitura!valor_individual
            rsAuxiliar!id_leitura = rsLeitura!id_leitura
          Else
            rsAuxiliar!ant_individual = 0
            rsAuxiliar!atual_individual = 0
            rsAuxiliar!gasto_individual = 0
            rsAuxiliar!valor_individual = 0
            rsAuxiliar!id_leitura = -1
          End If
        End If
        rsAuxiliar!id_associado = !id_associado
        rsAuxiliar!nome_associado = NomeCompleto(!id_associado)
        
        'despesas particulares
        Set rsParticular = db.OpenRecordset("select * from porfora where mes = '" & cpMes.Text _
            & "' and associado = " & !id_associado & ";", dbOpenDynaset)
        vFora = 0
        If rsParticular.RecordCount > 0 Then
          rsParticular.MoveFirst
          If id_fora <> rsParticular!Associado Then
            id_fora = rsParticular!Associado
            Do While Not rsParticular.EOF
              vFora = vFora + rsParticular!valor
              rsParticular.MoveNext
            Loop
          End If
        End If
        Set rsParticular = Nothing
        rsAuxiliar!desp_individual = vFora
        rsAuxiliar!desp_coef = !somadeValor
        rsAuxiliar!total_pagar = rsAuxiliar!desp_coef + rsAuxiliar!desp_individual + rsAuxiliar!valor_individual
        rsAuxiliar!valor_pro = !somadeVALOR_PROPRIETARIO
        rsAuxiliar!Descricao = !Descricao
        tbAssociados.Index = "codigoid"
        rsAuxiliar!Fracao = !SomaDeFRACAO
        rsAuxiliar.Update
        Set rsLeitura = Nothing
        .MoveNext
      Loop
    End If
  End With
  
  Set rsAuxiliar = Nothing
  
  If iGeral < 4 Then
    vProp = "desppropri|" & TotalDespProp(iCod, cpMes.Text) & "|"
  End If
  sTit = cpNome.Text & vbCrLf & "Distribuição das despesas mês base: " & Format$(cpMes.Text, "MM/yyyy")

  RelatoriosRPT.Carregar "", Parametros.dados, "", sTit, sFormataCaminho(App.Path) & "distdespesas.rpt", , , , "LEITURA_AUXILIAR"
  cmdPrint.Enabled = True
End Sub

Private Function PegaNomeAssociado(ByVal iCodigo As Long) As String
  Dim rs As Recordset
  Dim ret As String
  Set rs = db.OpenRecordset("select tipo, apartamento from associados where codigo = " & iCodigo & ";", dbOpenDynaset)
  ret = ""
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      ret = !Tipo & " " & !Apartamento
    End If
  End With
  Set rs = Nothing
  PegaNomeAssociado = ret
End Function

Private Function ViraFundoReserva(ByVal nCod As Long) As Long
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("select * from associados where condominio = " & nCod & " and virafundo = true;", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      ViraFundoReserva = !Codigo
    Else
      ViraFundoReserva = -1
    End If
  End With
  Set rs = Nothing
End Function

Private Function PegaValor(ByVal nCod As Long) As Double
  Dim rs As Recordset
  Set rs = db.OpenRecordset("select valor from despesainquilino where mes = '" & cpMes.Text & "' and id_associado = " & nCod & ";", dbOpenDynaset)
  
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      PegaValor = !valor
    Else
      PegaValor = 0
    End If
  End With
  Set rs = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set tbDespFixa = Nothing
  Set tbCondominio = Nothing
  Set TodosCond = Nothing
  Set Assoc = Nothing
  Set tbAssociados = Nothing
End Sub

Private Function FracaoVira(ByVal nCod As Long) As String
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("select descricao from fracao where id_associado = " & nCod & ";", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      FracaoVira = !Descricao & "|"
    Else
      FracaoVira = ""
    End If
  End With
  Set rs = Nothing
End Function

Private Function DescricaoVira(ByVal nCod As Long) As String
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("select descvira from associados where codigo = " & nCod & ";", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      DescricaoVira = !descvira & ""
    Else
      DescricaoVira = ""
    End If
  End With
  Set rs = Nothing
End Function
