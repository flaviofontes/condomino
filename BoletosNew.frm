VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{CCF682E3-14F1-11D7-80E2-10E08C102140}#1.0#0"; "zprogbar.ocx"
Begin VB.Form Boletos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de boletos bancários"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ControlBox      =   0   'False
   Icon            =   "BoletosNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin rdActiveText.ActiveText cpVencimento 
      Height          =   315
      Left            =   3780
      TabIndex        =   26
      Top             =   420
      Width           =   1335
      _ExtentX        =   2355
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
      TabIndex        =   23
      Top             =   4980
      Width           =   8175
      _ExtentX        =   14420
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
   Begin VB.CheckBox ChDebitos 
      Caption         =   "Incluir débitos anteriores"
      Height          =   195
      Left            =   4740
      TabIndex        =   22
      Top             =   780
      Visible         =   0   'False
      Width           =   2055
   End
   Begin rdActiveText.ActiveText cpMens1 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   6675
      _ExtentX        =   11774
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
      MaxLength       =   60
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMes 
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   420
      Width           =   1140
      _ExtentX        =   2011
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
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   90
      ScaleHeight     =   0
      ScaleWidth      =   8145
      TabIndex        =   19
      Top             =   4650
      Width           =   8205
   End
   Begin VB.TextBox cpMensagem 
      Height          =   285
      Index           =   3
      Left            =   90
      MaxLength       =   250
      TabIndex        =   12
      Top             =   4230
      Width           =   4665
   End
   Begin VB.TextBox cpMensagem 
      Height          =   285
      Index           =   0
      Left            =   90
      MaxLength       =   250
      TabIndex        =   9
      Text            =   "SR CAIXA, NÃO RECEBER APÓS 15 DIAS DE VENCIDO"
      Top             =   3345
      Width           =   4665
   End
   Begin VB.TextBox cpMensagem 
      Height          =   285
      Index           =   1
      Left            =   90
      MaxLength       =   250
      TabIndex        =   10
      Text            =   "APÓS VENCIMENTO, SÓ RECEBER COM MULTA DE "
      Top             =   3630
      Width           =   4665
   End
   Begin VB.TextBox cpMensagem 
      Height          =   285
      Index           =   2
      Left            =   90
      MaxLength       =   250
      TabIndex        =   11
      Top             =   3945
      Width           =   4665
   End
   Begin VB.CommandButton cmdEtiquetas 
      Caption         =   "&Imprimir"
      Height          =   795
      Left            =   7080
      Picture         =   "BoletosNew.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   1260
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Gerar"
      Height          =   795
      Left            =   7065
      Picture         =   "BoletosNew.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   540
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Fechar"
      Height          =   795
      Left            =   7080
      Picture         =   "BoletosNew.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2580
      Width           =   1260
   End
   Begin rdActiveText.ActiveText cpMens2 
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   1335
      Width           =   6675
      _ExtentX        =   11774
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
      MaxLength       =   60
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens3 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   1680
      Width           =   6675
      _ExtentX        =   11774
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
      MaxLength       =   60
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens4 
      Height          =   315
      Left            =   90
      TabIndex        =   6
      Top             =   2025
      Width           =   6675
      _ExtentX        =   11774
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
      MaxLength       =   60
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens5 
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   2370
      Width           =   6675
      _ExtentX        =   11774
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
      MaxLength       =   60
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens6 
      Height          =   315
      Left            =   90
      TabIndex        =   8
      Top             =   2715
      Width           =   6675
      _ExtentX        =   11774
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
      MaxLength       =   60
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   24
      Top             =   60
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   2820
      TabIndex        =   25
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Despesas do mês"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   165
      TabIndex        =   20
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Texto de responsabilidade do caixa"
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   3120
      Width           =   2505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mensagen promocional"
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   780
      Width           =   1650
   End
   Begin VB.Label Porcento 
      AutoSize        =   -1  'True
      Caption         =   "Progresso"
      Height          =   195
      Left            =   75
      TabIndex        =   16
      Top             =   4755
      Width           =   705
   End
End
Attribute VB_Name = "Boletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nCod As Long
Dim sTemp As String
Dim qDig As Integer
Dim WithEvents gBoleto As BoletoSicoob
Attribute gBoleto.VB_VarHelpID = -1
Dim tbCondominio As Recordset
Dim tbBoletos As Recordset

Private Sub cmdCancelar_Click()
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
      End If
    End With
  End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Errado
  Dim SelDesp As Recordset
  Dim SelAsso As Recordset
  Dim SelAtraso As Recordset
  Dim SelDesconto As Recordset
  Dim SqlStr  As String
  Dim nMen    As Double
  Dim sBarra  As String
  Dim nFator  As String * 4
  Dim fValor  As String
  Dim vbDig   As String * 1
  Dim Campo1  As String
  Dim Campo2  As String
  Dim Campo3  As String
  Dim Campo4  As String
  Dim Campo5  As String
  Dim QtdeAs  As Long
  Dim nVal    As Double
  Dim dtVenc  As Date
  Dim dFracao As Double
  Dim dDesconto As Double
  Dim dFora   As Double
  Dim dLeitura As Double
  Dim dTotal  As Double
  Dim selCondominio As Recordset
  Dim selBoletos As Recordset
  Dim selFora As Recordset
  Dim selLeitura As Recordset
  Dim NossoNumero As String
  Dim sLivre As String
  Dim dCredito As Double
  Dim dDifer As Double
  Dim sAviso As String
  Dim iResultado As Integer
  Dim dValorProprietario As Double
  Dim nDiasLimite As Integer
  Dim sGravaValor As String
  
  If Not IsDate("25/" & cpMes.Text) Then
    MsgBox "Por favor informe o mês das despesas.", vbCritical & vbOKOnly, "Aviso"
    cpMes.SetFocus
    GoTo Fim
  End If
  
  If Not IsDate(cpVencimento.Text) Then
    MsgBox "O vencimento informado não é válido.", vbCritical & vbOKOnly, "Aviso"
    cpVencimento.SetFocus
    GoTo Fim
  End If
  
  If CDate(cpVencimento.Text) < Date Then
    MsgBox "O vencimento informado é menor que a data atual.", vbCritical & vbOKOnly, "Aviso"
    cpVencimento.SetFocus
    GoTo Fim
  End If
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Por favor informe o condomínio.", vbCritical & vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    GoTo Fim
  End If
  
  cmdCancelar.Enabled = False
  cmdPrint.Enabled = False
  cmdEtiquetas.Enabled = False
  
  nCod = cpCodigo.Text
  
  Set selCondominio = db.OpenRecordset("Select * From CONDOMINIO where codigo = " & nCod & " order by nome;", dbOpenDynaset)
  
  With selCondominio
    If .RecordCount > 0 Then
      .MoveFirst
      If !Banco = 3 Then
        Set gBoleto = New BoletoSicoob
        DoEvents
        Dim SM As String
        Dim sc As String
        sc = cpMensagem(0).Text & "|" & cpMensagem(1).Text & "|" & cpMensagem(2).Text & "|" & cpMensagem(3).Text & "|"
        SM = cpMens1.Text & vbCrLf & cpMens2.Text & vbCrLf & cpMens3.Text & vbCrLf & cpMens4.Text & vbCrLf & cpMens5.Text & cpMens6.Text
        gBoleto.GerarBoletos cpMes.Text, cpVencimento.Text, nCod, SM, sc
        GoTo Fim
      End If
    End If
  End With
  
  SqlStr = "Select * From Despesainquilino Where id_Condominio = " & nCod & " And Mes = '" & cpMes.Text & "';"
  Set SelDesp = db.OpenRecordset(SqlStr, dbOpenDynaset)
  
  If SelDesp.RecordCount = 0 Then
    MsgBox "As despesa do mês '" & cpMes.Text & "' do condomínio: " & cpNome.Text & " não foram distribudas.", vbInformation + vbOKOnly, Caption
    GoTo Fim
  End If
  
  Set SelDesp = Nothing
  
  DBEngine.BeginTrans
  With selCondominio
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        
        'Apagando detalhes do boleto
        db.Execute "delete from boletodetalhe where id_condominio = " & !Codigo & " and mes = '" & cpMes.Text & "';"
        
        Set selBoletos = db.OpenRecordset("Select * from boletos where cond = " & !Codigo & " and tran = '" & cpMes.Text & "';", dbOpenDynaset)
        
        If selBoletos.RecordCount > 0 Then
          dtVenc = cpVencimento.Text
          
          SqlStr = "Select * From Associados Where Condominio = " & !Codigo & " Order By tipo, apartamento"
          Set SelAsso = db.OpenRecordset(SqlStr, dbOpenDynaset)
          
          With SelAsso
            If .RecordCount > 0 Then
              .MoveLast
              .MoveFirst
              QtdeAs = .RecordCount
            Else
              QtdeAs = 0
            End If
          End With
          
          With SelAsso
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                selBoletos.FindFirst "cdsc = " & !Codigo
                If Not selBoletos.NoMatch Then
                  If selBoletos!pago = "N" And (selBoletos!idStatus = 1 Or selBoletos!idStatus = 8 Or selBoletos!idStatus = 5) Then
                    selBoletos.Delete
                  Else
                    GoTo japagou
                  End If
                End If
                
                dFracao = IIf(IsNull(!Fracao), 0, !Fracao)
                
                'valor do rateio
                nMen = 0
                dValorProprietario = 0
                SqlStr = "Select * From Despesainquilino Where id_associado = " & !Codigo & " And Mes = '" & cpMes.Text & "';"
                Set SelDesp = db.OpenRecordset(SqlStr, dbOpenDynaset)
                With SelDesp
                  If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                      nMen = nMen + Round(!valor, 2)
                      dValorProprietario = dValorProprietario + IIf(IsNull(!VALOR_PROPRIETARIO), 0, !VALOR_PROPRIETARIO)
                      .MoveNext
                    Loop
                  End If
                End With
                
                db.Execute "insert into BOLETODETALHE " & SqlStr
                db.Execute "UPDATE BOLETODETALHE SET VALOR = VALOR - VALOR_PROPRIETARIO WHERE ID_ASSOCIADO = " & !Codigo
                
                If dValorProprietario > 0 Then
                  sGravaValor = Replace(Format$(dValorProprietario, "#0.00"), ",", ".")
                  db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                    & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                    & "VALUES ('" & cpNome.Text & "', '" & cpMes.Text & "', 'DESP. PROPRIETÁRIO', 0, 0, '" _
                    & NomeCompleto(!Codigo) & "', 0, " & sGravaValor _
                    & ", '" & NomeCompleto(!Codigo) & "', " & !Codigo & ", " & selCondominio!Codigo _
                    & ", 0);"
                End If
                
                'Débitos anteriores
                nDiasLimite = PrazoBoleto(selCondominio!Codigo)
                dTotal = 0
                If !acumulado = "S" And (nDiasLimite > 0 And nDiasLimite < 21) Then
                  Set SelAtraso = db.OpenRecordset("Select * From BOLETOS Where left(TRAN,2) <> 'AV' and CDSC = " & !Codigo & " And (PAGO = 'N' Or PAGO Is Null) And VCTO < #" & Format$(dtVenc, "mm/dd/yyyy") & "# Order By vcto;", dbOpenDynaset)
                  With SelAtraso
                    If .RecordCount > 0 Then
                      .MoveLast
                      .MoveFirst
                      .MoveLast
                      dTotal = dTotal + !corrigido
                    End If
                  End With
                  Set SelAtraso = Nothing
                End If
                If dTotal > 0 Then
                  db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                    & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                    & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', 'DÉB. ANTERIOR(ES)', 0, 0, '" _
                    & NomeCompleto(!Codigo) & "', 0, " & Replace(Format$(dTotal, "#0.00"), ",", ".") _
                    & ", '" & NomeCompleto(!Codigo) & "', " & !Codigo & ", " & selCondominio!Codigo _
                    & ", 0);"
                End If
                
                'Por fora
                Set selFora = db.OpenRecordset("Select * From porfora Where associado = " & !Codigo & " And mes = '" & cpMes.Text & "';", dbOpenDynaset)
                dFora = 0
                With selFora
                  If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    Do While Not .EOF
                      dFora = dFora + !valor
                      db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                        & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                        & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', '" & !Historico & "', 0, 0, '" _
                        & NomeCompleto(SelAsso!Codigo) & "', 0, " & Replace(Format$(dFora, "#0.00"), ",", ".") _
                        & ", '" & NomeCompleto(SelAsso!Codigo) & "', " & SelAsso!Codigo & ", " & selCondominio!Codigo _
                        & ", 0);"
                      .MoveNext
                    Loop
                  End If
                End With
                Set selFora = Nothing
                
                'Descontos se houver
                Set SelDesconto = db.OpenRecordset("Select * From descontos Where id_inquilino = " & !Codigo & " And mes = '" & cpMes.Text & "';", dbOpenDynaset)
                dDesconto = 0
                With SelDesconto
                  If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    Do While Not .EOF
                      dDesconto = dDesconto + !valor
                      .MoveNext
                    Loop
                  End If
                End With
                Set SelDesconto = Nothing
                If dDesconto > 0 Then
                  db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                    & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                    & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', 'DESCONTO(S)', 0, 0, '" _
                    & NomeCompleto(!Codigo) & "', 0, " & Replace(Format$(dDesconto * -1, "#0.00"), ",", ".") _
                    & ", '" & NomeCompleto(!Codigo) & "', " & !Codigo & ", " & selCondominio!Codigo _
                    & ", 0);"
                End If
                
                'Leituras
                Set selLeitura = db.OpenRecordset("SELECT CONTAS_LEITURA.DESC_LEITURA, CONTAS_LEITURA.MES_LEITURA, LEITURA_INDIVIDUAL.VALOR_INDIVIDUAL, LEITURA_INDIVIDUAL.ID_ASSOCIADO " _
                  & "FROM CONTAS_LEITURA LEFT JOIN LEITURA_INDIVIDUAL ON CONTAS_LEITURA.ID_LEITURA = LEITURA_INDIVIDUAL.ID_LEITURA " _
                  & "WHERE (((CONTAS_LEITURA.MES_LEITURA)='" & cpMes.Text & "') AND ((LEITURA_INDIVIDUAL.ID_ASSOCIADO)=" & !Codigo & "));", dbOpenDynaset)
                dLeitura = 0
                With selLeitura
                  If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    Do While Not .EOF
                      dLeitura = dLeitura + !valor_individual
                      If !valor_individual > 0 Then
                        db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                          & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                          & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', '" & !desc_leitura & "', 0, 0, '" _
                          & NomeCompleto(SelAsso!Codigo) & "', 0, " & Replace(Format$(!valor_individual, "#0.00"), ",", ".") _
                          & ", '" & NomeCompleto(SelAsso!Codigo) & "', " & SelAsso!Codigo & ", " & selCondominio!Codigo _
                          & ", 0);"
                      End If
                      .MoveNext
                    Loop
                  End If
                End With
                Set selLeitura = Nothing
                
                nVal = nMen + dTotal + dFora + dLeitura
                
                If nVal > 0 Then
                  iResultado = 0
                  sAviso = ""
                  If dDesconto > 0 Then  'tem desconto
                    dCredito = nVal - dDesconto
                    If dCredito = 0 Then 'desconto é igual ao débito
                      GoTo japagou
                    ElseIf dCredito > 0 Then  'desconto é menor que o débito - pagar restante
                      dDifer = nVal
                      nVal = dCredito
                      sAviso = "Taxa condomínio: " & Format$(dDifer, "#0.00") & " - Desconto: " & _
                        Format$(dDesconto, "#0.00") & " = Valor a pagar: " & Format$(nVal, "#0.00")
                    ElseIf dCredito < 0 Then  'desconto é maior que o débito - gerar credito
                      sAviso = "Taxa condomínio: " & Format$(nVal, "#0.00") & " - Desconto: " & _
                        Format$(dDesconto, "#0.00") & " = Crédito: " & Format$(dCredito * -1, "#0.00")
                      iResultado = 1
                      'gerar credito proximo mês
                      AcrescentaCredito Format$(DateAdd("m", 1, CDate("01/" & cpMes.Text)), "MM/yyyy"), "CRÉDITO GERADO MÊS ANTERIOR", !Codigo, (dCredito * -1)
                    End If
                  End If
                  
                  fValor = Format(nVal, "#0.00")
                  fValor = Left(fValor, InStr(fValor, ",") - 1) & Right(fValor, 2)
                  fValor = PadLeft(fValor, "0", 10)
                  nFator = CStr(dtVenc - CDate("07/10/1997"))
                  
                  If selCondominio!tipoboleto = 1 Then
                    
                    NossoNumero = Trim(selCondominio!carteira) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
                    NossoNumero = NossoNumero & PadLeft(selCondominio!Codigo, "0", 4)
                    NossoNumero = NossoNumero & PadLeft(!Codigo, "0", 7)
                    
                    sTemp = PadLeft(selCondominio!Conta & "", "0", 6)

                    sLivre = sTemp & DigitoNosso(sTemp)
                    sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
                    sLivre = sLivre & Mid$(NossoNumero, 6, 3)
                    sLivre = sLivre & Mid$(NossoNumero, 2, 1)
                    sLivre = sLivre & Mid$(NossoNumero, 9)
                        
                    sLivre = sLivre & DigitoNosso(sLivre)
                  ElseIf selCondominio!tipoboleto = 2 Then
                  
                    sTemp = selCondominio!Conta & ""
                    
                    NossoNumero = Trim(selCondominio!carteira) & PadLeft(!Codigo, "0", 6 - Len(Trim(selCondominio!carteira))) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
                    
                    sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & PadLeft(sTemp, "0", 8)
                  
                  End If
                  
                  sBarra = Left(selCondominio!CdAgencia, 3) & "9" & nFator & fValor & sLivre
                            
                  vbDig = DigitoBarra(sBarra)
                  
                  sBarra = Left(sBarra, 4) & vbDig & Mid(sBarra, 5)
                  Campo1 = Mid(sBarra, 1, 4) & Mid(sBarra, 20, 5)
                  Campo1 = Campo1 & DigitoCodigo(Campo1)
                  Campo1 = Left(Campo1, 5) & "." & Mid(Campo1, 6)
                  
                  Campo2 = Mid(sBarra, 25, 10)
                  Campo2 = Campo2 & DigitoCodigo(Campo2)
                  Campo2 = Left(Campo2, 5) & "." & Mid(Campo2, 6)
                  
                  Campo3 = Mid(sBarra, 35)
                  Campo3 = Campo3 & DigitoCodigo(Campo3)
                  Campo3 = Left(Campo3, 5) & "." & Mid(Campo3, 6)
                  
                  Campo4 = Mid(sBarra, 5, 1)
                  
                  Campo5 = Mid(sBarra, 6, 4) & Mid(sBarra, 10, 10)
                  
                  selBoletos.AddNew
                  If selCondominio!titularboleto = 1 Then
                    selBoletos!Condominio = selCondominio!Nome
                    selBoletos!CGC = selCondominio!CGC
                  Else
                    If selCondominio!razaoboleto & "" <> "" Then
                      selBoletos!Condominio = selCondominio!razaoboleto
                      selBoletos!CGC = selCondominio!cnpjboleto
                    Else
                      selBoletos!Condominio = selCondominio!Nome
                      selBoletos!CGC = selCondominio!CGC
                    End If
                  End If
                  selBoletos!vcto = dtVenc
                  selBoletos!MENS = nMen
                  selBoletos!EXTR = 0
                  selBoletos!Data = Format(Date, "dd/mm/yyyy")
                  selBoletos!valr = nVal
                  selBoletos!corrigido = nVal
                  If !boleto = 1 Then
                    selBoletos!COTA = PadLeft(!Codigo, "0", 4)
                    selBoletos!cdsc = !Codigo
                    selBoletos!Nome = NomeCompleto(!Codigo)
                    selBoletos!cpf = !pcpf
                    selBoletos!Ende = !PEndereco
                    selBoletos!Bair = !PBairro
                    selBoletos!Cida = !PCidade
                    selBoletos!Esta = !PEstado
                    selBoletos!cep = !PCep
                  Else
                    selBoletos!COTA = PadLeft(!Codigo, "0", 4)
                    selBoletos!cdsc = !Codigo
                    selBoletos!Nome = NomeCompleto(!Codigo)
                    selBoletos!cpf = !cpf
                    selBoletos!Ende = selCondominio!endereco
                    selBoletos!Bair = selCondominio!bairro
                    selBoletos!Cida = selCondominio!Cidade
                    selBoletos!Esta = selCondominio!estado
                    selBoletos!cep = selCondominio!cep
                  End If
                  selBoletos!tran = cpMes.Text
                  selBoletos!DIGITAVAL = NossoNumero & "." & DigitoNosso(NossoNumero)
                  selBoletos!agcedente = selCondominio!agcedente
                  selBoletos!carteira = selCondominio!carteira
                  If selCondominio!tipoboleto = 1 Then
                    selBoletos!Mensagem = selCondominio!agcedente & "/" & selCondominio!Conta & "-" & DigitoNosso(selCondominio!Conta)
                  Else
                    selBoletos!Mensagem = Trim(selCondominio!agcedente) & "." & Trim(selCondominio!Operacao) & "." & Trim(selCondominio!Conta) & "." & DigitoNosso(Trim(selCondominio!agcedente) & Trim(selCondominio!Operacao) & Trim(selCondominio!Conta))
                  End If
                  If iResultado = 0 Then
                    selBoletos!pago = "N"
                    selBoletos!CDBARRA = sBarra
                    selBoletos!CODI = Campo1 & " " & Campo2 & " " & Campo3 & " " & Campo4 & " " & Campo5
                  Else
                    selBoletos!pago = "S"
                    selBoletos!CDBARRA = ""
                    selBoletos!CODI = "SIMPLES CONFERÊNCIA"
                  End If
                  selBoletos!texto = cpMens1.Text & vbCrLf & cpMens2.Text & vbCrLf & cpMens3.Text & vbCrLf & cpMens4.Text & vbCrLf & cpMens5.Text & vbCrLf & cpMens6.Text & vbCrLf & vbCrLf & vbCrLf & sAviso
                  selBoletos!INST1 = cpMensagem(0).Text
                  selBoletos!INST2 = cpMensagem(1).Text
                  selBoletos!INST3 = cpMensagem(2).Text
                  selBoletos!INST4 = cpMensagem(3).Text
                  selBoletos!Banco = "CAIXA" 'selCondominio!Banco
                  selBoletos!CdBanco = selCondominio!CdAgencia
                  selBoletos!CANCELADO = "N"
                  selBoletos!cond = selCondominio!Codigo
                  selBoletos!acumulado = !acumulado
                  selBoletos!bole = NossoNumero
                  selBoletos!nosso = NossoNumero
                  selBoletos!desconto = dDesconto
                  selBoletos!idStatus = 1
                  selBoletos.Update
                  selBoletos.Bookmark = selBoletos.LastModified
                  db.Execute "update boletodetalhe set id_boleto = " & selBoletos!id & " where id_boleto is null;"
                  DoEvents
                End If
japagou:
                If Not .EOF Then
                  Porcento.Caption = "Montando... " & NomeCompleto(!Codigo)
                  Porcento.Refresh
                  Barra.Value = .PercentPosition
                End If
                .MoveNext
              Loop
            End If
          End With
        Else
          dtVenc = cpVencimento.Text
          
          SqlStr = "Select * From Associados Where Condominio = " & !Codigo & " Order By tipo, apartamento"
          Set SelAsso = db.OpenRecordset(SqlStr, dbOpenDynaset)
          
          With SelAsso
            If .RecordCount > 0 Then
              .MoveLast
              .MoveFirst
              QtdeAs = .RecordCount
            Else
              QtdeAs = 0
            End If
          End With
          
          With SelAsso
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not .EOF
                dFracao = IIf(IsNull(!Fracao), 0, !Fracao)
                
                'valor rateio
                nMen = 0
                dValorProprietario = 0
                SqlStr = "Select * From Despesainquilino Where id_associado = " & !Codigo & " And Mes = '" & cpMes.Text & "';"
                Set SelDesp = db.OpenRecordset(SqlStr, dbOpenDynaset)
                With SelDesp
                  If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                      nMen = nMen + Round(!valor, 2)
                      dValorProprietario = dValorProprietario + IIf(IsNull(!VALOR_PROPRIETARIO), 0, !VALOR_PROPRIETARIO)
                      .MoveNext
                    Loop
                  End If
                End With
                
                db.Execute "insert into BOLETODETALHE " & SqlStr
                db.Execute "UPDATE BOLETODETALHE SET VALOR = VALOR - VALOR_PROPRIETARIO WHERE ID_ASSOCIADO = " & !Codigo
                If dValorProprietario > 0 Then
                  db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                    & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                    & "VALUES ('" & cpNome.Text & "', '" & cpMes.Text & "', 'DESP. PROPRIETÁRIO', 0, 0, '" _
                    & NomeCompleto(!Codigo) & "', 0, " & Replace(Format$(dValorProprietario, "#0.00"), ",", ".") _
                    & ", '" & NomeCompleto(!Codigo) & "', " & !Codigo & ", " & selCondominio!Codigo _
                    & ", 0);"
                End If
                
                'Débitos anteriores--aqui
                nDiasLimite = PrazoBoleto(selCondominio!Codigo)
                dTotal = 0
                If !acumulado = "S" And (nDiasLimite > 0 And nDiasLimite < 21) Then
                  Set SelAtraso = db.OpenRecordset("Select * From BOLETOS Where left(TRAN,2) <> 'AV' and CDSC = " & !Codigo & " And (PAGO = 'N' Or PAGO Is Null) And VCTO < #" & Format$(Date, "mm/dd/yyyy") & "# Order By vcto;", dbOpenDynaset)
                  With SelAtraso
                    If .RecordCount > 0 Then
                      .MoveLast
                      .MoveFirst
                      .MoveLast
                      dTotal = dTotal + !corrigido
                    End If
                  End With
                  Set SelAtraso = Nothing
                End If
                If dTotal > 0 Then
                  db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                    & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                    & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', 'DÉB. ANTERIOR(ES)', 0, 0, '" _
                    & NomeCompleto(!Codigo) & "', 0, " & Replace(Format$(dTotal, "#0.00"), ",", ".") _
                    & ", '" & NomeCompleto(!Codigo) & "', " & !Codigo & ", " & selCondominio!Codigo _
                    & ", 0);"
                End If
                
                'Por fora
                Set selFora = db.OpenRecordset("Select * From porfora Where associado = " & !Codigo & " And mes = '" & cpMes.Text & "';", dbOpenDynaset)
                dFora = 0
                With selFora
                  If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    Do While Not .EOF
                      dFora = dFora + !valor
                      db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                        & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                        & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', '" & !Historico & "', 0, 0, '" _
                        & NomeCompleto(SelAsso!Codigo) & "', 0, " & Replace(Format$(dFora, "#0.00"), ",", ".") _
                        & ", '" & NomeCompleto(SelAsso!Codigo) & "', " & SelAsso!Codigo & ", " & selCondominio!Codigo _
                        & ", 0);"
                      .MoveNext
                    Loop
                  End If
                End With
                Set selFora = Nothing
                
                'Descontos se houver
                Set SelDesconto = db.OpenRecordset("Select * From descontos Where id_inquilino = " & !Codigo & " And mes = '" & cpMes.Text & "';", dbOpenDynaset)
                dDesconto = 0
                With SelDesconto
                  If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    Do While Not .EOF
                      dDesconto = dDesconto + !valor
                      .MoveNext
                    Loop
                  End If
                End With
                Set SelDesconto = Nothing
                If dDesconto > 0 Then
                  db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                    & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                    & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', 'DESCONTO(S)', 0, 0, '" _
                    & NomeCompleto(!Codigo) & "', 0, " & Replace(Format$(dDesconto * -1, "#0.00"), ",", ".") _
                    & ", '" & NomeCompleto(!Codigo) & "', " & !Codigo & ", " & selCondominio!Codigo _
                    & ", 0);"
                End If
                
                'Leituras
                Set selLeitura = db.OpenRecordset("SELECT CONTAS_LEITURA.DESC_LEITURA, CONTAS_LEITURA.MES_LEITURA, LEITURA_INDIVIDUAL.VALOR_INDIVIDUAL, LEITURA_INDIVIDUAL.ID_ASSOCIADO " _
                  & "FROM CONTAS_LEITURA LEFT JOIN LEITURA_INDIVIDUAL ON CONTAS_LEITURA.ID_LEITURA = LEITURA_INDIVIDUAL.ID_LEITURA " _
                  & "WHERE (((CONTAS_LEITURA.MES_LEITURA)='" & cpMes.Text & "') AND ((LEITURA_INDIVIDUAL.ID_ASSOCIADO)=" & !Codigo & "));", dbOpenDynaset)
                dLeitura = 0
                With selLeitura
                  If .RecordCount > 0 Then
                    .MoveLast
                    .MoveFirst
                    Do While Not .EOF
                      dLeitura = dLeitura + !valor_individual
                      If !valor_individual > 0 Then
                        db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
                          & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO) " _
                          & "VALUES ('" & cpCodigo.Text & "', '" & cpMes.Text & "', '" & !desc_leitura & "', 0, 0, '" _
                          & NomeCompleto(SelAsso!Codigo) & "', 0, " & Replace(Format$(!valor_individual, "#0.00"), ",", ".") _
                          & ", '" & NomeCompleto(SelAsso!Codigo) & "', " & SelAsso!Codigo & ", " & selCondominio!Codigo _
                          & ", 0);"
                      End If
                      .MoveNext
                    Loop
                  End If
                End With
                Set selLeitura = Nothing
                
                nVal = nMen + dTotal + dFora + dLeitura
                
                If nVal > 0 Then
                  
                  iResultado = 0
                  sAviso = ""
                  If dDesconto > 0 Then  'tem desconto
                    dCredito = nVal - dDesconto
                    If dCredito = 0 Then 'desconto é igual ao débito
                      GoTo japagou
                    ElseIf dCredito > 0 Then  'desconto é menor que o débito - pagar restante
                      dDifer = nVal
                      nVal = dCredito
                      sAviso = "Taxa condomínio: " & Format$(dDifer, "#0.00") & " - Desconto: " & _
                        Format$(dDesconto, "#0.00") & " = Valor a pagar: " & Format$(nVal, "#0.00")
                    ElseIf dCredito < 0 Then  'desconto é maior que o débito - gerar credito
                      sAviso = "Taxa condomínio: " & Format$(nVal, "#0.00") & " - Desconto: " & _
                        Format$(dDesconto, "#0.00") & " = Crédito: " & Format$(dCredito * -1, "#0.00")
                      iResultado = 1
                      'gerar credito proximo mês
                      AcrescentaCredito Format$(DateAdd("m", 1, CDate("01/" & cpMes.Text)), "MM/yyyy"), "CRÉDITO GERADO MÊS ANTERIOR", !Codigo, (dCredito * -1)
                    End If
                  End If
                  
                  fValor = Format(nVal, "#0.00")
                  fValor = Left(fValor, InStr(fValor, ",") - 1) & Right(fValor, 2)
                  fValor = PadLeft(fValor, "0", 10)
                  nFator = CStr(dtVenc - CDate("07/10/1997"))
                  
                  If selCondominio!tipoboleto = 1 Then
                    
                    NossoNumero = Trim(selCondominio!carteira) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
                    NossoNumero = NossoNumero & PadLeft(selCondominio!Codigo, "0", 4)
                    NossoNumero = NossoNumero & PadLeft(!Codigo, "0", 7)
                    
                    sTemp = PadLeft(selCondominio!Conta & "", "0", 6)

                    sLivre = sTemp & DigitoNosso(sTemp)
                    sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
                    sLivre = sLivre & Mid$(NossoNumero, 6, 3)
                    sLivre = sLivre & Mid$(NossoNumero, 2, 1)
                    sLivre = sLivre & Mid$(NossoNumero, 9)
                        
                    sLivre = sLivre & DigitoNosso(sLivre)
                  ElseIf selCondominio!tipoboleto = 2 Then
                  
                    sTemp = selCondominio!Conta & ""
                    
                    NossoNumero = Trim(selCondominio!carteira) & PadLeft(!Codigo, "0", 6 - Len(Trim(selCondominio!carteira))) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
                    
                    sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & PadLeft(sTemp, "0", 8)
                  
                  End If
                      
                  sBarra = Left(selCondominio!CdAgencia, 3) & "9" & nFator & fValor & sLivre
                            
                  vbDig = DigitoBarra(sBarra)
                  
                  sBarra = Left(sBarra, 4) & vbDig & Mid(sBarra, 5)
                  Campo1 = Mid(sBarra, 1, 4) & Mid(sBarra, 20, 5)
                  Campo1 = Campo1 & DigitoCodigo(Campo1)
                  Campo1 = Left(Campo1, 5) & "." & Mid(Campo1, 6)
                  
                  Campo2 = Mid(sBarra, 25, 10)
                  Campo2 = Campo2 & DigitoCodigo(Campo2)
                  Campo2 = Left(Campo2, 5) & "." & Mid(Campo2, 6)
                  
                  Campo3 = Mid(sBarra, 35)
                  Campo3 = Campo3 & DigitoCodigo(Campo3)
                  Campo3 = Left(Campo3, 5) & "." & Mid(Campo3, 6)
                  
                  Campo4 = Mid(sBarra, 5, 1)
                  
                  Campo5 = Mid(sBarra, 6, 4) & Mid(sBarra, 10, 10)
                  
                  tbBoletos.AddNew
                  If selCondominio!titularboleto = 1 Then
                    tbBoletos!Condominio = selCondominio!Nome
                    tbBoletos!CGC = selCondominio!CGC
                  Else
                    If selCondominio!razaoboleto & "" <> "" Then
                      tbBoletos!Condominio = selCondominio!razaoboleto
                      tbBoletos!CGC = selCondominio!cnpjboleto
                    Else
                      tbBoletos!Condominio = selCondominio!Nome
                      tbBoletos!CGC = selCondominio!CGC
                    End If
                  End If
                  tbBoletos!vcto = dtVenc
                  tbBoletos!MENS = nMen
                  tbBoletos!EXTR = 0
                  tbBoletos!Data = Format(Date, "dd/mm/yyyy")
                  tbBoletos!valr = nVal
                  tbBoletos!corrigido = nVal
                  If !boleto = 1 Then
                    tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
                    tbBoletos!cdsc = !Codigo
                    tbBoletos!Nome = NomeCompleto(!Codigo)
                    tbBoletos!cpf = !pcpf
                    tbBoletos!Ende = !PEndereco
                    tbBoletos!Bair = !PBairro
                    tbBoletos!Cida = !PCidade
                    tbBoletos!Esta = !PEstado
                    tbBoletos!cep = !PCep
                  Else
                    tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
                    tbBoletos!cdsc = !Codigo
                    tbBoletos!Nome = NomeCompleto(!Codigo)
                    tbBoletos!cpf = !cpf
                    tbBoletos!Ende = selCondominio!endereco
                    tbBoletos!Bair = selCondominio!bairro
                    tbBoletos!Cida = selCondominio!Cidade
                    tbBoletos!Esta = selCondominio!estado
                    tbBoletos!cep = selCondominio!cep
                  End If
                  tbBoletos!tran = cpMes.Text
                  tbBoletos!DIGITAVAL = NossoNumero & "." & DigitoNosso(NossoNumero)
                  tbBoletos!agcedente = selCondominio!agcedente
                  tbBoletos!carteira = selCondominio!carteira
                  If selCondominio!tipoboleto = 1 Then
                    tbBoletos!Mensagem = selCondominio!agcedente & "/" & selCondominio!Conta & "-" & DigitoNosso(selCondominio!Conta)
                  Else
                    tbBoletos!Mensagem = Trim(selCondominio!agcedente) & "." & Trim(selCondominio!Operacao) & "." & Trim(selCondominio!Conta) & "." & DigitoNosso(Trim(selCondominio!agcedente) & Trim(selCondominio!Operacao) & Trim(selCondominio!Conta))
                  End If
                  If iResultado = 0 Then
                    tbBoletos!pago = "N"
                    tbBoletos!CODI = Campo1 & " " & Campo2 & " " & Campo3 & " " & Campo4 & " " & Campo5
                    tbBoletos!CDBARRA = sBarra
                  Else
                    tbBoletos!pago = "S"
                    tbBoletos!CODI = "SIMPLES CONFERÊNCIA"
                    tbBoletos!CDBARRA = ""
                  End If
                  tbBoletos!texto = cpMens1.Text & vbCrLf & cpMens2.Text & vbCrLf & cpMens3.Text & vbCrLf & cpMens4.Text & vbCrLf & cpMens5.Text & vbCrLf & cpMens6.Text & vbCrLf & vbCrLf & vbCrLf & sAviso
                  tbBoletos!INST1 = cpMensagem(0).Text
                  tbBoletos!INST2 = cpMensagem(1).Text
                  tbBoletos!INST3 = cpMensagem(2).Text
                  tbBoletos!INST4 = cpMensagem(3).Text
                  tbBoletos!Banco = "CAIXA" 'selCondominio!Banco
                  tbBoletos!CdBanco = selCondominio!CdAgencia
                  tbBoletos!CANCELADO = "N"
                  tbBoletos!cond = selCondominio!Codigo
                  tbBoletos!acumulado = !acumulado
                  tbBoletos!bole = NossoNumero
                  tbBoletos!nosso = NossoNumero
                  tbBoletos!desconto = dDesconto
                  tbBoletos!idStatus = 1
                  tbBoletos.Update
                  tbBoletos.Bookmark = tbBoletos.LastModified
                  db.Execute "update boletodetalhe set id_boleto = " & tbBoletos!id & " where id_boleto is null;"
                  DoEvents
                End If
                If Not .EOF Then
                  Porcento.Caption = "Montando... " & NomeCompleto(!Codigo)
                  Porcento.Refresh
                  If Barra.Value <> Int(.PercentPosition) Then
                    Barra.Value = Int(.PercentPosition)
                  End If
                End If
                .MoveNext
              Loop
            End If
          End With
        End If
Proximo:
        .MoveNext
      Loop
    End If
  End With
  DBEngine.CommitTrans
  
Fim:
  Porcento.Caption = "Concluido..."
  Porcento.Refresh
  Barra.Value = 100
  cmdCancelar.Enabled = True
  cmdPrint.Enabled = True
  cmdEtiquetas.Enabled = True
  Exit Sub

Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  DBEngine.Rollback
  Resume Fim

End Sub

Private Sub cmdEtiquetas_Click()
  Dim sSql As String
  Dim nCod As Long
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Selecione um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  If Not IsDate("25/" & cpMes.Text) Then
    MsgBox "O mês/ano informado não é válido.", vbInformation + vbOKOnly, "Aviso"
    cpMes.SetFocus
    Exit Sub
  End If
  
'  cmdCancelar.Enabled = False
'  cmdPrint.Enabled = False
'  cmdEtiquetas.Enabled = False
  
  nCod = cpCodigo.Text
  
  sSql = "Select * from boletos where pago='N' and tran = '" & cpMes.Text & "' and cond = " & nCod & " order by bole;"
  
  Dim PerSind As Double
  
  PerSind = PersentualSindico(CodigoSindico(nCod))
  
  RelatoriosRPT.mnuSupExport.Visible = False
  RelatoriosRPT.mnuExportBoleto.Visible = True
  RelatoriosRPT.mnuOrdenar.Visible = True
  RelatoriosRPT.mnuEviarEmail.Visible = True
  RelatoriosRPT.Carregar "{boletos.bole};crAscendingOrder;boletos|", Parametros.dados, "{boletos.pago} = 'N' and {boletos.tran} = '" & cpMes.Text & "' and {boletos.cond} = " _
      & nCod, "Boletos", sFormataCaminho(App.Path) & "cobranca.rpt", , sSql, "persind|" & PerSind, , nCod, cpMes.Text, , "subdesccobranca;subcobranca.rpt;cob_despesas_sub.rpt"
  
'  cmdCancelar.Enabled = True
'  cmdPrint.Enabled = True
'  cmdEtiquetas.Enabled = True
End Sub

Private Sub cpCodigo_LostFocus()
  Dim nAno    As String
  Dim nMes    As String
  Dim nDia    As String
  If Val(cpCodigo.Text) = 0 Then
    cpMensagem(0).Text = "SR CAIXA, NÃO RECEBER APÓS 15 DIAS DE VENCIDO"
    cpMensagem(1).Text = "APÓS VENCIMENTO SÓ RECEBER"
    cpMensagem(2).Text = "COM JUROS DE " & Parametros.juros & "% AO MÊS + MULTA DE " & Parametros.Multa & "%"
    cpMensagem(3).Text = ""
  Else
    If cpMes.Text <> "" Then
      cpMes.Text = FormataMesAno(cpMes.Text)
      nDia = StrZero(PegaVencimento(cpCodigo.Text), 2)
      nMes = Left(cpMes.Text, 2)
      nAno = Right(cpMes.Text, 4)
      If Val(nMes) = 12 Then
        nMes = "01"
        nAno = StrZero(Val(nAno) + 1, 4)
      Else
        nMes = StrZero(Val(nMes) + 1, 2)
      End If
      cpVencimento.Text = nDia + "/" + nMes + "/" + nAno
    End If
    If PrazoBoleto(cpCodigo.Text) = 0 Then
      cpMensagem(0).Text = "SR CAIXA"
    Else
      cpMensagem(0).Text = "SR CAIXA, NÃO RECEBER APÓS " & PegaDias(cpCodigo.Text) & " DIAS DE VENCIDO"
    End If
    cpMensagem(1).Text = "APÓS VENCIMENTO SÓ RECEBER"
    cpMensagem(2).Text = "COM JUROS DE " & PegaJuros(cpCodigo.Text) & _
          "% AO MÊS + MULTA DE " & PegaMulta(cpCodigo.Text) & "%"
    cpMensagem(3).Text = ""
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

Private Sub cpMens1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens2.SetFocus
  End If
End Sub

Private Sub cpMens2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens3.SetFocus
  End If
End Sub

Private Sub cpMens3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens4.SetFocus
  End If
End Sub

Private Sub cpMens4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens5.SetFocus
  End If
End Sub

Private Sub cpMens5_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens6.SetFocus
  End If
End Sub

Private Sub cpMens6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMensagem(0).SetFocus
  End If
End Sub

Private Sub cpMensagem_GotFocus(Index As Integer)
  cpMensagem(Index).SelStart = 0
  cpMensagem(Index).SelLength = Len(cpMensagem(Index).Text)
End Sub

Private Sub cpMensagem_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      KeyAscii = 0
      If Index < 3 Then
        cpMensagem(Index + 1).SetFocus
      Else
        cmdPrint.SetFocus
      End If
    Case Else
      KeyAscii = vTexto(KeyAscii)
  End Select
End Sub

Private Sub cpMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens1.SetFocus
  End If
End Sub

Private Sub cpMes_LostFocus()
  Dim nAno    As String
  Dim nMes    As String
  Dim nDia    As String
  If cpMes.Text <> "" Then
    cpMes.Text = FormataMesAno(cpMes.Text)
    nDia = StrZero(PegaVencimento(cpCodigo.Text), 2)
    nMes = Left(cpMes.Text, 2)
    nAno = Right(cpMes.Text, 4)
    If Val(nMes) = 12 Then
      nMes = "01"
      nAno = StrZero(Val(nAno) + 1, 4)
    Else
      nMes = StrZero(Val(nMes) + 1, 2)
    End If
    cpVencimento.Text = nDia + "/" + nMes + "/" + nAno
  End If
End Sub

Private Sub cpVencimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens1.SetFocus
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Set tbBoletos = db.OpenRecordset("boletos", dbOpenTable)
  Refresh
  DoEvents
  KeyPreview = True
  cpMensagem(0).Text = "SR CAIXA, NÃO RECEBER APÓS 15 DIAS DE VENCIDO"
  cpMensagem(1).Text = "APÓS VENCIMENTO SÓ RECEBER"
  cpMensagem(2).Text = "COM JUROS DE " & Parametros.juros & "% AO MÊS + MULTA DE " & Parametros.Multa & "%"
  cpMensagem(3).Text = ""
End Sub

Private Function PegaDias(ByVal nCod As Long) As Integer
  Dim rs As Recordset
  Dim sRet As Integer
  Set rs = db.OpenRecordset("Select * From CONDOMINIO where codigo = " & nCod & " order by nome;", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      sRet = !dias
    Else
      sRet = 15
    End If
  End With
  Set rs = Nothing
  PegaDias = sRet
End Function

Private Function PegaVencimento(ByVal nCond As Long) As Integer
  Dim rs As Recordset
  Dim sRet As Integer
  Set rs = db.OpenRecordset("Select vencimento From CONDOMINIO where codigo = " & nCod & ";", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      sRet = !Vencimento
    Else
      sRet = 10
    End If
  End With
  Set rs = Nothing
  PegaVencimento = sRet
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
  Set tbBoletos = Nothing
End Sub

Private Sub gBoleto_Erro(sMensagem As String, sTitulo As String)
  MsgBox sMensagem, vbCritical + vbOKOnly, sTitulo
End Sub

Private Sub gBoleto_Progresso(p As Integer, Inquilino As String)
  Barra.Value = p
  Porcento.Caption = "Processando " & Inquilino
  Porcento.Refresh
End Sub

'                  If selCondominio!tipoboleto = 1 Then
'
'                    NossoNumero = Trim(selCondominio!carteira) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
'                    NossoNumero = NossoNumero & PadLeft(selCondominio!Codigo, "0", 4)
'                    NossoNumero = NossoNumero & PadLeft(!Codigo, "0", 7)
'
'                    sTemp = selCondominio!Conta & ""
'                    qDig = DigitosCedente(1)
'                    If Len(sTemp) > qDig Then
'                      sTemp = Right(sTemp, qDig)
'                    ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
'                      sTemp = PadLeft(sTemp, "0", qDig)
'                    End If
'
'                    sLivre = sTemp & DigitoNosso(sTemp)
'                    sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
'                    sLivre = sLivre & Mid$(NossoNumero, 6, 3)
'                    sLivre = sLivre & Mid$(NossoNumero, 2, 1)
'                    sLivre = sLivre & Mid$(NossoNumero, 9)
'
'                    sLivre = sLivre & DigitoNosso(sLivre)
'                  ElseIf selCondominio!tipoboleto = 2 Then
'
'                    sTemp = selCondominio!Conta & ""
'                    qDig = DigitosCedente(2)
'                    If Len(sTemp) > qDig Then
'                      sTemp = Right(sTemp, qDig)
'                    ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
'                      sTemp = PadLeft(sTemp, "0", qDig)
'                    End If
'
'                    NossoNumero = Trim(selCondominio!carteira) & PadLeft(!Codigo, "0", 6 - Len(Trim(selCondominio!carteira))) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
'
'                    sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & sTemp
'
'                  End If
'                  If selCondominio!tipoboleto = 1 Then
'
'                    NossoNumero = Trim(selCondominio!carteira) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
'                    NossoNumero = NossoNumero & PadLeft(selCondominio!Codigo, "0", 4)
'                    NossoNumero = NossoNumero & PadLeft(!Codigo, "0", 7)
'
'                    sTemp = selCondominio!Conta & ""
'                    qDig = DigitosCedente(1)
'                    If Len(sTemp) > qDig Then
'                      sTemp = Right(sTemp, qDig)
'                    ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
'                      sTemp = PadLeft(sTemp, "0", qDig)
'                    End If
'
'                    sLivre = sTemp & DigitoNosso(sTemp)
'                    sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
'                    sLivre = sLivre & Mid$(NossoNumero, 6, 3)
'                    sLivre = sLivre & Mid$(NossoNumero, 2, 1)
'                    sLivre = sLivre & Mid$(NossoNumero, 9)
'
'                    sLivre = sLivre & DigitoNosso(sLivre)
'                  ElseIf selCondominio!tipoboleto = 2 Then
'
'                    sTemp = selCondominio!Conta & ""
'                    qDig = DigitosCedente(2)
'                    If Len(sTemp) > qDig Then
'                      sTemp = Right(sTemp, qDig)
'                    ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
'                      sTemp = PadLeft(sTemp, "0", qDig)
'                    End If
'
'                    NossoNumero = Trim(selCondominio!carteira) & PadLeft(!Codigo, "0", 6 - Len(Trim(selCondominio!carteira))) & Left$(cpMes.Text, 2) & Right$(cpMes.Text, 2)
'
'                    sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & sTemp
'
'                  End If

