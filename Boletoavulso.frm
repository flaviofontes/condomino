VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form BoletosAvulso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de Boletos Bancários"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAnteriores 
      Caption         =   "Gerar de débitos anteriores"
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   3660
      Width           =   3075
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1740
      TabIndex        =   11
      Top             =   4200
      Width           =   555
   End
   Begin rdActiveText.ActiveText cpHistorico 
      Height          =   315
      Left            =   900
      TabIndex        =   13
      Top             =   4620
      Width           =   7095
      _ExtentX        =   12515
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
      MaxLength       =   200
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpValor 
      Height          =   315
      Left            =   2880
      TabIndex        =   15
      Top             =   5040
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
   Begin rdActiveText.ActiveText cpVenc 
      Height          =   315
      Left            =   1140
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.CommandButton cmdReimpime 
      Caption         =   "&Reimprimir"
      Height          =   795
      Left            =   6960
      Picture         =   "Boletoavulso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      Width           =   1260
   End
   Begin rdActiveText.ActiveText cpCodigo 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   4200
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
   Begin VB.TextBox cpNome 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2340
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Top             =   4200
      Width           =   5685
   End
   Begin VB.TextBox cpMensagem 
      Height          =   315
      Index           =   3
      Left            =   195
      MaxLength       =   80
      TabIndex        =   9
      Top             =   3675
      Width           =   4485
   End
   Begin VB.TextBox cpMensagem 
      Height          =   315
      Index           =   0
      Left            =   195
      MaxLength       =   80
      TabIndex        =   6
      Top             =   2670
      Width           =   4485
   End
   Begin VB.TextBox cpMensagem 
      Height          =   315
      Index           =   1
      Left            =   195
      MaxLength       =   80
      TabIndex        =   7
      Top             =   3015
      Width           =   4485
   End
   Begin VB.TextBox cpMensagem 
      Height          =   315
      Index           =   2
      Left            =   195
      MaxLength       =   80
      TabIndex        =   8
      Top             =   3330
      Width           =   4485
   End
   Begin VB.CommandButton cmdEtiquetas 
      Caption         =   "&Imprimir"
      Height          =   795
      Left            =   6960
      Picture         =   "Boletoavulso.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   135
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Fechar"
      Height          =   795
      Left            =   6960
      Picture         =   "Boletoavulso.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1020
      Width           =   1260
   End
   Begin rdActiveText.ActiveText cpMens1 
      Height          =   315
      Left            =   195
      TabIndex        =   0
      Top             =   330
      Width           =   5565
      _ExtentX        =   9816
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
      MaxLength       =   80
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens2 
      Height          =   315
      Left            =   195
      TabIndex        =   1
      Top             =   675
      Width           =   5565
      _ExtentX        =   9816
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
      MaxLength       =   80
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens3 
      Height          =   315
      Left            =   195
      TabIndex        =   2
      Top             =   1020
      Width           =   5565
      _ExtentX        =   9816
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
      MaxLength       =   80
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens4 
      Height          =   315
      Left            =   195
      TabIndex        =   3
      Top             =   1365
      Width           =   5565
      _ExtentX        =   9816
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
      MaxLength       =   80
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens5 
      Height          =   315
      Left            =   195
      TabIndex        =   4
      Top             =   1710
      Width           =   5565
      _ExtentX        =   9816
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
      MaxLength       =   80
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText cpMens6 
      Height          =   315
      Left            =   195
      TabIndex        =   5
      Top             =   2055
      Width           =   5565
      _ExtentX        =   9816
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
      MaxLength       =   80
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   2460
      TabIndex        =   24
      Top             =   5100
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   180
      TabIndex        =   22
      Top             =   5100
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Associado"
      Height          =   195
      Left            =   150
      TabIndex        =   21
      Top             =   4260
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Texto de responsabilidade do caixa"
      Height          =   195
      Left            =   195
      TabIndex        =   20
      Top             =   2445
      Width           =   2505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mensagens promocionais"
      Height          =   195
      Left            =   195
      TabIndex        =   19
      Top             =   90
      Width           =   1800
   End
End
Attribute VB_Name = "BoletosAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tbBoletos As Recordset
Dim tbAssociados As Recordset
Public DebAnt As Boolean
Dim cCondominio As Long

Private Sub cmdAnteriores_Click()
  LocMensalidade.Show 1
  If DebAnt = True Then
    cpCodigo_KeyPress 13
  End If
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub Montar()
On Error GoTo Errado
  Dim SelAsso As Recordset
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
  Dim nVal    As Double
  Dim dtVenc  As Date
  Dim selCondominio As Recordset
  Dim selBoletos As Recordset
  Dim NossoNumero As String
  Dim sLivre As String
  Dim nBoleto As Long
  Dim sSql As String
  Dim sBol As String
  Dim i As Integer
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Informe o inquilino.", vbCritical + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    GoTo Fim
  End If
  
  If Not IsDate(cpVenc.Text) Then
    MsgBox "Data de vencimento inválida.", vbCritical + vbOKOnly, "Aviso"
    cpVenc.SetFocus
    GoTo Fim
  End If
  
  If cpValor.Text = "" Then
    MsgBox "Valor inválido.", vbCritical + vbOKOnly, "Aviso"
    cpValor.SetFocus
    GoTo Fim
  End If
  
  If CDbl(cpValor.Text) <= 0# Then
    MsgBox "O valor deve ser maior que zero.", vbCritical + vbOKOnly, "Aviso"
    cpValor.SetFocus
    GoTo Fim
  End If
  
'  cmdCancelar.Enabled = False
'  cmdReimpime.Enabled = False
'  cmdEtiquetas.Enabled = False
  
  If DebAnt = True Then
    Resp = MsgBox("ATENÇÃO:" & Chr(10) & "Este boleto está sendo gerado através de débitos anteriores que serão excluidos para dar lugar a um novo. Continuar?", vbQuestion + vbYesNo, "Aviso")
    If Resp = vbNo Then
      Exit Sub
    End If
  End If
  
  Set selBoletos = db.OpenRecordset("boletos", dbOpenTable)
  
  Set selCondominio = db.OpenRecordset("Select * From CONDOMINIO where codigo = " & cCondominio & " order by nome;", dbOpenDynaset)
  
  DBEngine.BeginTrans
  With selCondominio
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      
      nBoleto = !ultboleto + 1
      If nBoleto > 9999 Then
        nBoleto = 1
      End If
      
      dtVenc = CDate(cpVenc.Text)
      
      SqlStr = "Select * From Associados Where Condominio = " & !Codigo & " and codigo = " & cpCodigo.Text & " Order By Proprietario"
      Set SelAsso = db.OpenRecordset(SqlStr, dbOpenDynaset)
      
      With SelAsso
        If .RecordCount > 0 Then
          .MoveFirst
            
          nVal = cpValor.Text
          
          If nVal > 0 Then
            
            fValor = Format(nVal, "#0.00")
            fValor = Left(fValor, InStr(fValor, ",") - 1) & Right(fValor, 2)
            fValor = PadLeft(fValor, "0", 10)
            nFator = CStr(dtVenc - CDate("07/10/1997"))
            
            If selCondominio!tipoboleto = 1 Then
              NossoNumero = Trim(selCondominio!carteira) & Format$(cpVenc.Text, "MMyy")
              NossoNumero = NossoNumero & PadLeft(!Codigo, "0", 5)
              NossoNumero = NossoNumero & PadLeft(nBoleto, "0", 6)
              
              sLivre = Trim(selCondominio!conta) & DigitoNosso(Trim(selCondominio!conta))
              sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
              sLivre = sLivre & Mid$(NossoNumero, 6, 3)
              sLivre = sLivre & Mid$(NossoNumero, 2, 1)
              sLivre = sLivre & Mid$(NossoNumero, 9)
                  
              sLivre = sLivre & DigitoNosso(sLivre)
            ElseIf selCondominio!tipoboleto = 2 Then
              
              sBol = PadLeft(cCondominio, "0", 4) & PadLeft(nBoleto, "0", 4)
              NossoNumero = Trim(selCondominio!carteira) & PadLeft(sBol, "0", 10 - Len(Trim(selCondominio!carteira)))
              sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & PadLeft(selCondominio!conta, "0", 8)
            
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
            tbBoletos!MENS = nVal
            tbBoletos!EXTR = 0
            tbBoletos!Data = Format(Date, "dd/mm/yyyy")
            tbBoletos!valr = nVal
            tbBoletos!corrigido = nVal
            tbBoletos!cdsc = !Codigo
            If !boleto = 1 Then
              tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
              tbBoletos!Nome = NomeCompleto(!Codigo)
              tbBoletos!cpf = !pcpf
              tbBoletos!Ende = !PEndereco
              tbBoletos!Bair = !PBairro
              tbBoletos!Cida = !PCidade
              tbBoletos!Esta = !PEstado
              tbBoletos!cep = !PCep
            Else
              If !Proprietario = "" Then
                tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
                tbBoletos!Nome = NomeCompleto(!Codigo)
                tbBoletos!cpf = !pcpf
                tbBoletos!Ende = !PEndereco
                tbBoletos!Bair = !PBairro
                tbBoletos!Cida = !PCidade
                tbBoletos!Esta = !PEstado
                tbBoletos!cep = !PCep
              Else
                tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
                tbBoletos!Nome = NomeCompleto(!Codigo)
                tbBoletos!cpf = !cpf
                tbBoletos!Ende = !endereco
                tbBoletos!Bair = !bairro
                tbBoletos!Cida = !Cidade
                tbBoletos!Esta = !estado
                tbBoletos!cep = !cep
              End If
            End If
            tbBoletos!tran = "AV" & sBol   'PadLeft(nBoleto, "0", 8)
            tbBoletos!CDBARRA = sBarra 'Bar25I(sBarra)
            tbBoletos!DIGITAVAL = NossoNumero & "." & DigitoNosso(NossoNumero)
            tbBoletos!agcedente = selCondominio!agcedente
            tbBoletos!carteira = selCondominio!carteira
            If selCondominio!tipoboleto = 1 Then
              tbBoletos!Mensagem = selCondominio!agcedente & "/" & selCondominio!conta & "-" & DigitoNosso(selCondominio!conta)
            Else
              tbBoletos!Mensagem = Trim(selCondominio!agcedente) & "." & Trim(selCondominio!Operacao) & "." & Trim(selCondominio!conta) & "." & DigitoNosso(Trim(selCondominio!agcedente) & Trim(selCondominio!Operacao) & Trim(selCondominio!conta))
            End If
            tbBoletos!CODI = Campo1 & " " & Campo2 & " " & Campo3 & " " & Campo4 & " " & Campo5
            tbBoletos!texto = cpHistorico.Text & vbCrLf & vbCrLf & cpMens1.Text & vbCrLf & cpMens2.Text & vbCrLf & cpMens3.Text & vbCrLf & cpMens4.Text & vbCrLf & cpMens5.Text & vbCrLf & cpMens6.Text
            tbBoletos!INST1 = cpMensagem(0).Text
            tbBoletos!INST2 = cpMensagem(1).Text
            tbBoletos!INST3 = cpMensagem(2).Text
            tbBoletos!INST4 = cpMensagem(3).Text
            tbBoletos!Banco = "CAIXA" 'selCondominio!Banco
            tbBoletos!CdBanco = selCondominio!CdAgencia
            tbBoletos!pago = "N"
            tbBoletos!CANCELADO = "N"
            tbBoletos!cond = selCondominio!Codigo
            tbBoletos!acumulado = !acumulado
            tbBoletos!bole = NossoNumero
            tbBoletos!nosso = NossoNumero
            tbBoletos!idStatus = 1
            tbBoletos.Update
            sLivre = "AV" & sBol   'PadLeft(nBoleto, "0", 8)
            DoEvents
          End If
        End If
      End With
      .Edit
      !ultboleto = nBoleto
      .Update
      If DebAnt = True Then
        If VetorIniciado(mRetDeb) Then
          For i = 0 To UBound(mRetDeb)
            db.Execute "delete from boletos where id = " & mRetDeb(i) & ";"
          Next i
        End If
      End If
    End If
  End With
  
  DBEngine.CommitTrans
  sSql = "Select * from boletos where cdsc = " & cpCodigo.Text & " and nosso = '" & NossoNumero & "' order by nome;"

  RelatoriosRPT.mnuSupExport.Visible = False
  RelatoriosRPT.mnuExportBoleto.Visible = True
  RelatoriosRPT.mnuEviarEmail.Visible = True
  RelatoriosRPT.Carregar "", Parametros.dados, "{boletos.cdsc} = " & cpCodigo.Text & " and {boletos.nosso} = '" _
      & NossoNumero & "'", "Boletos", sFormataCaminho(App.Path) & "avulso.rpt", , sSql

Fim:
'  cmdCancelar.Enabled = True
'  cmdReimpime.Enabled = True
'  cmdEtiquetas.Enabled = True
  Exit Sub

Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  DBEngine.Rollback
  Resume Fim
  
End Sub

Private Sub cmdEtiquetas_Click()
  cmdEtiquetas.Enabled = False
  cmdCancelar.Enabled = False
  cmdReimpime.Enabled = False
  DoEvents
  Montar
  DebAnt = False
  cmdEtiquetas.Enabled = True
  cmdCancelar.Enabled = True
  cmdReimpime.Enabled = True
End Sub

Private Sub cmdLocalizar_Click()
  cpCodigo.Text = ""
  cpCodigo_KeyPress 13
End Sub

Private Sub cpCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMens1.SetFocus
  End If
End Sub

Private Sub AcertaMensagem(nCod As Long)
  If nCod <= 0 Then
    cpMensagem(0).Text = "SR CAIXA, NÃO RECEBER APÓS 15 DIAS DE VENCIDO"
    cpMensagem(1).Text = "APÓS VENCIMENTO SÓ RECEBER"
    cpMensagem(2).Text = "COM JUROS DE " & Parametros.juros & "% AO MÊS + MULTA DE " & Parametros.Multa & "%"
    cpMensagem(3).Text = ""
  Else
    If PrazoBoleto(nCod) = 0 Then
      cpMensagem(0).Text = "SR CAIXA"
    Else
      cpMensagem(0).Text = "SR CAIXA, NÃO RECEBER APÓS " & PegaDias(nCod) & " DIAS DE VENCIDO"
    End If
    cpMensagem(1).Text = "APÓS VENCIMENTO SÓ RECEBER"
    cpMensagem(2).Text = "COM JUROS DE " & PegaJuros(nCod) & _
          "% AO MÊS + MULTA DE " & PegaMulta(nCod) & "%"
    cpMensagem(3).Text = ""
  End If
End Sub

Private Sub cmdReimpime_Click()
  ReAvulsa.Show
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodigo.Text) > 0 Then
      With tbAssociados
        .Index = "codigoid"
        .Seek "=", Val(cpCodigo.Text)
        If Not .NoMatch Then
          If !boleto = 0 Then
            cpNome.Text = !Tipo & " " & !Apartamento & " " & !Proprietario
          Else
            cpNome.Text = !Tipo & " " & !Apartamento & " " & !Nome
          End If
          AcertaMensagem !Condominio
          cCondominio = !Condominio
          cpHistorico.SetFocus
        Else
          MsgBox "Código não encontrado!", vbInformation + vbOKOnly, "Aviso"
        End If
      End With
    Else
      RetCodigo = 0
      lAssociado.Show 1
      If RetCodigo > 0 Then
        With tbAssociados
          .Index = "codigoid"
          .Seek "=", RetCodigo
          If Not .NoMatch Then
            cpCodigo.Text = RetCodigo
            If !boleto = 0 Then
              cpNome.Text = !Tipo & " " & !Apartamento & " " & !Proprietario
            Else
              cpNome.Text = !Tipo & " " & !Apartamento & " " & !Nome
            End If
            AcertaMensagem !Condominio
            cCondominio = !Condominio
            cpHistorico.SetFocus
          End If
        End With
      End If
    End If
  End If
End Sub

Private Sub cpHistorico_GotFocus()
  cpHistorico.SelStart = 0
  cpHistorico.SelLength = Len(cpHistorico.Text)
End Sub

Private Sub cpHistorico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpVenc.SetFocus
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
        cpCodigo.SetFocus
      End If
    Case Else
      KeyAscii = vTexto(KeyAscii)
  End Select
End Sub

Private Sub cpNome_GotFocus()
  cpNome.SelStart = 0
  cpNome.SelLength = Len(cpNome.Text)
End Sub

Private Sub cpNome_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpVenc.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpValor_GotFocus()
  cpValor.SelStart = 0
  cpValor.SelLength = Len(cpValor.Text)
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdEtiquetas.SetFocus
  End If
End Sub

Private Sub cpVenc_GotFocus()
  cpVenc.SelStart = 0
  cpVenc.SelLength = Len(cpVenc.Text)
End Sub

Private Sub cpVenc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
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
  Set tbBoletos = db.OpenRecordset("boletos", dbOpenTable)
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  Refresh
  KeyPreview = True
  DebAnt = False
  cpMensagem(0).Text = "SR CAIXA, NÃO RECEBER APÓS 15 DIAS DE VENCIDO"
  cpMensagem(1).Text = "APÓS VENCIMENTO SÓ RECEBER"
  cpMensagem(2).Text = "COM JUROS DE " & Parametros.juros & "% AO MÊS + MULTA DE " & Parametros.Multa & "%"
  cpMensagem(3).Text = ""
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

Private Function PegaJuros(ByVal nCod As Long) As Double
  Dim rs As Recordset
  Dim sRet As Double
  Set rs = db.OpenRecordset("Select * From CONDOMINIO where codigo = " & nCod & " order by nome;", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      sRet = !juros
    Else
      sRet = Parametros.juros
    End If
  End With
  Set rs = Nothing
  PegaJuros = sRet
End Function

Private Function PegaMulta(ByVal nCod As Long) As Double
  Dim rs As Recordset
  Dim sRet As Double
  Set rs = db.OpenRecordset("Select * From CONDOMINIO where codigo = " & nCod & " order by nome;", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      sRet = !Multa
    Else
      sRet = Parametros.Multa
    End If
  End With
  Set rs = Nothing
  PegaMulta = sRet
End Function

Private Function VetorIniciado(vVetor() As Long) As Boolean
    Dim aux As Integer
    '
    On Error Resume Next
    aux = UBound(vVetor)
    '
    If Err.Number = 0 Then
        VetorIniciado = True
    Else
        VetorIniciado = False
    End If
    '
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set tbBoletos = Nothing
  Set tbAssociados = Nothing
End Sub
