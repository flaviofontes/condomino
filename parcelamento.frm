VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Parcelamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parcelamento de débitos"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "parcelamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lsParcelas 
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   4260
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descrição"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vencimento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsDebitos 
      Height          =   2115
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   3731
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Referência"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vencimento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Juros"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Histórico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdConfirmacao 
      BackColor       =   &H0000C000&
      Caption         =   "Confirmar parcelamento"
      Height          =   375
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6060
      Width           =   2055
   End
   Begin rdActiveText.ActiveText cpDataVencimento 
      Height          =   315
      Left            =   5220
      TabIndex        =   6
      Top             =   3840
      Width           =   1515
      _ExtentX        =   2672
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
   Begin VB.CommandButton cmdGerarParcelas 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   7020
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox cpParcelas 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CheckBox ChLiquido 
      Caption         =   "Valor líquido"
      Height          =   195
      Left            =   7140
      TabIndex        =   4
      Top             =   3300
      Width           =   1215
   End
   Begin VB.TextBox cpTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3180
      Width           =   1395
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
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Top             =   510
      Width           =   6015
   End
   Begin VB.TextBox cpCota 
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
      Left            =   810
      TabIndex        =   0
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data vencimento 1ª parcela"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   3900
      Width           =   1980
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Número de parcelas"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      X1              =   60
      X2              =   8340
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   5160
      TabIndex        =   11
      Top             =   3300
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inquilino"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "Parcelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim vaSoma      As Double
Dim vaSelect    As Double
Dim vaJuros     As Double
Dim tbAssociados As Recordset
Dim tbBoletos As Recordset

Private Sub ChLiquido_Click()
  Dim ValMens As Double
  Dim i As Integer
  Dim MaxParc As Integer
  vaSelect = 0
  If ChLiquido.Value = 1 Then
    For i = 1 To lsDebitos.ListItems.Count
      If lsDebitos.ListItems.Item(i).Checked = True Then
        vaSelect = vaSelect + CDbl(lsDebitos.ListItems.Item(i).SubItems(2))
      End If
    Next i
    cpTotal.Text = Format$(vaSelect, "#,##0.00")
  Else
    For i = 1 To lsDebitos.ListItems.Count
      If lsDebitos.ListItems.Item(i).Checked = True Then
        vaSelect = vaSelect + (CDbl(lsDebitos.ListItems.Item(i).SubItems(2)) + CDbl(lsDebitos.ListItems.Item(i).SubItems(3)))
      End If
    Next i
    cpTotal.Text = Format(vaSelect, "#,##0.00")
  End If
End Sub

Private Sub cmdConfirmacao_Click()
  Dim i As Integer
  Dim vParcela As Double
  Dim rs As Recordset
  Dim Sql As String
  Dim Item As MSComctlLib.ListItem
  Dim prints() As Long
  Dim sqlCrystal As String
  
  If lsParcelas.ListItems.Count = 0 Then
    MsgBox "Gere as parcelas primeiro.", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  Resp = MsgBox("Este processo irá eliminar o(s) débito(s) selecionado(s) e gerar novos conforme o parcelamento. Prosseguir?", vbQuestion + vbYesNo, "Parcelar")
  If Resp = vbYes Then
    ReDim prints(lsParcelas.ListItems.Count - 1) As Long
    i = 0
    For Each Item In lsParcelas.ListItems
      prints(i) = Acrescenta(Val(cpCota.Text), Item.Text, Item.SubItems(1), Item.SubItems(2))
      i = i + 1
    Next
    For i = 1 To lsDebitos.ListItems.Count
      If lsDebitos.ListItems.Item(i).Checked = True Then
        'saber o status do boleto
        Sql = "select * From BOLETOS Where id = " & lsDebitos.ListItems.Item(i).Tag & ";"
        Set rs = db.OpenRecordset(Sql, dbOpenDynaset)
        With rs
          .MoveFirst
          Select Case !idStatus
            Case 1, 5, 7, 8
              db.Execute "delete from boletodetalhe where id_boleto = " & !id & ";"
              .Delete
            Case 2, 4, 6
              .Edit
              !idStatus = 9
              .Update
          End Select
        End With
'        Sql = "delete From [BOLETOS] Where [CDSC] = " + cpCota.Text + " And ([PAGO] <> 'S' Or [PAGO] Is Null) " _
'              + " and id = " & lsDebitos.ListItems.Item(i).Tag & ";"
'        db.Execute Sql
      End If
    Next i
    lsParcelas.ListItems.Clear
    Gerar
    Sql = ""
    If VetorLongIniciado(prints) Then
      For i = 0 To UBound(prints)
        Sql = Sql & "id = " & prints(i) & " Or "
      Next i
      
      If Right(Sql, 4) = " Or " Then
        Sql = Left(Sql, Len(Sql) - 4)
      End If
      
      sqlCrystal = Replace(Sql, "id", "{boletos.ID}")
      
      Sql = "Select * from boletos where " & Sql & " order by nome;"
      RelatoriosRPT.mnuSupExport.Visible = False
      RelatoriosRPT.mnuExportBoleto.Visible = True
      RelatoriosRPT.mnuEviarEmail.Visible = True
      RelatoriosRPT.Carregar "", Parametros.dados, sqlCrystal, "Boletos", sFormataCaminho(App.Path) & "avulso.rpt", , Sql
    End If
  End If
End Sub

Private Sub cmdGerarParcelas_Click()
  Dim i As Integer
  Dim Maxp As Integer
  Dim vParcela As Double
  Dim pParcela As Double
  Dim tParcela As Double
  Dim dVencimento As Date
  Dim nItem As MSComctlLib.ListItem
  
  Maxp = Val(cpParcelas.Text)
  
  If Maxp = 0 Then
    MsgBox "Escolha o número de parcelas.", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  If Not IsDate(cpDataVencimento.Text) Then
    MsgBox "Informe a data de vencimento da 1ª parcela.", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  If CDate(cpDataVencimento.Text) < Date Then
    MsgBox "A data de vencimento da 1ª parcela deve ser igual ou maior que " & Format$(Date, "dd/MM/yyyy") & ".", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  lsParcelas.ListItems.Clear
  
  vParcela = Round(cpTotal.Text / Maxp, 2)
  Maxp = Maxp - 1
  tParcela = vParcela * Maxp
  pParcela = cpTotal.Text - tParcela
  dVencimento = cpDataVencimento.Text
  
  Set nItem = lsParcelas.ListItems.Add(, "P" & 1)
  nItem.Text = "1ª PARCELA - DÉBITOS EM ATRAZO   "
  nItem.SubItems(1) = Format$(dVencimento, "dd/MM/yyyy")
  nItem.SubItems(2) = Format$(pParcela, "#,##0.00")
  For i = 2 To Maxp + 1
    dVencimento = DateAdd("m", 1, dVencimento)
    Set nItem = lsParcelas.ListItems.Add(, "P" & i)
    nItem.Text = i & "ª PARCELA - DÉBITOS EM ATRAZO   "
    nItem.SubItems(1) = Format$(dVencimento, "dd/MM/yyyy")
    nItem.SubItems(2) = Format$(vParcela, "#,##0.00")
  Next i
  Call AutoAjusteListView2(lsParcelas, 0)
End Sub

Private Sub cpCota_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Len(cpCota.Text) > 0 Then
      With tbAssociados
        .Index = "Codigoid"
        .Seek "=", Val(cpCota.Text)
        If Not .NoMatch Then
          cpNome.Text = NomeCompleto(!Codigo)
          Call Gerar
        Else
          Beep
          MsgBox "Código não encontrado!", vbInformation + vbOKOnly, "Aviso"
        End If
      End With
    Else
      RetCodigo = 0
      lAssociado.Show 1
      If RetCodigo > 0 Then
        With tbAssociados
          .Index = "Codigoid"
          .Seek "=", RetCodigo
          If Not .NoMatch Then
            cpNome.Text = NomeCompleto(!Codigo)
            cpCota.Text = !Codigo
            Gerar
          End If
        End With
      End If
    End If
  End If
End Sub

Private Function Gerar()
On Error GoTo Errado
  Dim StrSel2 As String
  Dim NumBole As String
  Dim ValMens As Double
  Dim i As Integer
  Dim MaxParc As Integer
  Dim nNode As MSComctlLib.ListItem
  Dim rs As Recordset
  
  StrSel2 = "Select boletos.*, statusboleto.descricao From BOLETOS inner join statusboleto on statusboleto.idstatus = boletos.idstatus Where CDSC = " + cpCota.Text + " And ([PAGO] <> 'S' Or [PAGO] Is Null) Order By VCTO"
  Set rs = db.OpenRecordset(StrSel2, dbOpenDynaset)
  DoEvents
  
  vaSoma = 0
  vaJuros = 0
  
  DBEngine.Idle dbRefreshCache
  lsDebitos.ListItems.Clear
  With rs
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      While Not .EOF
        Set nNode = lsDebitos.ListItems.Add(, "A" & .AbsolutePosition)
        nNode.Text = !tran & "  "
        nNode.SubItems(1) = Format$(!vcto, "dd/MM/yyyy")
        nNode.SubItems(2) = Format$(!valr, "#,##0.00")
        nNode.SubItems(3) = Format$(!corrigido - !valr, "#,##0.00")
        nNode.SubItems(4) = Replace(!texto & "  ", vbCrLf, "")
        nNode.SubItems(5) = RetornaStatus(!idStatus)
        nNode.Checked = True
        nNode.Tag = !id
        vaSoma = vaSoma + IIf(IsNull(!valr), 0, !valr)
        vaJuros = vaJuros + (!corrigido - !valr)
        DoEvents
        .MoveNext
      Wend
    End If
    If .RecordCount > 0 Then
      .MoveFirst
      Call AutoAjusteListView2(lsDebitos, 0)
      Call AutoAjusteListView2(lsDebitos, 1)
      Call AutoAjusteListView2(lsDebitos, 4)
    End If
  End With
  
  cpTotal.Text = Format(vaSoma + vaJuros, "#,##0.00")
  cpParcelas.Clear
  If cpTotal.Text > 0 Then
    MaxParc = 12
    For i = 2 To MaxParc
      cpParcelas.AddItem (Format$(i, "00"))
    Next i
  End If
  DoEvents

Fim:
  Set rs = Nothing
  Exit Function
  
Errado:
  MsgBox "Erro No. " + Str(Err.Number) + vbCrLf + Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Fim

End Function

Private Sub cpDataVencimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGerarParcelas.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbAssociados = db.OpenRecordset("associados", dbOpenTable)
  Set tbBoletos = db.OpenRecordset("boletos", dbOpenTable)
  Refresh
  DoEvents
  KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbAssociados = Nothing
  Set tbBoletos = Nothing
End Sub

Private Sub lsDebitos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  vaSelect = CDbl(cpTotal.Text)
  If Item.Checked = True Then
    If ChLiquido.Value = 0 Then
      vaSelect = vaSelect + (CDbl(Item.SubItems(2)) + CDbl(Item.SubItems(3)))
    Else
      vaSelect = vaSelect + CDbl(Item.SubItems(2))
    End If
  Else
    If ChLiquido.Value = 0 Then
      vaSelect = vaSelect - (CDbl(Item.SubItems(2)) + CDbl(Item.SubItems(3)))
    Else
      vaSelect = vaSelect - CDbl(Item.SubItems(2))
    End If
  End If
  cpTotal.Text = Format$(vaSelect, "#,##0.00")
End Sub

Private Function Acrescenta(ByVal nCod As Long, ByVal sHist As String, ByVal vcto As Date, ByVal nValor As Double) As Long
  Dim SelAsso As Recordset
  Dim SelCond As Recordset
  
  Dim sBarra  As String
  Dim nFator  As String * 4
  Dim fValor  As String
  Dim vbDig   As String * 1
  Dim Campo1  As String
  Dim Campo2  As String
  Dim Campo3  As String
  Dim Campo4  As String
  Dim Campo5  As String
  Dim NossoNumero As String
  Dim sLivre As String
  Dim nBoleto As Long
  Dim sBol As String


  Set SelAsso = db.OpenRecordset("Select * from associados where codigo = " & nCod & ";", dbOpenDynaset)
  With SelAsso
    If .RecordCount > 0 Then
      .MoveFirst
      Set SelCond = db.OpenRecordset("Select * from condominio where codigo = " & !Condominio & ";", dbOpenDynaset)
    End If
  End With
  
  With SelCond
    If .RecordCount > 0 Then
      .MoveFirst
      
      nBoleto = !ultboleto + 1
      If nBoleto > 9999 Then
        nBoleto = 1
      End If

      With SelAsso
        If .RecordCount > 0 Then
          .MoveFirst
            
          If nValor > 0 Then
            
            fValor = Format(nValor, "#0.00")
            fValor = Left(fValor, InStr(fValor, ",") - 1) & Right(fValor, 2)
            fValor = PadLeft(fValor, "0", 10)
            nFator = CStr(vcto - CDate("07/10/1997"))
            
            If SelCond!tipoboleto = 1 Then
              NossoNumero = Trim(SelCond!carteira) & Format$(vcto, "MMyy")
              NossoNumero = NossoNumero & PadLeft(!Codigo, "0", 5)
              NossoNumero = NossoNumero & PadLeft(nBoleto, "0", 6)
              
              sLivre = Trim(SelCond!Conta) & DigitoNosso(Trim(SelCond!Conta))
              sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
              sLivre = sLivre & Mid$(NossoNumero, 6, 3)
              sLivre = sLivre & Mid$(NossoNumero, 2, 1)
              sLivre = sLivre & Mid$(NossoNumero, 9)
                  
              sLivre = sLivre & DigitoNosso(sLivre)
            ElseIf SelCond!tipoboleto = 2 Then
              
              sBol = PadLeft(SelCond!Codigo, "0", 4) & PadLeft(nBoleto, "0", 4)
              NossoNumero = Trim(SelCond!carteira) & PadLeft(sBol, "0", 10 - Len(Trim(SelCond!carteira)))
              sLivre = NossoNumero & SelCond!agcedente & SelCond!Operacao & PadLeft(SelCond!Conta, "0", 8)
            
            End If
                
            sBarra = Left(SelCond!CdAgencia, 3) & "9" & nFator & fValor & sLivre
                      
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
            If SelCond!titularboleto = 1 Then
              tbBoletos!Condominio = SelCond!Nome
              tbBoletos!CGC = SelCond!CGC
            Else
              If SelCond!razaoboleto & "" <> "" Then
                tbBoletos!Condominio = SelCond!razaoboleto
                tbBoletos!CGC = SelCond!cnpjboleto
              Else
                tbBoletos!Condominio = SelCond!Nome
                tbBoletos!CGC = SelCond!CGC
              End If
            End If
            tbBoletos!vcto = vcto
            tbBoletos!MENS = nValor
            tbBoletos!EXTR = 0
            tbBoletos!Data = Format(Date, "dd/mm/yyyy")
            tbBoletos!valr = nValor
            tbBoletos!corrigido = nValor
            tbBoletos!cdsc = !Codigo
            If !boleto = 1 Then
              tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
              tbBoletos!Nome = NomeCompleto(!Codigo)
              tbBoletos!Cpf = !PCpf
              tbBoletos!Ende = !PEndereco
              tbBoletos!Bair = !PBairro
              tbBoletos!Cida = !PCidade
              tbBoletos!Esta = !PEstado
              tbBoletos!cep = !PCep
            Else
              If !Proprietario = "" Then
                tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
                tbBoletos!Nome = NomeCompleto(!Codigo)
                tbBoletos!Cpf = !PCpf
                tbBoletos!Ende = !PEndereco
                tbBoletos!Bair = !PBairro
                tbBoletos!Cida = !PCidade
                tbBoletos!Esta = !PEstado
                tbBoletos!cep = !PCep
              Else
                tbBoletos!COTA = PadLeft(!Codigo, "0", 4)
                tbBoletos!Nome = NomeCompleto(!Codigo)
                tbBoletos!Cpf = !Cpf
                tbBoletos!Ende = !endereco
                tbBoletos!Bair = !bairro
                tbBoletos!Cida = !Cidade
                tbBoletos!Esta = !estado
                tbBoletos!cep = !cep
              End If
            End If
            tbBoletos!tran = "PR" & sBol   'PadLeft(nBoleto, "0", 8)
            tbBoletos!CDBARRA = sBarra     'Bar25I(sBarra)
            tbBoletos!DIGITAVAL = NossoNumero & "." & DigitoNosso(NossoNumero)
            tbBoletos!agcedente = SelCond!agcedente
            tbBoletos!carteira = SelCond!carteira
            If SelCond!tipoboleto = 1 Then
              tbBoletos!Mensagem = SelCond!agcedente & "/" & SelCond!Conta & "-" & DigitoNosso(SelCond!Conta)
            Else
              tbBoletos!Mensagem = Trim(SelCond!agcedente) & "." & Trim(SelCond!Operacao) & "." & Trim(SelCond!Conta) & "." & DigitoNosso(Trim(SelCond!agcedente) & Trim(SelCond!Operacao) & Trim(SelCond!Conta))
            End If
            tbBoletos!CODI = Campo1 & " " & Campo2 & " " & Campo3 & " " & Campo4 & " " & Campo5
            tbBoletos!texto = sHist
            
            If PrazoBoleto(SelCond!Codigo) = 0 Then
              tbBoletos!INST1 = "SR CAIXA"
            Else
              tbBoletos!INST1 = "SR CAIXA, NÃO RECEBER APÓS " & PegaDias(SelCond!Codigo) & " DIAS DE VENCIDO"
            End If
            tbBoletos!INST2 = "APÓS VENCIMENTO SÓ RECEBER"
            tbBoletos!INST3 = "COM JUROS DE " & PegaJuros(SelCond!Codigo) & _
                  "% AO MÊS + MULTA DE " & PegaMulta(SelCond!Codigo) & "%"
            tbBoletos!INST4 = ""
            tbBoletos!Banco = "CAIXA" 'selcond!Banco
            tbBoletos!CdBanco = SelCond!CdAgencia
            tbBoletos!pago = "N"
            tbBoletos!CANCELADO = "N"
            tbBoletos!cond = SelCond!Codigo
            tbBoletos!acumulado = !acumulado
            tbBoletos!bole = NossoNumero
            tbBoletos!nosso = NossoNumero
            tbBoletos!idStatus = 1
            tbBoletos.Update
            tbBoletos.Bookmark = tbBoletos.LastModified
            Acrescenta = tbBoletos!id
            sLivre = "AV" & sBol   'PadLeft(nBoleto, "0", 8)
            DoEvents
          End If
        End If
      End With
      .Edit
      !ultboleto = nBoleto
      .Update
    End If
  End With
End Function

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

Private Sub lsParcelas_BeforeLabelEdit(Cancel As Integer)
  Call AutoAjusteListView2(lsParcelas, 0)
End Sub
