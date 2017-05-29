VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelBolAt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de boletos em atraso"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "RelBolAt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   6120
      TabIndex        =   3
      Top             =   480
      Width           =   1080
   End
   Begin rdActiveText.ActiveText cpDias 
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Top             =   480
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   3
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2370
      TabIndex        =   6
      Top             =   60
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "dias em atraso"
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   540
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mostrar os boletos com mais de"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   2220
   End
End
Attribute VB_Name = "RelBolAt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dtMax As Date
Dim tbCondominio As Recordset

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
  Dim sOrdem As String
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  'cmdPrint.Enabled = False
  
  dtMax = DateAdd("d", (Val(cpDias.Text) * -1), Date)
  
  MontaRelatorio
  DoEvents
  
  sOrdem = ""  '"{relextrato.vencimento};crAscendingOrder;relextrato|"
  RelatoriosRPT.Carregar sOrdem, Parametros.dados, "", "Boletos com mais de " & cpDias.Text & " dias de atraso do condominio " & cpNome.Text, sFormataCaminho(App.Path) & "boletoatrazado.rpt", , , , "relextrato", , , , "subatrasados"
  'cmdPrint.Enabled = True
  
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
          cpDias.SetFocus
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
        cpDias.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpDias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdPrint.Enabled Then
      cmdPrint.SetFocus
    End If
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
  Refresh
  KeyPreview = True
End Sub

Private Sub MontaRelatorio()
  Dim rs As Recordset
  Dim rsBol As Recordset
  Dim rsrel As Recordset
  Dim rsDet As Recordset
  Dim lcDet As Recordset
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Escolha um condomínio.", vbInformation + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  dbLocal.Execute "delete from relextrato;"
  dbLocal.Execute "delete from boletodetalhe;"
  Set rsrel = dbLocal.OpenRecordset("relextrato", dbOpenTable)
  Set lcDet = dbLocal.OpenRecordset("boletodetalhe", dbOpenTable)
  Set rs = db.OpenRecordset("select * from associados where condominio = " _
    & cpCodigo.Text & " order by codigo;", dbOpenDynaset)
  
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
'          Set rsBol = db.OpenRecordset("select * from boletos where cdsc = " & !Codigo & " and (PAGO <> 'S' Or PAGO is null) and vcto <= #" _
'                & Format$(Date, "MM/dd/yyyy") & "# order by vcto;", dbOpenDynaset)
'          If rsBol.RecordCount > 0 Then
'            rsBol.MoveFirst
'            If rsBol!VCTO <= dtMax Then
'              Do While Not rsBol.EOF
'                rsrel.AddNew
'                rsrel!id_associado = !Codigo
'                rsrel!id_condominio = cpCodigo.Text
'                rsrel!COTA = rsBol!TRAN
'                rsrel!data = rsBol!data
'                rsrel!valor = rsBol!valr
'                rsrel!juros = rsBol!corrigido - rsBol!valr
'                rsrel!Vencimento = rsBol!VCTO
'                rsrel!nosso_numero = rsBol!DIGITAVAL
'                rsrel!somar = 1
'                rsrel.Update
'                rsBol.MoveNext
'              Loop
'            End If
'          End If  'and cdbarra = ''
'          Set rsBol = Nothing
'          Set rsBol = db.OpenRecordset("select * from boletos where left(tran,2) = 'GE' and cdsc = " & !Codigo & " and (PAGO <> 'S' Or PAGO is null) and vcto <= #" _
'                & Format$(Date, "MM/dd/yyyy") & "# order by vcto;", dbOpenDynaset)
'          If rsBol.RecordCount > 0 Then
'            rsBol.MoveFirst
'            If rsBol!VCTO <= dtMax Then
'              Do While Not rsBol.EOF
'                rsrel.AddNew
'                rsrel!id_associado = !Codigo
'                rsrel!id_condominio = cpCodigo.Text
'                rsrel!COTA = rsBol!TRAN
'                rsrel!data = rsBol!data
'                rsrel!valor = rsBol!valr
'                rsrel!juros = rsBol!corrigido - rsBol!valr
'                rsrel!Vencimento = rsBol!VCTO
'                rsrel!nosso_numero = rsBol!DIGITAVAL
'                rsrel!somar = 1
'                rsrel.Update
'                rsBol.MoveNext
'              Loop
'            End If
'          End If  'and cdbarra = ''
'          Set rsBol = Nothing -left(tran,2) <> 'AV' and
          Set rsBol = db.OpenRecordset("select * from boletos where cancelado = 'N' and cdsc = " & !Codigo & " and PAGO <> 'S' and vcto <= #" _
                & Format$(Date, "MM/dd/yyyy") & "# order by vcto;", dbOpenDynaset)
          With rsBol
            If .RecordCount > 0 Then
              .MoveFirst
              If !vcto <= dtMax Then
                Do While Not .EOF
                  If (IsNumeric(Left(!CODI, 5))) Then
                    rsrel.AddNew
                    rsrel!id_associado = !cdsc
                    rsrel!id_condominio = cpCodigo.Text
                    rsrel!COTA = !tran
                    rsrel!Data = !Data
                    rsrel!valor = !valr
                    rsrel!Historico = NomeCompleto(!cdsc)
                    rsrel!juros = !corrigido - rsBol!valr
                    rsrel!Vencimento = !vcto
                    rsrel!nosso_numero = !DIGITAVAL
                    rsrel!somar = 1
                    rsrel!id_boleto = !id
                    If IsNumeric(Left(!tran, 2)) Then
                      If rs!acumulado = "S" Then
                        If (.AbsolutePosition + 1) < .RecordCount And .RecordCount > 1 Then
                          rsrel!somar = 0
                        Else
                          rsrel!somar = 1
                        End If
                      Else
                        rsrel!somar = 1
                      End If
                    End If
                    rsrel.Update
                    Set rsDet = db.OpenRecordset("select * from boletodetalhe where id_boleto = " & !id & ";", dbOpenDynaset)
                    If rsDet.RecordCount > 0 Then
                      rsDet.MoveFirst
                      Do While Not rsDet.EOF
                        lcDet.AddNew
                        lcDet!Condominio = rsDet!Condominio
                        lcDet!mes = rsDet!mes
                        lcDet!Descricao = rsDet!Descricao
                        lcDet!GERAIS = rsDet!GERAIS
                        lcDet!Sindico = rsDet!Sindico
                        lcDet!Proprietario = rsDet!Proprietario
                        lcDet!Fracao = rsDet!Fracao
                        lcDet!valor = rsDet!valor
                        lcDet!BOX = rsDet!BOX
                        lcDet!id_associado = rsDet!id_associado
                        lcDet!id_condominio = rsDet!id_condominio
                        lcDet!VALOR_PROPRIETARIO = rsDet!VALOR_PROPRIETARIO
                        lcDet!id_boleto = rsDet!id_boleto
                        lcDet.Update
                        rsDet.MoveNext
                      Loop
                    End If
                    .MoveNext
                  End If
                Loop
              End If
            End If
          End With
          Set rsBol = Nothing
        .MoveNext
      Loop
    Else
      MsgBox "Este condomínio não possui inquilinos.", vbInformation + vbOKOnly, "Aviso"
    End If
  End With
  Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
