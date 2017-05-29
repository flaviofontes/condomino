VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelRecebidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de boletos quitados"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "RelRecebidos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin rdActiveText.ActiveText cpFim 
      Height          =   315
      Left            =   3540
      TabIndex        =   3
      Top             =   480
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
   Begin rdActiveText.ActiveText cpInicio 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1155
      _ExtentX        =   2037
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   6060
      TabIndex        =   4
      Top             =   480
      Width           =   1140
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2340
      TabIndex        =   8
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   195
      Left            =   3180
      TabIndex        =   6
      Top             =   540
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quitados no período de"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1680
   End
End
Attribute VB_Name = "RelRecebidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
  
  If Not IsDate(cpInicio.Text) Then
    MsgBox "A data inicial não foi informada ou não é válida.", vbInformation + vbOKOnly, "Aviso"
    cpInicio.SetFocus
    Exit Sub
  End If
    
  If Not IsDate(cpFim.Text) Then
    MsgBox "A data final não foi informada ou não é válida.", vbInformation + vbOKOnly, "Aviso"
    cpFim.SetFocus
    Exit Sub
  End If
    
  If CDate(cpInicio.Text) > CDate(cpFim.Text) Then
    MsgBox "A data final não pode ser menor que a data inicial.", vbInformation + vbOKOnly, "Aviso"
    cpFim.SetFocus
    Exit Sub
  End If
    
  'cmdPrint.Enabled = False
  MontaRelatorio
  sOrdem = "{relextrato.vencimento};crAscendingOrder;relextrato|"
  RelatoriosRPT.Carregar sOrdem, Parametros.dados, "", "Boletos quitados entre " & cpInicio.Text & " e " & cpFim.Text & " - " & cpNome.Text, sFormataCaminho(App.Path) & "boletospagos.rpt", , , , "relextrato", , , , "detalhes"
  'cmdPrint.Enabled = True
  
End Sub

Private Sub cpCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpInicio.SetFocus
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
          cpInicio.SetFocus
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
        cpInicio.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpFim_GotFocus()
  If IsDate(cpInicio.Text) Then
    cpFim.Text = UltimoDia(Month(cpInicio.Text) & "/" & Year(cpInicio.Text))
  End If
End Sub

Private Sub cpFim_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdPrint.Enabled Then
      cmdPrint.SetFocus
    End If
  End If
End Sub

Private Sub cpInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFim.SetFocus
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
  
  dbLocal.Execute "delete from relextrato;"
  Set rsrel = dbLocal.OpenRecordset("relextrato", dbOpenTable)
  Set rs = db.OpenRecordset("select * from associados where condominio = " _
    & cpCodigo.Text & " order by codigo;", dbOpenDynaset)
  
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
          Set rsBol = db.OpenRecordset("select * from boletos where cancelado = 'N' and left(tran,2) = 'AV' and cdsc = " & !Codigo & " and PAGO = 'S' and dtpgto between #" _
                & Format$(cpInicio.Text, "MM/dd/yyyy") & "# and #" & Format$(cpFim.Text, "MM/dd/yyyy") & "# order by vcto;", dbOpenDynaset)
          If rsBol.RecordCount > 0 Then
            rsBol.MoveFirst
            Do While Not rsBol.EOF
              rsrel.AddNew
              rsrel!id_associado = !Codigo
              rsrel!id_condominio = cpCodigo.Text
              rsrel!COTA = rsBol!tran
              rsrel!Data = rsBol!dtpgto
              rsrel!valor = rsBol!valr
              rsrel!juros = rsBol!vlpago - rsBol!valr
              rsrel!Vencimento = rsBol!vcto
              rsrel!nosso_numero = rsBol!DIGITAVAL
              rsrel!tarifas = PegaTaxa(rsBol!nosso)
              rsrel!somar = 1
              rsrel.Update
              rsBol.MoveNext
            Loop
          End If
          Set rsBol = Nothing
          Set rsBol = db.OpenRecordset("select * from boletos where cancelado = 'N' and left(tran,2) = 'GE' and cdsc = " & !Codigo & " and PAGO = 'S' and dtpgto between #" _
                & Format$(cpInicio.Text, "MM/dd/yyyy") & "# and #" & Format$(cpFim.Text, "MM/dd/yyyy") & "# order by vcto;", dbOpenDynaset)
          If rsBol.RecordCount > 0 Then
            rsBol.MoveFirst
            Do While Not rsBol.EOF
              rsrel.AddNew
              rsrel!id_associado = !Codigo
              rsrel!id_condominio = cpCodigo.Text
              rsrel!COTA = rsBol!tran
              rsrel!Data = rsBol!dtpgto
              rsrel!valor = rsBol!valr
              rsrel!juros = rsBol!vlpago - rsBol!valr
              rsrel!Vencimento = rsBol!vcto
              rsrel!nosso_numero = rsBol!DIGITAVAL
              rsrel!tarifas = PegaTaxa(rsBol!nosso)
              rsrel!somar = 1
              rsrel.Update
              rsBol.MoveNext
            Loop
          End If
          Set rsBol = Nothing
          Set rsBol = db.OpenRecordset("select * from boletos where cancelado = 'N' and left(tran,2) <> 'AV' and cdsc = " & !Codigo & " and PAGO = 'S' and dtpgto between #" _
                & Format$(cpInicio.Text, "MM/dd/yyyy") & "# and #" & Format$(cpFim.Text, "MM/dd/yyyy") & "# order by vcto;", dbOpenDynaset)
          If rsBol.RecordCount > 0 Then
            rsBol.MoveFirst
            Do While Not rsBol.EOF
              rsrel.AddNew
              rsrel!id_associado = !Codigo
              rsrel!id_condominio = cpCodigo.Text
              rsrel!COTA = rsBol!tran
              rsrel!Data = rsBol!dtpgto
              rsrel!valor = rsBol!valr
              rsrel!juros = rsBol!vlpago - rsBol!valr
              rsrel!Vencimento = rsBol!vcto
              rsrel!nosso_numero = rsBol!DIGITAVAL
              rsrel!tarifas = PegaTaxa(rsBol!nosso)
              If !acumulado = "S" Then
                If (rsBol.AbsolutePosition + 1) < rsBol.RecordCount And rsBol.RecordCount > 1 Then
                  rsrel!somar = 0
                Else
                  rsrel!somar = 1
                End If
              Else
                rsrel!somar = 1
              End If
              rsrel.Update
              rsBol.MoveNext
            Loop
          End If
          Set rsBol = Nothing
        .MoveNext
      Loop
    Else
      MsgBox "Este condomínio não possui inquilinos.", vbInformation + vbOKOnly, "Aviso"
    End If
  End With
  Set rs = Nothing
End Sub

Private Function PegaTaxa(sBol As String) As Double
  Dim rs As Recordset
  Dim dRet As Double
  
  dRet = 0
  
  Set rs = db.OpenRecordset("select * from baixados where titulo = '" & sBol & "' order by data;", dbOpenDynaset)
  If rs.RecordCount > 0 Then
    rs.MoveLast
    dRet = rs!tarifas
  End If
  rs.Close
  Set rs = Nothing
  PegaTaxa = dRet
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
