VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form RelCpfs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de CPF/CNPJ com problemas"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   Icon            =   "RelCpfs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   2010
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   6060
      TabIndex        =   2
      Top             =   540
      Width           =   1260
   End
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2490
      TabIndex        =   3
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
      Left            =   1230
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
      Left            =   240
      TabIndex        =   4
      Top             =   135
      Width           =   855
   End
End
Attribute VB_Name = "RelCpfs"
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
  Dim sSql As String
  Dim sCpf As String
  Dim sPcpf As String
  Dim rs As Recordset
  Dim rsrel As Recordset
 
  cmdPrint.Enabled = False
  
  dbLocal.Execute "delete from relcpfs;"
  Set rsrel = dbLocal.OpenRecordset("relcpfs", dbOpenTable)
  If Val(cpCodigo.Text) > 0 Then
    sSql = "select * from associados where condominio = " & cpCodigo.Text & " order by codigo;"
  Else
    sSql = "select * from associados order by codigo;"
  End If
  Set rs = db.OpenRecordset(sSql, dbOpenDynaset)
  
  With rs
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        If (!cpf & "" = "") And (!pcpf & "" = "") Then
          rsrel.AddNew
          rsrel!Condominio = !Condominio
          rsrel!Inquilino = !Codigo
          rsrel!cpf = "Sem CPF/CNPJ"
          rsrel!pcpf = "Sem CPF/CNPJ"
          rsrel!nomei = NomeCompleto(!Codigo)
          rsrel!nomec = NomeCondominio(!Condominio)
          rsrel.Update
        Else
          If (!cpf & "" <> "") Then
            sCpf = SoNumeros(!cpf)
            If Len(sCpf) < 12 Then
              If MinhaDll.Valida_Cpf(sCpf) = sdSucesso Then
                sCpf = ""
              End If
            Else
              If MinhaDll.Valida_CGC(sCpf) = sdSucesso Then
                sCpf = ""
              End If
            End If
          End If
          If (!pcpf & "" <> "") Then
            sPcpf = SoNumeros(!pcpf)
            If Len(sPcpf) < 12 Then
              If MinhaDll.Valida_Cpf(sPcpf) = sdSucesso Then
                sPcpf = ""
              End If
            Else
              If MinhaDll.Valida_CGC(sPcpf) = sdSucesso Then
                sPcpf = ""
              End If
            End If
          End If
          If (sCpf <> "") Or (sPcpf <> "") Then
            rsrel.AddNew
            rsrel!Condominio = !Condominio
            rsrel!Inquilino = !Codigo
            If Len(sCpf) < 12 Then
              sCpf = Format$(sCpf, "000\.###\.###\-##")
            Else
              sCpf = Format$(sCpf, "00\.###\.###\/####\-##")
            End If
            If Len(sPcpf) < 12 Then
              sPcpf = Format$(sPcpf, "000\.###\.###\-##")
            Else
              sPcpf = Format$(sPcpf, "00\.###\.###\/####\-##")
            End If
            rsrel!cpf = sCpf
            rsrel!pcpf = sPcpf
            rsrel!nomei = NomeCompleto(!Codigo)
            rsrel!nomec = NomeCondominio(!Condominio)
            rsrel.Update
          End If
        End If
        .MoveNext
      Loop
    End If
  End With
  
  RelatoriosRPT.Carregar "", Parametros.dados, "", "Relatório de CPF/CNPJ com problemas.", sFormataCaminho(App.Path) & "lstcpfs.rpt", , , , "RELCPFS;"
  cmdPrint.Enabled = True

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

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  KeyPreview = True
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
End Sub
