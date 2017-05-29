VERSION 5.00
Begin VB.Form frmMensagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensagem"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   ControlBox      =   0   'False
   Icon            =   "frmMensagem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   10200
      TabIndex        =   1
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9060
      TabIndex        =   0
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label lbMensagem 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbMensagem"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   11175
   End
End
Attribute VB_Name = "frmMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  Dim sLinhas() As String
  Dim linha As Single
  Dim i As Integer
  Dim daParametros As Recordset
  
  Set daParametros = db.OpenRecordset("parametros", dbOpenTable)
  daParametros.MoveFirst
  Printer.ScaleMode = vbCentimeters
  Printer.PaperSize = vbPRPSLegal
  PrnDireita 0.6, 3, daParametros!Empresa, "Arial", 14, True
  PrnDireita 1.2, 3, daParametros!Endereco, "Arial", 9, False
  PrnDireita 1.2, 10, daParametros!Bairro, "Arial", 9, False
  PrnDireita 1.2, 15, daParametros!Cidade, "Arial", 9, False
  PrnDireita 1.6, 3, daParametros!Estado, "Arial", 9, False
  PrnDireita 1.6, 4, daParametros!Cep, "Arial", 9, False
  PrnDireita 1.6, 6.5, daParametros!CNPJ, "Arial", 9, False
  PrnDireita 1.6, 10.5, daParametros!Inscricao, "Arial", 9, False
  PrnDireita 1.6, 15, daParametros!Telefone, "Arial", 9, False
  PrnLinha 2, 1, 2.02, 20
  sLinhas = Split(lbMensagem.Caption, vbCrLf)
  linha = 2.3
  PrnDireita linha, 1, sLinhas(i), "Arial", 10, False
  linha = linha + 0.4
  For i = 1 To UBound(sLinhas)
    If sLinhas(i) <> "" Then
      PrnDireita linha, 1, Left(sLinhas(i), InStr(sLinhas(i), " ") - 1), "Arial", 10, False
      PrnDireita linha, 5, Mid(sLinhas(i), InStr(sLinhas(i), " ") + 1, 40), lbMensagem.FontName, lbMensagem.FontSize, False
      PrnEsquerda linha, 20, Mid(sLinhas(i), InStrRev(sLinhas(i), " ")), lbMensagem.FontName, lbMensagem.FontSize, False
      linha = linha + 0.4
    Else
      PrnDireita linha, 1, sLinhas(i), "Arial", 10, False
      linha = linha + 0.4
    End If
    If linha > 27 Then
      Printer.NewPage
      PrnDireita 0.6, 3, daParametros!Empresa, "Arial", 14, True
      PrnDireita 1.2, 3, daParametros!Endereco, "Arial", 9, False
      PrnDireita 1.2, 10, daParametros!Bairro, "Arial", 9, False
      PrnDireita 1.2, 15, daParametros!Cidade, "Arial", 9, False
      PrnDireita 1.6, 3, daParametros!Estado, "Arial", 9, False
      PrnDireita 1.6, 4, daParametros!Cep, "Arial", 9, False
      PrnDireita 1.6, 6.5, daParametros!CNPJ, "Arial", 9, False
      PrnDireita 1.6, 10.5, daParametros!Inscricao, "Arial", 9, False
      PrnDireita 1.6, 15, daParametros!Telefone, "Arial", 9, False
      PrnLinha 2, 1, 2.02, 20
      linha = 2.3
    End If
  Next i
  Printer.EndDoc
  Set daParametros = Nothing
End Sub
