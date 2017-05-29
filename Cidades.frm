VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.3#0"; "ACTIVETEXT.OCX"
Begin VB.Form Cidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de cidades"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "Cidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin rdActiveText.ActiveText cpCep 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   465
      Width           =   1335
      _ExtentX        =   2355
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
      MaxLength       =   9
      TextMask        =   6
      RawText         =   6
      Mask            =   "#####-###"
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpEstado 
      Height          =   315
      Left            =   735
      TabIndex        =   1
      Top             =   465
      Width           =   660
      _ExtentX        =   1164
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
      MaxLength       =   2
      TextCase        =   1
      RawText         =   0
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin rdActiveText.ActiveText cpCidade 
      Height          =   315
      Left            =   735
      TabIndex        =   0
      Top             =   120
      Width           =   5100
      _ExtentX        =   8996
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
      MaxLength       =   50
      TextCase        =   1
      RawText         =   0
      FontName        =   "Arial"
      FontSize        =   9,75
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cep"
      Height          =   195
      Left            =   1665
      TabIndex        =   5
      Top             =   540
      Width           =   285
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   195
      Width           =   495
   End
End
Attribute VB_Name = "Cidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cpCep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    With tbCidades
      .AddNew
      !Nome = cpCidade.Text
      !Estado = cpEstado.Text
      !Cep = cpCep.Text
      .Update
    End With
    Unload Me
  End If
End Sub

Private Sub cpCidade_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEstado.SetFocus
  End If
End Sub

Private Sub cpEstado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCep.SetFocus
  End If
End Sub

Private Sub Form_Load()
 Refresh
 KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RetCidade(0) = cpCidade.Text
  RetCidade(1) = cpEstado.Text
  RetCidade(2) = cpCep.Text
End Sub
