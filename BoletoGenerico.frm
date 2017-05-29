VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{CCF682E3-14F1-11D7-80E2-10E08C102140}#1.0#0"; "zprogbar.ocx"
Begin VB.Form BoletoGenerico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de boletos genéricos"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   Icon            =   "BoletoGenerico.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "..."
      Height          =   315
      Left            =   2115
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Fechar"
      Height          =   795
      Left            =   9060
      Picture         =   "BoletoGenerico.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   60
      Width           =   1260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9128
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Prestações"
      TabPicture(0)   =   "BoletoGenerico.frx":0316
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cpMesInicio"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cpDiaVencimento"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cpPrestacoes"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cpValor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cpHistorico"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cpFracao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdPrestacoes"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ListaPrestacoes"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cpIniciarEm"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cpDe"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Boletos"
      TabPicture(1)   =   "BoletoGenerico.frx":0332
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cpMens6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cpMens5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cpMens4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cpMens3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cpMens2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cpMens1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cpMensagem(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cpMensagem(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cpMensagem(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cpMensagem(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdPrint"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Impressão"
      TabPicture(2)   =   "BoletoGenerico.frx":034E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cpSelecao"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdEtiquetas"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cpHistoricoPrint"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin rdActiveText.ActiveText cpDe 
         Height          =   315
         Left            =   -68400
         TabIndex        =   7
         Top             =   1260
         Width           =   555
         _ExtentX        =   979
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
         MaxLength       =   3
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpIniciarEm 
         Height          =   315
         Left            =   -69420
         TabIndex        =   6
         Top             =   1260
         Width           =   615
         _ExtentX        =   1085
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
         MaxLength       =   3
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.ComboBox cpHistoricoPrint 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   720
         Width           =   7155
      End
      Begin VB.CommandButton cmdEtiquetas 
         Caption         =   "&Imprimir"
         Height          =   795
         Left            =   7080
         Picture         =   "BoletoGenerico.frx":036A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1920
         Width           =   1260
      End
      Begin VB.ComboBox cpSelecao 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1200
         Width           =   1695
      End
      Begin ComctlLib.TreeView ListaPrestacoes 
         Height          =   2535
         Left            =   -74820
         TabIndex        =   11
         Top             =   2460
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4471
         _Version        =   327682
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Gerar"
         Height          =   795
         Left            =   -67920
         Picture         =   "BoletoGenerico.frx":0674
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   540
         Width           =   1260
      End
      Begin VB.TextBox cpMensagem 
         Height          =   285
         Index           =   2
         Left            =   -74880
         MaxLength       =   250
         TabIndex        =   20
         Top             =   4005
         Width           =   5025
      End
      Begin VB.TextBox cpMensagem 
         Height          =   285
         Index           =   1
         Left            =   -74880
         MaxLength       =   250
         TabIndex        =   19
         Text            =   "APÓS VENCIMENTO, SÓ RECEBER COM MULTA DE "
         Top             =   3690
         Width           =   5025
      End
      Begin VB.TextBox cpMensagem 
         Height          =   285
         Index           =   0
         Left            =   -74880
         MaxLength       =   250
         TabIndex        =   18
         Text            =   "SR CAIXA, NÃO RECEBER APÓS 15 DIAS DE VENCIDO"
         Top             =   3405
         Width           =   5025
      End
      Begin VB.TextBox cpMensagem 
         Height          =   285
         Index           =   3
         Left            =   -74880
         MaxLength       =   250
         TabIndex        =   21
         Top             =   4290
         Width           =   5025
      End
      Begin VB.CommandButton cmdPrestacoes 
         Caption         =   "Gerar"
         Height          =   315
         Left            =   -67665
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cpFracao 
         Height          =   315
         ItemData        =   "BoletoGenerico.frx":097E
         Left            =   -73680
         List            =   "BoletoGenerico.frx":0988
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   4575
      End
      Begin rdActiveText.ActiveText cpHistorico 
         Height          =   315
         Left            =   -73680
         TabIndex        =   3
         Top             =   900
         Width           =   7215
         _ExtentX        =   12726
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
      End
      Begin rdActiveText.ActiveText cpValor 
         Height          =   315
         Left            =   -73680
         TabIndex        =   4
         Top             =   1260
         Width           =   1515
         _ExtentX        =   2672
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
         FloatFormat     =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpPrestacoes 
         Height          =   315
         Left            =   -70965
         TabIndex        =   5
         Top             =   1260
         Width           =   675
         _ExtentX        =   1191
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
         MaxLength       =   3
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpDiaVencimento 
         Height          =   315
         Left            =   -73365
         TabIndex        =   8
         Top             =   1620
         Width           =   855
         _ExtentX        =   1508
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
         MaxLength       =   2
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpMesInicio 
         Height          =   315
         Left            =   -71280
         TabIndex        =   9
         Top             =   1620
         Width           =   1095
         _ExtentX        =   1931
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
      Begin rdActiveText.ActiveText cpMens1 
         Height          =   315
         Left            =   -74880
         TabIndex        =   12
         Top             =   750
         Width           =   5235
         _ExtentX        =   9234
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
      Begin rdActiveText.ActiveText cpMens2 
         Height          =   315
         Left            =   -74880
         TabIndex        =   13
         Top             =   1095
         Width           =   5235
         _ExtentX        =   9234
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
         Left            =   -74880
         TabIndex        =   14
         Top             =   1440
         Width           =   5235
         _ExtentX        =   9234
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
         Left            =   -74880
         TabIndex        =   15
         Top             =   1785
         Width           =   5235
         _ExtentX        =   9234
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   2130
         Width           =   5235
         _ExtentX        =   9234
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
         Left            =   -74880
         TabIndex        =   17
         Top             =   2475
         Width           =   5235
         _ExtentX        =   9234
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "de"
         Height          =   195
         Left            =   -68700
         TabIndex        =   44
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Iniciar em"
         Height          =   195
         Left            =   -70140
         TabIndex        =   43
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tecle F5 para reimpressão"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mensagen promocional"
         Height          =   195
         Left            =   -74880
         TabIndex        =   33
         Top             =   540
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Texto de responsabilidade do caixa"
         Height          =   195
         Left            =   -74880
         TabIndex        =   32
         Top             =   3180
         Width           =   2505
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Iniciar no mês"
         Height          =   195
         Left            =   -72360
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Dia do vencimento"
         Height          =   195
         Left            =   -74805
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "N. Prestações"
         Height          =   195
         Left            =   -72045
         TabIndex        =   29
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         Height          =   195
         Left            =   -74760
         TabIndex        =   27
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Distribuição"
         Height          =   195
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   825
      End
   End
   Begin ZealProgressBar.ProgressBar Barra 
      Height          =   255
      Left            =   420
      TabIndex        =   24
      Top             =   6120
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
   Begin rdActiveText.ActiveText cpNome 
      Height          =   315
      Left            =   2595
      TabIndex        =   35
      Top             =   120
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
      Left            =   1320
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Condomínio"
      Height          =   195
      Left            =   180
      TabIndex        =   36
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Porcento 
      AutoSize        =   -1  'True
      Caption         =   "Progresso"
      Height          =   195
      Left            =   435
      TabIndex        =   23
      Top             =   5895
      Width           =   705
   End
End
Attribute VB_Name = "BoletoGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nCod As Long
Dim sTemp As String
Dim qDig As Integer
Dim i As Integer
Dim h As Integer
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

Private Sub cmdPrestacoes_Click()
  
  If Not IsDate("25/" & cpMesInicio.Text) Then
    MsgBox "Por favor informe o mês de início.", vbCritical & vbOKOnly, "Aviso"
    cpMesInicio.SetFocus
    GoTo Fim
  End If
  
  If Val(cpDiaVencimento.Text) < 1 Or Val(cpDiaVencimento.Text) > 31 Then
    MsgBox "O dia do vencimento informado não é válido.", vbCritical & vbOKOnly, "Aviso"
    cpDiaVencimento.SetFocus
    GoTo Fim
  End If
  
  If CDate(cpDiaVencimento.Text & "/" & cpMesInicio.Text) < Date Then
    MsgBox "O a primeira prestação tem vencimento menor que a data atual.", vbCritical & vbOKOnly, "Aviso"
    cpDiaVencimento.SetFocus
    GoTo Fim
  End If
  
  If Val(cpCodigo.Text) <= 0 Then
    MsgBox "Por favor informe o condomínio.", vbCritical & vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    GoTo Fim
  End If

  If Trim(cpHistorico.Text) = "" Then
    MsgBox "Preencha o histórico da despesa.", vbCritical + vbOKOnly, "Aviso"
    cpHistorico.SetFocus
    GoTo Fim
  End If
  
  If Val(cpValor.Text) <= 0 Then
    MsgBox "O valor informado deve ser maior que zero.", vbCritical + vbOKOnly, "Aviso"
    cpValor.SetFocus
    GoTo Fim
  End If
  
  If Val(cpPrestacoes.Text) < 1 Then
    MsgBox "O número de prestações deve ser maior que 1.", vbCritical + vbOKOnly, "Aviso"
    cpPrestacoes.SetFocus
    GoTo Fim
  End If
  
  If Val(cpIniciarEm.Text) < 1 Or Val(cpIniciarEm.Text) > Val(cpDe.Text) Then
    MsgBox "O número para primeira prestação deve ser maior que 1 e menor que " & cpDe.Text & ".", vbCritical + vbOKOnly, "Aviso"
    cpIniciarEm.SetFocus
    GoTo Fim
  End If
  
  If Val(cpDe.Text) < (Val(cpIniciarEm.Text) + Val(cpPrestacoes.Text) - 1) Then
    MsgBox "O número da última prestação (campo de) deve ser " & (Val(cpIniciarEm.Text) + Val(cpPrestacoes.Text) - 1) & " ou maior.", vbCritical + vbOKOnly, "Aviso"
    cpIniciarEm.SetFocus
    GoTo Fim
  End If
  
  If JaExiste("boletos", "historico", cpHistorico.Text) Then
    MsgBox "Este histórico já existe.", vbCritical + vbOKOnly, "Aviso"
    cpHistorico.SetFocus
    GoTo Fim
  End If
    
  
  Dim rsPrestacoes As Recordset
  Dim rsInquilinos As Recordset
  Dim SomaQuotaFracao As Double
  Dim PegaFracao As Double
  Dim ValorInquilino As Double
  Dim ValorGeral As Double
  Dim ValorCotaFracao As Double
  Dim tpTaxa() As crTaxa
  Dim nPrest As Integer
  Dim dVencimento As Date
  Dim dInicio As Date
  Dim nNode As node
  Dim IniciarEm As Integer
  
  Me.MousePointer = 11
  Me.Refresh
  Me.Enabled = False
  
  dbLocal.Execute "delete from prestacoes;"
  Set rsInquilinos = db.OpenRecordset("Select * from associados where condominio = " & cpCodigo.Text & " order by codigo;", dbOpenDynaset)
  Set rsPrestacoes = dbLocal.OpenRecordset("prestacoes", dbOpenTable)
  
  ListaPrestacoes.Nodes.Clear
  DoEvents
  
  SomaQuotaFracao = 0
  i = 0
  With rsInquilinos
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If cpFracao.ListIndex = 0 Then
          PegaFracao = RetornaFracao(!Codigo)
        Else
          PegaFracao = 1
        End If
        If PegaFracao > 0 Then
          ReDim Preserve tpTaxa(i) As crTaxa
          tpTaxa(i).iFracao = PegaFracao
          tpTaxa(i).iCodigo = !Codigo
          tpTaxa(i).iValor = 0
          tpTaxa(i).iInquilino = NomeCompleto(!Codigo)
          tpTaxa(i).iSindico = !Sindico
          tpTaxa(i).iPerc = 0
          tpTaxa(i).iValorProp = 0
          SomaQuotaFracao = SomaQuotaFracao + tpTaxa(i).iFracao
          i = i + 1
        End If
        Barra.Value = Int(.PercentPosition / 2)
        .MoveNext
      Loop
    End If
  End With
  
  IniciarEm = Val(cpIniciarEm.Text) - 1
  nPrest = Val(cpPrestacoes.Text)
  ValorCotaFracao = CDbl(cpValor.Text) / SomaQuotaFracao
  
  For i = 0 To UBound(tpTaxa)
    If cpFracao.ListIndex = 0 Then
      tpTaxa(i).iValor = ValorCotaFracao * tpTaxa(i).iFracao
    Else
      tpTaxa(i).iValor = ValorCotaFracao
    End If
    tpTaxa(i).iValorProp = tpTaxa(i).iValor / nPrest
  Next i
  
  
  If IsDate(cpDiaVencimento.Text & "/" & cpMesInicio.Text) Then
    dInicio = CDate(cpDiaVencimento.Text & "/" & cpMesInicio.Text)
  Else
    dInicio = UltimoDia(cpMesInicio.Text)
  End If
  
  ValorGeral = 0
  For i = 0 To UBound(tpTaxa)
    Set nNode = ListaPrestacoes.Nodes.Add(, , "P" & tpTaxa(i).iCodigo)
    nNode.Text = tpTaxa(i).iInquilino
    ValorInquilino = 0
    For h = 1 To nPrest
      dVencimento = DateAdd("m", h - 1, dInicio)
      Set nNode = ListaPrestacoes.Nodes.Add("P" & tpTaxa(i).iCodigo, tvwChild, "F" & tpTaxa(i).iCodigo & "|" & h)
      nNode.Text = PadLeft(IniciarEm + h, "0", 3) & "/" & cpDe.Text & " | " & Format$(dVencimento, "dd/MM/yyyy") & " | " & Format$(tpTaxa(i).iValorProp, "#,##0.00")
      ValorInquilino = ValorInquilino + tpTaxa(i).iValorProp
    Next h
    Set nNode = ListaPrestacoes.Nodes.Add("P" & tpTaxa(i).iCodigo, tvwChild, "T" & tpTaxa(i).iCodigo)
    nNode.Text = "Total:  " & Format$(ValorInquilino, "#,##0.00")
    ValorGeral = ValorGeral + ValorInquilino
    Barra.Value = Int(50 + ((i / UBound(tpTaxa) * 100) / 2))
  Next i
  Set nNode = ListaPrestacoes.Nodes.Add(, , "Tgeral")
  nNode.Text = "Total:  " & Format$(ValorGeral, "#,##0.00")
  nNode.EnsureVisible
  
  
Fim:
  Me.MousePointer = 0
  Me.Enabled = True
  Me.Refresh
  DoEvents
  Exit Sub
  
End Sub

Private Sub cmdPrint_Click()
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
  Dim sDoc As String
  Dim selCondominio As Recordset
  Dim rsI As Recordset
  Dim NossoNumero As String
  Dim sLivre As String
  Dim nNode As node
  Dim fNode As node
  Dim Inquilino As Integer
  Dim nPrest As Integer
  
  If ListaPrestacoes.Nodes.Count = 0 Then
    MsgBox "Gere as prestações primeiro.", vbCritical + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  cmdCancelar.Enabled = False
  cmdPrint.Enabled = False
  cmdEtiquetas.Enabled = False
  
  nCod = cpCodigo.Text
  
  Set selCondominio = db.OpenRecordset("Select * From CONDOMINIO where codigo = " & nCod & " order by nome;", dbOpenDynaset)
  
  DBEngine.BeginTrans
  With selCondominio
    If .RecordCount > 0 Then
      .MoveFirst
      i = 0
      For Each nNode In ListaPrestacoes.Nodes
        If InStr(nNode.Key, "F") > 0 Then
          Inquilino = CInt(GetPiece(Mid$(nNode.Key, InStr(nNode.Key, "F") + 1), "|", 1))
          dtVenc = CDate(GetPiece(nNode.Text, "|", 2))
          nVal = CDbl(GetPiece(nNode.Text, "|", 3))
          sDoc = GetPiece(nNode.Text, "|", 1)
          nPrest = Val(Left(sDoc, InStr(sDoc, "/") - 1))
          If nVal > 0 Then
            
            Set rsI = db.OpenRecordset("select * from associados where codigo = " & Inquilino & ";", dbOpenDynaset)
            rsI.MoveFirst
            
            fValor = Format(nVal, "#0.00")
            fValor = Left(fValor, InStr(fValor, ",") - 1) & Right(fValor, 2)
            fValor = PadLeft(fValor, "0", 10)
            nFator = CStr(dtVenc - CDate("07/10/1997"))
            
            If selCondominio!tipoboleto = 1 Then
              
              NossoNumero = Trim(selCondominio!carteira) & PadLeft(nPrest, "0", 4)
              NossoNumero = NossoNumero & PadLeft(selCondominio!Codigo, "0", 4)
              NossoNumero = NossoNumero & PadLeft(Inquilino, "0", 7)
              
              sTemp = selCondominio!conta & ""
              qDig = DigitosCedente(1)
              If Len(sTemp) > qDig Then
                sTemp = Right(sTemp, qDig)
              ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
                sTemp = PadLeft(sTemp, "0", qDig)
              End If

              sLivre = sTemp & DigitoNosso(sTemp)
              sLivre = sLivre & Mid$(NossoNumero, 3, 3) & Left(NossoNumero, 1)
              sLivre = sLivre & Mid$(NossoNumero, 6, 3)
              sLivre = sLivre & Mid$(NossoNumero, 2, 1)
              sLivre = sLivre & Mid$(NossoNumero, 9)
                  
              sLivre = sLivre & DigitoNosso(sLivre)
            ElseIf selCondominio!tipoboleto = 2 Then
            
              sTemp = selCondominio!conta & ""
              qDig = DigitosCedente(2)
              If Len(sTemp) > qDig Then
                sTemp = Right(sTemp, qDig)
              ElseIf Len(sTemp) > 0 And Len(sTemp) < qDig Then
                sTemp = PadLeft(sTemp, "0", qDig)
              End If
              
              NossoNumero = Trim(selCondominio!carteira) & PadLeft(Inquilino, "0", 5) & PadLeft(nPrest, "0", 3)
              
              sLivre = NossoNumero & selCondominio!agcedente & selCondominio!Operacao & sTemp
            
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
            tbBoletos!Historico = cpHistorico.Text
            tbBoletos!corrigido = nVal
            If rsI!boleto = 1 Then
              tbBoletos!COTA = PadLeft(rsI!Codigo, "0", 4)
              tbBoletos!cdsc = rsI!Codigo
              tbBoletos!Nome = NomeCompleto(rsI!Codigo)
              tbBoletos!cpf = rsI!pcpf
              tbBoletos!Ende = rsI!PEndereco
              tbBoletos!Bair = rsI!PBairro
              tbBoletos!Cida = rsI!PCidade
              tbBoletos!Esta = rsI!PEstado
              tbBoletos!cep = rsI!PCep
            Else
              tbBoletos!COTA = PadLeft(rsI!Codigo, "0", 4)
              tbBoletos!cdsc = rsI!Codigo
              tbBoletos!Nome = NomeCompleto(rsI!Codigo)
              tbBoletos!cpf = rsI!cpf
              tbBoletos!Ende = selCondominio!endereco
              tbBoletos!Bair = selCondominio!bairro
              tbBoletos!Cida = selCondominio!Cidade
              tbBoletos!Esta = selCondominio!estado
              tbBoletos!cep = selCondominio!cep
            End If
            tbBoletos!tran = "GE" & PadLeft(Inquilino, "0", 5) & PadLeft(nPrest, "0", 3)
            tbBoletos!DIGITAVAL = NossoNumero & "." & DigitoNosso(NossoNumero)
            tbBoletos!agcedente = selCondominio!agcedente
            tbBoletos!carteira = selCondominio!carteira
            If selCondominio!tipoboleto = 1 Then
              tbBoletos!Mensagem = selCondominio!agcedente & "/" & selCondominio!conta & "-" & DigitoNosso(selCondominio!conta)
            Else
              tbBoletos!Mensagem = Trim(selCondominio!agcedente) & "." & Trim(selCondominio!Operacao) & "." & Trim(selCondominio!conta) & "." & DigitoNosso(Trim(selCondominio!agcedente) & Trim(selCondominio!Operacao) & Trim(selCondominio!conta))
            End If
            tbBoletos!pago = "N"
            tbBoletos!CODI = Campo1 & " " & Campo2 & " " & Campo3 & " " & Campo4 & " " & Campo5
            tbBoletos!CDBARRA = sBarra
            tbBoletos!texto = cpMens1.Text & vbCrLf & cpMens2.Text & vbCrLf & cpMens3.Text & vbCrLf & cpMens4.Text & vbCrLf & cpMens5.Text & vbCrLf & cpMens6.Text & vbCrLf & cpHistorico.Text
            tbBoletos!INST1 = cpMensagem(0).Text
            tbBoletos!INST2 = cpMensagem(1).Text
            tbBoletos!INST3 = cpMensagem(2).Text
            tbBoletos!INST4 = cpMensagem(3).Text
            tbBoletos!Banco = "CAIXA" 'selCondominio!Banco
            tbBoletos!CdBanco = selCondominio!CdAgencia
            tbBoletos!CANCELADO = "N"
            tbBoletos!cond = selCondominio!Codigo
            tbBoletos!acumulado = "N"
            tbBoletos!bole = NossoNumero
            tbBoletos!nosso = NossoNumero
            tbBoletos!desconto = 0
            tbBoletos!idStatus = 1
            tbBoletos.Update
            tbBoletos.Bookmark = tbBoletos.LastModified
            db.Execute "insert into BOLETODETALHE (CONDOMINIO, MES, DESCRICAO, GERAIS, SINDICO, " _
              & "PROPRIETARIO, FRACAO, VALOR, BOX, ID_ASSOCIADO, ID_CONDOMINIO, VALOR_PROPRIETARIO, ID_BOLETO) " _
              & "VALUES ('" & cpCodigo.Text & "', '" & Format$(dtVenc, "MM/yyyy") & "', 'PRESTAÇÃO " & sDoc & "', 0, 0, '" _
              & NomeCompleto(Inquilino) & "', 0, " & Replace(Format$(nVal, "#0.00"), ",", ".") _
              & ", '" & rsI!Tipo & " " & rsI!Apartamento & "', " & Inquilino & ", " & selCondominio!Codigo _
              & ", 0, " & tbBoletos!id & ");"
            
          End If
        End If
        Barra.Value = Int(i / ListaPrestacoes.Nodes.Count * 100)
        DoEvents
        i = i + 1
      Next
    End If
  End With
  DBEngine.CommitTrans
  Dim rsSel As Recordset
  
  cpHistoricoPrint.AddItem (cpHistorico.Text)
  cpHistoricoPrint.ListIndex = 0
  
  Set rsSel = db.OpenRecordset("select distinct vcto from boletos where left(tran,2) = 'GE' and cond = " & cpCodigo.Text & " order by vcto;", dbOpenDynaset)
  With rsSel
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cpSelecao.AddItem (Format(!vcto, "dd/MM/yyyy"))
        .MoveNext
      Loop
    End If
  End With
  Set rsSel = Nothing

Fim:
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
  
  If cpHistoricoPrint.ListIndex < 0 Then
    MsgBox "Selecione um histórico.", vbInformation + vbOKOnly, "Aviso"
    cpHistoricoPrint.SetFocus
    Exit Sub
  End If
  
  If cpSelecao.ListIndex < 0 Then
    MsgBox "Selecione um vencimento.", vbInformation + vbOKOnly, "Aviso"
    cpSelecao.SetFocus
    Exit Sub
  End If
    
  
'  cmdCancelar.Enabled = False
'  cmdPrint.Enabled = False
'  cmdEtiquetas.Enabled = False
  
  nCod = cpCodigo.Text
  
  sSql = "Select * from boletos where left(tran,2) = 'GE' and cond = " & nCod & " and vcto = #" & Format(cpSelecao.Text, "MM/dd/yyyy") & "# order by bole;"
  
  Dim PerSind As Double
  
  PerSind = PersentualSindico(CodigoSindico(nCod))
  
  RelatoriosRPT.mnuSupExport.Visible = False
  RelatoriosRPT.mnuExportBoleto.Visible = True
  RelatoriosRPT.mnuOrdenar.Visible = True
  RelatoriosRPT.mnuEviarEmail.Visible = True
  RelatoriosRPT.Carregar "{boletos.bole};crAscendingOrder;boletos|", Parametros.dados, "{boletos.historico}= '" & cpHistoricoPrint.Text & "' and Left({boletos.tran},2) = 'GE' and {boletos.cond} = " _
      & nCod & " AND {BOLETOS.VCTO}=date(" & Year(cpSelecao.Text) & "," & Month(cpSelecao.Text) & "," & Day(cpSelecao.Text) _
      & ")", "Boletos", sFormataCaminho(App.Path) & "generico.rpt", , sSql, "persind|" & PerSind, , nCod, "0000"
  
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
          cpFracao.SetFocus
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
        cpFracao.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpDe_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDiaVencimento.SetFocus
  End If
End Sub

Private Sub cpDiaVencimento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMesInicio.SetFocus
  End If
End Sub

Private Sub cpFracao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpHistorico.SetFocus
  End If
End Sub

Private Sub cpHistorico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  End If
End Sub

Private Sub cpHistoricoPrint_LostFocus()
  Dim rsSel As Recordset
  cpSelecao.Clear
  Set rsSel = db.OpenRecordset("select distinct vcto from boletos where historico = '" & cpHistoricoPrint.Text & "' and left(tran,2) = 'GE' and cond = " & cpCodigo.Text & " order by vcto;", dbOpenDynaset)
  With rsSel
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cpSelecao.AddItem (Format(!vcto, "dd/MM/yyyy"))
        .MoveNext
      Loop
    End If
  End With
  Set rsSel = Nothing
End Sub

Private Sub cpIniciarEm_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDe.SetFocus
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

Private Sub cpMesInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrestacoes.SetFocus
  End If
End Sub

Private Sub cpMesInicio_LostFocus()
  cpMesInicio.Text = FormataMesAno(cpMesInicio.Text)
End Sub

Private Sub cpPrestacoes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpIniciarEm.SetFocus
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpPrestacoes.SetFocus
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 116 Then
    ReImpressao
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
  SSTab1.Tab = 0
End Sub

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

Private Sub ReImpressao()
  Dim rsSel As Recordset
  
  If Val(cpCodigo.Text) = 0 Then
    MsgBox "Selecione um condomínio.", vbCritical + vbOKOnly, "Aviso"
    cpCodigo.SetFocus
    Exit Sub
  End If
  
  cpHistoricoPrint.Clear
  Set rsSel = db.OpenRecordset("select distinct historico from boletos where left(tran,2) = 'GE' and cond = " & cpCodigo.Text & ";", dbOpenDynaset)
  With rsSel
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cpHistoricoPrint.AddItem (!Historico & "")
        .MoveNext
      Loop
      cpHistoricoPrint.ListIndex = 0
    End If
  End With
  Set rsSel = Nothing
  
  cpSelecao.Clear
  Set rsSel = db.OpenRecordset("select distinct vcto from boletos where historico = '" & cpHistoricoPrint.Text & "' and left(tran,2) = 'GE' and cond = " & cpCodigo.Text & " order by vcto;", dbOpenDynaset)
  With rsSel
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        cpSelecao.AddItem (Format(!vcto, "dd/MM/yyyy"))
        .MoveNext
      Loop
    End If
  End With
  Set rsSel = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbCondominio = Nothing
  Set tbBoletos = Nothing
End Sub
