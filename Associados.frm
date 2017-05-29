VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Associados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de inquilinos/proprietários"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "Associados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Abas 
      Height          =   5715
      Left            =   60
      TabIndex        =   31
      Top             =   1080
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   10081
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Inquilino"
      TabPicture(0)   =   "Associados.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbCpfCnpj"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbFracao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label15"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label16"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cpFracao"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label17"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label20"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cpCodCond"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cpNomeCond"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cpCodigo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cpApartamento"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cpNome"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cpEndereco"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cpBairro"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cpCidade"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cpEstado"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cpCep"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cpCpf"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cpIdentidade"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cpFone"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cpFone2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cpObs"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cpTipo"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "ChAcumulado"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cpSindico"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cpValor"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdFracao"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cpBlocos"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cpEmail"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdLocCond"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Propietário"
      TabPicture(1)   =   "Associados.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(1)=   "Label22"
      Tab(1).Control(2)=   "lbCpfProp"
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(4)=   "Label25"
      Tab(1).Control(5)=   "Label26"
      Tab(1).Control(6)=   "Label27"
      Tab(1).Control(7)=   "Label28"
      Tab(1).Control(8)=   "Label29"
      Tab(1).Control(9)=   "Label30"
      Tab(1).Control(10)=   "Label31"
      Tab(1).Control(11)=   "Label19"
      Tab(1).Control(12)=   "cpValCondominio"
      Tab(1).Control(13)=   "cpNomeP"
      Tab(1).Control(14)=   "cpEnderecoP"
      Tab(1).Control(15)=   "cpBairroP"
      Tab(1).Control(16)=   "cpCidadeP"
      Tab(1).Control(17)=   "cpEstadoP"
      Tab(1).Control(18)=   "cpCepP"
      Tab(1).Control(19)=   "cpCpfP"
      Tab(1).Control(20)=   "cpIdentidadeP"
      Tab(1).Control(21)=   "cpFone1P"
      Tab(1).Control(22)=   "cpFone2P"
      Tab(1).Control(23)=   "CpPgto"
      Tab(1).Control(24)=   "chViraFundo"
      Tab(1).Control(25)=   "cpDescVira"
      Tab(1).ControlCount=   26
      Begin rdActiveText.ActiveText cpDescVira 
         Height          =   315
         Left            =   -69960
         TabIndex        =   78
         Top             =   3960
         Width           =   2235
         _ExtentX        =   3942
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
         MaxLength       =   20
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.CheckBox chViraFundo 
         Caption         =   "Este valor vira:"
         Height          =   195
         Left            =   -71340
         TabIndex        =   77
         Top             =   4020
         Width           =   1395
      End
      Begin VB.CommandButton cmdLocCond 
         Caption         =   "..."
         Height          =   315
         Left            =   1860
         TabIndex        =   19
         Top             =   5220
         Width           =   435
      End
      Begin rdActiveText.ActiveText cpEmail 
         Height          =   315
         Left            =   900
         TabIndex        =   14
         Top             =   4140
         Width           =   6375
         _ExtentX        =   11245
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
         MaxLength       =   100
         TextCase        =   2
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.ComboBox cpBlocos 
         Height          =   315
         Left            =   4155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   2220
         Width           =   3135
      End
      Begin VB.CommandButton cmdFracao 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   16
         Top             =   4560
         Width           =   555
      End
      Begin rdActiveText.ActiveText cpValor 
         Height          =   315
         Left            =   5940
         TabIndex        =   67
         Top             =   3240
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
         MaxLength       =   10
         Text            =   "0,00"
         TextMask        =   4
         RawText         =   4
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.ComboBox cpSindico 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   3240
         Width           =   4455
      End
      Begin VB.CheckBox ChAcumulado 
         Caption         =   "Acumular débito(s) anterior(es) no boleto"
         Height          =   195
         Left            =   1080
         TabIndex        =   17
         Top             =   4920
         Width           =   3135
      End
      Begin VB.ComboBox cpTipo 
         Height          =   315
         Left            =   2730
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   495
         Width           =   1590
      End
      Begin VB.CheckBox CpPgto 
         Caption         =   "O boleto será em nome do proprietário"
         Height          =   195
         Left            =   -73965
         TabIndex        =   30
         Top             =   3540
         Width           =   2985
      End
      Begin VB.TextBox cpObs 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   900
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   3600
         Width           =   6390
      End
      Begin rdActiveText.ActiveText cpFone2 
         Height          =   315
         Left            =   4035
         TabIndex        =   12
         Top             =   2880
         Width           =   1965
         _ExtentX        =   3466
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpFone 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   2880
         Width           =   2070
         _ExtentX        =   3651
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpIdentidade 
         Height          =   315
         Left            =   4035
         TabIndex        =   10
         Top             =   2535
         Width           =   1965
         _ExtentX        =   3466
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCpf 
         Height          =   315
         Left            =   900
         TabIndex        =   9
         Top             =   2535
         Width           =   2070
         _ExtentX        =   3651
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
         MaxLength       =   14
         TextMask        =   7
         RawText         =   7
         Mask            =   "###.###.###-##"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCep 
         Height          =   315
         Left            =   2190
         TabIndex        =   8
         Top             =   2190
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpEstado 
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   2190
         Width           =   660
         _ExtentX        =   1164
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
         MaxLength       =   2
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCidade 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   1845
         Width           =   5100
         _ExtentX        =   8996
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
         MaxLength       =   30
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpBairro 
         Height          =   315
         Left            =   900
         TabIndex        =   5
         Top             =   1515
         Width           =   5100
         _ExtentX        =   8996
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
         MaxLength       =   30
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpEndereco 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   1185
         Width           =   5100
         _ExtentX        =   8996
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
      Begin rdActiveText.ActiveText cpNome 
         Height          =   315
         Left            =   900
         TabIndex        =   3
         Top             =   840
         Width           =   5100
         _ExtentX        =   8996
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
      Begin rdActiveText.ActiveText cpApartamento 
         Height          =   315
         Left            =   4785
         TabIndex        =   2
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCodigo 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   495
         Width           =   1020
         _ExtentX        =   1799
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpFone2P 
         Height          =   315
         Left            =   -70830
         TabIndex        =   29
         Top             =   2955
         Width           =   1965
         _ExtentX        =   3466
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpFone1P 
         Height          =   315
         Left            =   -73965
         TabIndex        =   28
         Top             =   2940
         Width           =   2070
         _ExtentX        =   3651
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpIdentidadeP 
         Height          =   315
         Left            =   -70830
         TabIndex        =   27
         Top             =   2610
         Width           =   1965
         _ExtentX        =   3466
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCpfP 
         Height          =   315
         Left            =   -73965
         TabIndex        =   26
         Top             =   2595
         Width           =   2070
         _ExtentX        =   3651
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCepP 
         Height          =   315
         Left            =   -72675
         TabIndex        =   25
         Top             =   2250
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpEstadoP 
         Height          =   315
         Left            =   -73965
         TabIndex        =   24
         Top             =   2250
         Width           =   660
         _ExtentX        =   1164
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
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpCidadeP 
         Height          =   315
         Left            =   -73965
         TabIndex        =   23
         Top             =   1905
         Width           =   5100
         _ExtentX        =   8996
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
         MaxLength       =   30
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpBairroP 
         Height          =   315
         Left            =   -73965
         TabIndex        =   22
         Top             =   1575
         Width           =   5100
         _ExtentX        =   8996
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
         MaxLength       =   30
         TextCase        =   1
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpEnderecoP 
         Height          =   315
         Left            =   -73965
         TabIndex        =   21
         Top             =   1245
         Width           =   5100
         _ExtentX        =   8996
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
      Begin rdActiveText.ActiveText cpNomeP 
         Height          =   315
         Left            =   -73965
         TabIndex        =   20
         Top             =   900
         Width           =   5100
         _ExtentX        =   8996
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
      Begin rdActiveText.ActiveText cpValCondominio 
         Height          =   315
         Left            =   -72780
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1170
         _ExtentX        =   2064
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
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpNomeCond 
         Height          =   315
         Left            =   2340
         TabIndex        =   76
         Top             =   5220
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
      Begin rdActiveText.ActiveText cpCodCond 
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   5220
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
         Height          =   195
         Left            =   360
         TabIndex        =   75
         Top             =   4200
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Bloco"
         Height          =   195
         Left            =   3660
         TabIndex        =   74
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Valor do último condomínio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   72
         Top             =   4020
         Width           =   1920
      End
      Begin MSForms.ComboBox cpFracao 
         Height          =   315
         Left            =   2100
         TabIndex        =   15
         Top             =   4560
         Width           =   2010
         DisplayStyle    =   7
         Size            =   "3545;556"
         ListWidth       =   5291
         BoundColumn     =   2
         ColumnCount     =   2
         ListRows        =   5
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         Object.Width           =   "1834;3245"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Condomínio"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   2280
         TabIndex        =   69
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   5700
         TabIndex        =   68
         Top             =   3300
         Width           =   120
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Ctrl + M = Mesmo do inquilino."
         Height          =   195
         Left            =   -69840
         TabIndex        =   65
         Top             =   5160
         Width           =   2115
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Identidade"
         Height          =   195
         Left            =   -71670
         TabIndex        =   64
         Top             =   2700
         Width           =   750
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   -73050
         TabIndex        =   63
         Top             =   2325
         Width           =   285
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -74460
         TabIndex        =   62
         Top             =   990
         Width           =   420
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   -74730
         TabIndex        =   61
         Top             =   1335
         Width           =   690
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   -74430
         TabIndex        =   60
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   -74520
         TabIndex        =   59
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   -74520
         TabIndex        =   58
         Top             =   2325
         Width           =   495
      End
      Begin VB.Label lbCpfProp 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   195
         Left            =   -74340
         TabIndex        =   57
         Top             =   2670
         Width           =   300
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Fone 1"
         Height          =   195
         Left            =   -74535
         TabIndex        =   56
         Top             =   3015
         Width           =   495
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Fone 2"
         Height          =   195
         Left            =   -71415
         TabIndex        =   55
         Top             =   3030
         Width           =   495
      End
      Begin VB.Label lbFracao 
         AutoSize        =   -1  'True
         Caption         =   "% Fração ideal (despesas)"
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   4620
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   315
         TabIndex        =   53
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No."
         Height          =   195
         Left            =   4425
         TabIndex        =   52
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   405
         TabIndex        =   51
         Top             =   930
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   135
         TabIndex        =   50
         Top             =   1275
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   435
         TabIndex        =   49
         Top             =   1605
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   345
         TabIndex        =   48
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   345
         TabIndex        =   47
         Top             =   2265
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   1830
         TabIndex        =   46
         Top             =   2265
         Width           =   285
      End
      Begin VB.Label lbCpfCnpj 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   195
         Left            =   525
         TabIndex        =   45
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Identidade"
         Height          =   195
         Left            =   3210
         TabIndex        =   44
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fone 1"
         Height          =   195
         Left            =   330
         TabIndex        =   43
         Top             =   2955
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fone 2"
         Height          =   195
         Left            =   3450
         TabIndex        =   42
         Top             =   2970
         Width           =   495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Síndico"
         Height          =   195
         Left            =   270
         TabIndex        =   41
         Top             =   3285
         Width           =   555
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
         Height          =   195
         Left            =   495
         TabIndex        =   40
         Top             =   3585
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   60
      ScaleHeight     =   0
      ScaleWidth      =   7215
      TabIndex        =   39
      Top             =   975
      Width           =   7275
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   45
      Picture         =   "Associados.frx":0044
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1035
      Picture         =   "Associados.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2025
      Picture         =   "Associados.frx":0658
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3015
      Picture         =   "Associados.frx":0962
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4005
      Picture         =   "Associados.frx":0C6C
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   4995
      Picture         =   "Associados.frx":0F76
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6345
      Picture         =   "Associados.frx":1280
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   45
      Width           =   990
   End
End
Attribute VB_Name = "Associados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vbGrava As Integer
Private vbBook  As Variant
Private vbIndex As String
Private vbRetg  As Boolean
Dim iSind As Integer
Dim rsInquilino As Recordset
Dim tbCondominio As Recordset

Private Sub cmdAlterar_Click()
  Travar (False)
  Botoes (False)
  iSind = cpSindico.ListIndex
  vbGrava = 2
  cpCodigo.SetFocus
End Sub

Private Sub cmdDesfazer_Click()
  Resp = MsgBox("Desfazer as alterações?", vbQuestion + vbYesNo + vbDefaultButton2, "Desfazer")
  If Resp = vbYes Then
    With rsInquilino
      If vbGrava = 1 Then
        If .RecordCount > 0 Then
          .MoveFirst
          LerDados
        Else
          Limpar
        End If
      ElseIf vbGrava = 2 Then
        If .EOF And .BOF Then
          Limpar
        Else
          LerDados
        End If
      End If
    End With
    Botoes (True)
    Travar (True)
  End If
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Errado
  Resp = MsgBox("Todos os dados relacionados a este inquilino serão excluidos." & vbCrLf & "Confirma a exclusão de '" + cpNome.Text + "' do cadastro?", vbQuestion + vbYesNo + vbDefaultButton2, "Excluir")
  If Resp = vbYes Then
    DBEngine.BeginTrans
    With rsInquilino
      db.Execute "delete from fracao where id_associado = " & cpCodigo.Text
      db.Execute "delete from boletos where cdsc = " & cpCodigo.Text
      db.Execute "delete from descontos where id_inquilino = " & cpCodigo.Text
      db.Execute "delete from despesainquilino where id_associado = " & cpCodigo.Text
      db.Execute "delete from fundos where id_associado = " & cpCodigo.Text
      db.Execute "delete from leitura_individual where id_associado = " & cpCodigo.Text
      db.Execute "delete from porfora where associado = " & cpCodigo.Text
      db.Execute "delete from subdesp_fixa where id_associado = " & cpCodigo.Text
      .Delete
      DBEngine.CommitTrans
      .MovePrevious
      If .BOF Then
        .MoveNext
        If .EOF Then
          Limpar
        Else
          LerDados
        End If
      Else
        LerDados
      End If
    End With
    MsgBox "Todos os dados foram excluidos com sucesso...", vbInformation + vbOKOnly, "Aviso"
  End If
Sair:
  Exit Sub
Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbCritical + vbOKOnly, "Erro"
  DBEngine.Rollback
  Resume Sair
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdFracao_Click()
  If Val(cpCodigo.Text) > 0 Then
    CadFracao.lCod = cpCodigo.Text
    CadFracao.Show vbModal
    Call PreencheCombo2(cpCodigo.Text)
    If cpFracao.ListCount > 0 Then
      cpFracao.ListIndex = 0
    End If
  End If
End Sub

Private Sub cmdIncluir_Click()
  Limpar
  Travar (False)
  Botoes (False)
  iSind = -1
  cpCodigo.Text = ProximoCodigo("associados", "Codigo")
  vbGrava = 1
  cpTipo.SetFocus
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  lAssociado.Show 1
  If RetCodigo > 0 Then
    With rsInquilino
      .FindFirst "codigo = " & RetCodigo
      If Not .NoMatch Then
        LerDados
      End If
    End With
  End If
End Sub

Private Sub cmdLocCond_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        cpNomeCond.Text = !Nome
        cpCodCond.Text = !Codigo
        CombinaInfo
      End If
    End With
  End If
End Sub

Private Sub cmdSalvar_Click()
  Dim CodConferi As Long
  
  If Val(cpCodCond.Text) <= 0 Then
    MsgBox "Escolha um condominio.", vbInformation + vbOKOnly, "Aviso"
    cpCodCond.SetFocus
    Exit Sub
  End If
  
  If CpPgto.Value = 0 Then
    If cpCpf.Text <> "" Then
      If lbCpfCnpj.Caption = "CNPJ" Then
        If MinhaDll.Valida_CGC(cpCpf.Text) = sdErro Then
          MsgBox "O CNPJ informado para o inquilino não é válido.", vbCritical + vbOKOnly, "Aviso"
          cpCpf.SetFocus
          Exit Sub
        End If
      Else
        If MinhaDll.Valida_Cpf(cpCpf.Text) = sdErro Then
          MsgBox "O CPF informado para o inquilino não é válido.", vbCritical + vbOKOnly, "Aviso"
          cpCpf.SetFocus
          Exit Sub
        End If
      End If
    End If
  Else
    If cpCpfP.Text <> "" Then
      If lbCpfProp.Caption = "CNPJ" Then
        If MinhaDll.Valida_CGC(cpCpfP.Text) = sdErro Then
          MsgBox "O CNPJ informado para o proprietário não é válido.", vbCritical + vbOKOnly, "Aviso"
          cpCpfP.SetFocus
          Exit Sub
        End If
      Else
        If MinhaDll.Valida_Cpf(cpCpfP.Text) = sdErro Then
          MsgBox "O CPF informado para o proprietário não é válido.", vbCritical + vbOKOnly, "Aviso"
          cpCpfP.SetFocus
          Exit Sub
        End If
      End If
    End If
  End If
  If ChAcumulado.Value = 1 Then
    CodConferi = PrazoBoleto(cpCodCond.Text)
    If CodConferi < 1 Or CodConferi > 20 Then
      MsgBox "O prazo limite para pagamento do boleto depois do vencimento infomado para o condomínio não permite acumular débitos.", vbInformation + vbOKOnly, "Aviso"
      ChAcumulado.SetFocus
      Exit Sub
    End If
  End If
  
  If cpTipo.ListIndex < 0 Then
    MsgBox "Escolha um tipo.", vbInformation + vbOKOnly, "Aviso"
    cpTipo.SetFocus
    Exit Sub
  End If
  
  If cpApartamento.Text = "" Then
    MsgBox "Informe o(s) número(s).", vbInformation + vbOKOnly, "Aviso"
    cpApartamento.SetFocus
    Exit Sub
  End If
  
'  If cpFracao.ListCount = 0 Then
'    MsgBox "Informe " & lbFracao.Caption & " do inquilino.", vbInformation + vbOKOnly, "Aviso"
'    cmdFracao.SetFocus
'    Exit Sub
'  End If
  
  If cpBlocos.ListIndex < 0 Then
    MsgBox "Escolha 0 bloco a qual pertence a unidade.", vbInformation + vbOKOnly, "Aviso"
    cpBlocos.SetFocus
    Exit Sub
  End If
  
  Select Case cpSindico.ListIndex
    Case 0, 1
      cpValor.Text = 0
    Case 2, 3
      If CDbl(cpValor.Text) <= 0 Then
        MsgBox "Informe o % a pagar para o sindico so valor das despesas.", vbInformation + vbOKOnly, "Aviso"
        cpValor.SetFocus
        Exit Sub
      End If
  End Select
  
  If cpSindico.ListIndex > 0 Then
    CodConferi = CodigoSindico(cpCodCond.Text)
    If (CodConferi > 0 And CodConferi <> cpCodigo.Text) Then
      MsgBox "O condomínio " & cpNomeCond.Text & " já possui um síndico.", vbInformation + vbOKOnly, "Aviso"
      cpSindico.SetFocus
      Exit Sub
    End If
  End If
  
  If Len(cpEmail.Text) > 0 Then
    If Not ValidEmail(cpEmail.Text) Then
        MsgBox "O e-mail informado não é válido.", vbInformation + vbOKOnly, "Aviso"
        cpEmail.SetFocus
        Exit Sub
    End If
  End If
  
  If cpValCondominio.Text = "" Then
    cpValCondominio.Text = 0
  End If
  
  If vbGrava = 1 Then
    If GravaDados(1) Then
      Travar (True)
      Botoes (True)
    End If
  ElseIf vbGrava = 2 Then
    If GravaDados(2, vbBook) Then
      Travar (True)
      Botoes (True)
    End If
  End If
End Sub

Private Sub cpApartamento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpNome.SetFocus
  End If
End Sub

Private Sub cpBairro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCidade.SetFocus
  End If
End Sub

Private Sub cpBairroP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCidadeP.SetFocus
  End If
End Sub

Private Sub cpCep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCpf.SetFocus
  End If
End Sub

Private Sub cpCepP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCpfP.SetFocus
  End If
End Sub

Private Sub cpCidade_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEstado.SetFocus
  End If
End Sub

Private Sub cpCidade_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    If KeyCode = 113 Then
      If Not cpCidade.Locked Then
        RetCidade(0) = ""
        RetCidade(1) = ""
        RetCidade(2) = ""
        lCidades.Show 1
        If Not (RetCidade(0) = "") Then
          cpCidade.Text = RetCidade(0)
          cpEstado.Text = RetCidade(1)
          cpCep.Text = RetCidade(2)
        End If
      End If
    End If
  End If
End Sub

Private Sub cpCidadeP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEstadoP.SetFocus
  End If
End Sub

Private Sub cpCidadeP_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift = 0 Then
    If KeyCode = 113 Then
      If Not cpCidadeP.Locked Then
        RetCidade(0) = ""
        RetCidade(1) = ""
        RetCidade(2) = ""
        lCidades.Show 1
        If Not (RetCidade(0) = "") Then
          cpCidadeP.Text = RetCidade(0)
          cpEstadoP.Text = RetCidade(1)
          cpCepP.Text = RetCidade(2)
        End If
      End If
    End If
  End If
End Sub

Private Sub cpCodCond_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If Val(cpCodCond.Text) > 0 Then
      With tbCondominio
        .Index = "codigoid"
        .Seek "=", cpCodCond.Text
        If Not .NoMatch Then
          cpNomeCond.Text = !Nome
          cpNomeP.SetFocus
        Else
          MsgBox "O código informado não existe!", vbInformation + vbOKOnly, "Aviso"
          cpCodCond.SetFocus
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
            cpNomeCond.Text = !Nome
            cpCodCond.Text = !Codigo
          End If
        End With
        cpNomeP.SetFocus
      End If
    End If
  End If
End Sub

Private Sub cpCodCond_LostFocus()
  CombinaInfo
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpTipo.SetFocus
  End If
End Sub

Private Sub CombinaInfo()
  If Val(cpCodCond.Text) > 0 Then
    Select Case TipoDistDespesa(cpCodCond.Text)
      Case 0, 1
        lbFracao.Caption = "% Fração ideal (despesas)"
      Case 2, 3
        lbFracao.Caption = "Cota ideal (despesas)"
      Case 4
        lbFracao.Caption = "Valor fixo"
    End Select
    PegaEndereco cpCodCond.Text
  End If
End Sub

Private Sub cpCpf_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpIdentidade.SetFocus
  End If
End Sub

Private Sub cpCpfP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpIdentidadeP.SetFocus
  End If
End Sub

Private Sub cpEmail_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdFracao.Enabled Then cmdFracao.SetFocus
  End If
End Sub

Private Sub cpEndereco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpBairro.SetFocus
  End If
End Sub

Private Sub cpEnderecoP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpBairroP.SetFocus
  End If
End Sub

Private Sub cpEstado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCep.SetFocus
  End If
End Sub

Private Sub cpEstadoP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCepP.SetFocus
  End If
End Sub

Private Sub cpFone_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFone2.SetFocus
  End If
End Sub

Private Sub cpFone1P_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFone2P.SetFocus
  End If
End Sub

Private Sub cpFone2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpSindico.SetFocus
  End If
End Sub

Private Sub cpFone2P_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    CpPgto.SetFocus
  End If
End Sub

Private Sub cpIdentidade_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFone.SetFocus
  End If
End Sub

Private Sub cpIdentidadeP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFone1P.SetFocus
  End If
End Sub

Private Sub cpNome_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEndereco.SetFocus
  End If
End Sub

Private Sub cpNomeP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEnderecoP.SetFocus
  End If
End Sub

Private Sub cpObs_GotFocus()
  cpObs.SelStart = 0
  cpObs.SelLength = Len(cpObs.Text)
End Sub

Private Sub cpObs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEmail.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

'Private Sub cpPercentoal_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'    KeyAscii = 0
'    ChAcumulado.SetFocus
'  End If
'End Sub

Private Sub cpSindico_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  End If
End Sub

Private Sub cpSindico_LostFocus()
  Select Case cpSindico.ListIndex
    Case 0, 1
      cpValor.Locked = True
      cpObs.SetFocus
    Case 2, 3, 4, 5
      cpValor.Locked = False
  End Select
End Sub

Private Sub cpTipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpApartamento.SetFocus
  End If
End Sub

Private Sub cpValCondominio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpObs.SetFocus
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpObs.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If cmdSalvar.Enabled = False Then
    If Shift = 0 Then
      Select Case KeyCode
        Case 33
          With rsInquilino
            If Not .BOF Then
              .MovePrevious
              If .BOF Then
                .MoveNext
                If .EOF Then
                  Limpar
                Else
                  LerDados
                End If
              Else
                LerDados
              End If
            End If
          End With
        Case 34
          With rsInquilino
            If Not .EOF Then
              .MoveNext
              If .EOF Then
                .MovePrevious
                If .BOF Then
                  Limpar
                Else
                  LerDados
                End If
              Else
                LerDados
              End If
            End If
          End With
      End Select
    ElseIf Shift = 2 Then
      Select Case KeyCode
        Case 33
          With rsInquilino
            If .RecordCount > 0 Then
              .MoveFirst
              LerDados
            End If
          End With
        Case 34
          With rsInquilino
            If .RecordCount > 0 Then
              .MoveLast
              LerDados
            End If
          End With
        Case 77
          oMesmo
      End Select
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
  Refresh
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  KeyPreview = True
  Call PreencheCombo(cpTipo, "tipos", "codigo", "descricao")
  Call PreencheCombo(cpBlocos, "blocos", "id_bloco", "nome_bloco")
  cpSindico.AddItem ("NÃO")
  cpSindico.AddItem ("SIM")
  cpSindico.AddItem ("SIM - Recebe % do Valor da despesa")
  cpSindico.AddItem ("SIM - Recebe % do Valor da despesa, e distribui para os demais")
  Set rsInquilino = db.OpenRecordset("select * from associados order by condominio, tipo, apartamento;", dbOpenDynaset)
  With rsInquilino
    If .RecordCount > 0 Then
      .MoveFirst
      LerDados
    End If
  End With
  Botoes (True)
  Travar (True)
  Abas.Tab = 0
End Sub

Private Sub TravaBotoes()
  cmdDesfazer.Enabled = False
  cmdSalvar.Enabled = False
  cmdAlterar.Enabled = False
  cmdExcluir.Enabled = False
  cmdFechar.Enabled = True
  cmdIncluir.Enabled = False
  cmdLocalizar.Enabled = False
End Sub

Private Sub Botoes(Tipo As Boolean)
  cmdDesfazer.Enabled = Not Tipo
  cmdSalvar.Enabled = Not Tipo
  cmdAlterar.Enabled = Tipo
  cmdExcluir.Enabled = Tipo
  cmdFechar.Enabled = Tipo
  cmdIncluir.Enabled = Tipo
  cmdLocalizar.Enabled = Tipo
  cmdFracao.Enabled = Tipo
End Sub

Private Sub LerDados()
  cpCpfP.ForeColor = vbBlack
  With rsInquilino
    If .RecordCount > 0 Then
      cpApartamento.Text = IIf(IsNull(!Apartamento), "", !Apartamento)
      cpBairro.Text = IIf(IsNull(!bairro), "", !bairro)
      cpCep.Text = IIf(IsNull(!cep), "", !cep)
      cpCidade.Text = IIf(IsNull(!Cidade), "", !Cidade)
      cpCodigo.Text = IIf(IsNull(!Codigo), "", !Codigo)
      cpBlocos.ListIndex = MostraBloco(!id_bloco)
      cpDescVira.Text = !descvira & ""
      cpEmail.Text = !email & ""
      cpCodCond.Text = !Condominio
      cpNomeCond.Text = NomeCondominio(!Condominio)
      If !acumulado = "S" Then
        ChAcumulado.Value = 1
      Else
        ChAcumulado.Value = 0
      End If
      If !ViraFundo = True Then
        chViraFundo.Value = 1
      Else
        chViraFundo.Value = 0
      End If
      If Len(SoNumeros(!cpf & "")) > 13 Then
        lbCpfCnpj.Caption = "CNPJ"
        cpCpf.TextMask = [CGC Mask]
        cpCpf.Text = IIf(IsNull(!cpf), "", !cpf)
        If MinhaDll.Valida_CGC(cpCpf.Text) = sdErro Then
          cpCpf.ForeColor = vbRed
        Else
          cpCpf.ForeColor = vbBlack
        End If
      Else
        lbCpfCnpj.Caption = "CPF"
        cpCpf.TextMask = [CPF Mask]
        cpCpf.Text = IIf(IsNull(!cpf), "", !cpf)
        If MinhaDll.Valida_Cpf(cpCpf.Text) = sdErro Then
          cpCpf.ForeColor = vbRed
        Else
          cpCpf.ForeColor = vbBlack
        End If
      End If
      cpEndereco.Text = IIf(IsNull(!endereco), "", !endereco)
      cpEstado.Text = IIf(IsNull(!estado), "", !estado)
      cpFone.Text = IIf(IsNull(!Fone1), "", !Fone1)
      cpFone2.Text = IIf(IsNull(!Fone2), "", !Fone2)
      cpIdentidade.Text = IIf(IsNull(!Identidade), "", !Identidade)
      cpNome.Text = IIf(IsNull(!Proprietario), "", !Proprietario)
      cpObs.Text = IIf(IsNull(!Observacoes), "", !Observacoes)
      If !Sindico >= 0 Then
        cpSindico.ListIndex = !Sindico
      Else
        cpSindico.ListIndex = 0
      End If
      cpValor.Text = !valorpagar
      cpValCondominio.Text = IIf(IsNull(!ValCondominio), "0,00", Format$(!ValCondominio, "#,##0.00"))
      Call PreencheCombo2(!Codigo)
      If cpFracao.ListCount > 0 Then
        cpFracao.ListIndex = 0
      End If
      cpBairroP.Text = IIf(IsNull(!PBairro), "", !PBairro)
      cpCepP.Text = IIf(IsNull(!PCep), "", !PCep)
      cpCidadeP.Text = IIf(IsNull(!PCidade), "", !PCidade)
      If Len(SoNumeros(!pcpf & "")) > 13 Then
        lbCpfProp.Caption = "CNPJ"
        cpCpfP.TextMask = [CGC Mask]
        cpCpfP.Text = IIf(IsNull(!pcpf), "", !pcpf)
        If MinhaDll.Valida_CGC(cpCpfP.Text) = sdErro Then
          cpCpfP.ForeColor = vbRed
        Else
          cpCpfP.ForeColor = vbBlack
        End If
      Else
        lbCpfProp.Caption = "CPF"
        cpCpfP.TextMask = [CPF Mask]
        cpCpfP.Text = IIf(IsNull(!pcpf), "", !pcpf)
        If MinhaDll.Valida_Cpf(cpCpfP.Text) = sdErro Then
          cpCpfP.ForeColor = vbRed
        Else
          cpCpfP.ForeColor = vbBlack
        End If
      End If
      cpEnderecoP.Text = IIf(IsNull(!PEndereco), "", !PEndereco)
      cpEstadoP.Text = IIf(IsNull(!PEstado), "", !PEstado)
      cpFone1P.Text = IIf(IsNull(!PFone1), "", !PFone1)
      cpFone2P.Text = IIf(IsNull(!PFone2), "", !PFone2)
      cpIdentidadeP.Text = IIf(IsNull(!PIdentidade), "", !PIdentidade)
      cpNomeP.Text = IIf(IsNull(!Nome), "", !Nome)
      CpPgto.Value = IIf(IsNull(!boleto), 0, !boleto)
      If Not IsNull(!Tipo) Then
        cpTipo.ListIndex = AchaTipo(!Tipo)
      Else
        cpTipo.ListIndex = -1
      End If
      Select Case TipoDistDespesa(!Condominio)
        Case 0, 1
          lbFracao.Caption = "% Fração ideal (despesas)"
        Case 2, 3
          lbFracao.Caption = "Cota ideal (despesas)"
        Case 4
          lbFracao.Caption = "Valor fixo"
      End Select
      Me.Caption = "Cadastro de inquilino: " & !Tipo & " " & !Apartamento
    End If
  End With
End Sub

Private Function GravaDados(nTipo As Integer, Optional nBook As Variant) As Boolean
On Error GoTo Errado
  GravaDados = False
  With rsInquilino
    If nTipo = 1 Then
      .AddNew
      !Codigo = cpCodigo.Text
    ElseIf nTipo = 2 Then
      .Edit
    End If
    !Apartamento = cpApartamento.Text
    !bairro = cpBairro.Text
    If cpBlocos.ListIndex >= 0 Then
      !id_bloco = cpBlocos.ItemData(cpBlocos.ListIndex)
    Else
      !id_bloco = 1
    End If
    !descvira = cpDescVira.Text & "."
    !email = cpEmail.Text
    !cep = cpCep.Text
    !Cidade = cpCidade.Text
    !Condominio = cpCodCond.Text
    !cpf = cpCpf.Text
    !endereco = cpEndereco.Text
    If chViraFundo.Value = 1 Then
      !ViraFundo = True
    Else
      !ViraFundo = False
    End If
    !estado = cpEstado.Text
    !Fone1 = cpFone.Text
    !Fone2 = cpFone2.Text
    !Identidade = cpIdentidade.Text
    !Proprietario = cpNome.Text
    !Observacoes = cpObs.Text
    !Postagem = "N"
    Select Case cpSindico.ListIndex
      Case 0
        !Sindico = 0
        !pagar = 0
        !valorpagar = 0
        !distpagar = 0
      Case 1
        !Sindico = 1
        !pagar = 0
        !valorpagar = 0
        !distpagar = 0
      Case 2
        !Sindico = 2
        !pagar = 1
        !valorpagar = cpValor.Text
        !distpagar = 0
      Case 3
        !Sindico = 3
        !pagar = 1
        !valorpagar = cpValor.Text
        !distpagar = 0
    End Select
    !ValCondominio = cpValCondominio.Text
    !Fracao = 0
    !PBairro = cpBairroP.Text
    !PCep = cpCepP.Text
    !PCidade = cpCidadeP.Text
    !pcpf = cpCpfP.Text
    !PEndereco = cpEnderecoP.Text
    !PEstado = cpEstadoP.Text
    !PFone1 = cpFone1P.Text
    !PFone2 = cpFone2P.Text
    !PIdentidade = cpIdentidadeP.Text
    !Nome = cpNomeP.Text
    !boleto = CpPgto.Value
    !Tipo = cpTipo.Text
    If ChAcumulado.Value = 1 Then
      !acumulado = "S"
    Else
      !acumulado = "N"
    End If
    .Update
    .Bookmark = .LastModified
  End With
  LerDados
  GravaDados = True

Sair:
  Exit Function

Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Function

Private Sub Limpar()
  cpApartamento.Text = ""
  chViraFundo.Value = 0
  cpDescVira.Text = ""
  cpBlocos.ListIndex = BlocoUnico
  cpEmail.Text = ""
  cpBairro.Text = ""
  cpCep.Text = "36570-000"
  cpCidade.Text = ""
  cpCodigo.Text = ""
  cpCpf.Text = ""
  cpEndereco.Text = ""
  cpEstado.Text = "MG"
  cpFone.Text = ""
  cpFone2.Text = ""
  cpIdentidade.Text = ""
  cpNome.Text = ""
  cpObs.Text = ""
  cpValor.Text = 0
  cpSindico.ListIndex = 0
  cpValCondominio.Text = "0"
  cpFracao.ListIndex = -1
  cpBairroP.Text = ""
  cpCepP.Text = "36570-000"
  cpCidadeP.Text = "VIÇOSA"
  cpCpfP.Text = ""
  cpEnderecoP.Text = ""
  cpEstadoP.Text = "MG"
  cpFone1P.Text = ""
  cpFone2P.Text = ""
  cpIdentidadeP.Text = ""
  cpNomeP.Text = ""
  CpPgto.Value = 0
  ChAcumulado.Value = 0
  cpTipo.ListIndex = -1
  cpNomeCond.Text = ""
  cpCodCond.Text = ""
End Sub

Private Sub Travar(Tipo As Boolean)
  cpApartamento.Locked = Tipo
  chViraFundo.Enabled = Not Tipo
  cpDescVira.Locked = Tipo
  cpBairro.Locked = Tipo
  cpEmail.Locked = Tipo
  cpBlocos.Locked = Tipo
  cpCep.Locked = Tipo
  cpCidade.Locked = Tipo
  cpCodigo.Locked = Tipo
  cpCpf.Locked = Tipo
  cpEndereco.Locked = Tipo
  cpEstado.Locked = Tipo
  cpFone.Locked = Tipo
  cpFone2.Locked = Tipo
  cpIdentidade.Locked = Tipo
  cpNome.Locked = Tipo
  cpObs.Locked = Tipo
  cpSindico.Locked = Tipo
  cpValCondominio.Locked = Tipo
  'cpFracao.Locked = Tipo
  cpBairroP.Locked = Tipo
  cpCepP.Locked = Tipo
  cpCidadeP.Locked = Tipo
  cpCpfP.Locked = Tipo
  cpEnderecoP.Locked = Tipo
  cpEstadoP.Locked = Tipo
  cpFone1P.Locked = Tipo
  cpFone2P.Locked = Tipo
  cpIdentidadeP.Locked = Tipo
  cpNomeP.Locked = Tipo
  CpPgto.Enabled = Not Tipo
  ChAcumulado.Enabled = Not Tipo
  cpTipo.Locked = Tipo
  cpCodCond.Locked = Tipo
  cpNomeCond.Locked = Tipo
  cmdLocCond.Enabled = Not Tipo
End Sub

Private Sub oMesmo()
  cpBairroP.Text = cpBairro.Text
  cpCepP.Text = cpCep.Text
  cpCidadeP.Text = cpCidade.Text
  cpCpfP.Text = cpCpf.Text
  cpEnderecoP.Text = cpEndereco.Text
  cpEstadoP.Text = cpEstado.Text
  cpFone1P.Text = cpFone.Text
  cpFone2P.Text = cpFone2.Text
  cpIdentidadeP.Text = cpIdentidade.Text
  cpNomeP.Text = cpNome.Text
End Sub

Private Function AchaTipo(ByVal sTipo As String) As Integer
  Dim ii As Integer
  AchaTipo = -1
  For ii = 0 To cpTipo.ListCount - 1
    If cpTipo.List(ii) = sTipo Then
      AchaTipo = ii
      Exit For
    End If
  Next ii
End Function

Private Sub PegaEndereco(ByVal nCod As Long)
  With tbCondominio
    .Index = "codigoid"
    .Seek "=", nCod
    If Not .NoMatch Then
      cpEndereco.Text = !endereco & ""
      cpBairro.Text = !bairro & ""
      cpEstado.Text = !estado & ""
      cpCidade.Text = !Cidade & ""
      cpCep.Text = !cep & ""
    End If
  End With
End Sub

Private Function MostraBloco(ByVal iCod As Long) As Integer
  Dim i As Integer
  MostraBloco = -1
  For i = 0 To cpBlocos.ListCount - 1
    If cpBlocos.ItemData(i) = iCod Then
      MostraBloco = i
      Exit For
    End If
  Next i
End Function

Private Function BlocoUnico() As Integer
  Dim i As Integer
  BlocoUnico = -1
  For i = 0 To cpBlocos.ListCount - 1
    If UCase(cpBlocos.List(i)) = "ÚNICO" Or UCase(cpBlocos.List(i)) = "UNICO" Then
      BlocoUnico = i
      Exit For
    End If
  Next i
End Function

Private Sub PreencheCombo2(ByVal aCod As Long)
  
  Dim rs As Recordset
  cpFracao.Clear
  Set rs = db.OpenRecordset("Select * from fracao where id_associado = " & aCod & " order by fracao;", dbOpenDynaset)
  
  With rs
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        cpFracao.AddItem
        cpFracao.List(cpFracao.ListCount - 1, 0) = Format$(.Fields("fracao").Value, "#0.0000####")
        cpFracao.List(cpFracao.ListCount - 1, 1) = .Fields("descricao").Value
        .MoveNext
      Loop
    End If
  End With
  Set rs = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set rsInquilino = Nothing
  Set tbCondominio = Nothing
End Sub

Private Sub lbCpfCnpj_Click()
  If lbCpfCnpj.Caption = "CPF" Then
    lbCpfCnpj.Caption = "CNPJ"
    cpCpf.TextMask = [CGC Mask]
  Else
    lbCpfCnpj.Caption = "CPF"
    cpCpf.TextMask = [CPF Mask]
  End If
End Sub

Private Sub lbCpfProp_Click()
  If lbCpfProp.Caption = "CPF" Then
    lbCpfProp.Caption = "CNPJ"
    cpCpfP.TextMask = [CGC Mask]
  Else
    lbCpfProp.Caption = "CPF"
    cpCpfP.TextMask = [CPF Mask]
  End If
End Sub
