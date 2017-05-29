VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Condominio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro do condomínio"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ControlBox      =   0   'False
   Icon            =   "Condominio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Abas 
      Height          =   4575
      Left            =   60
      TabIndex        =   40
      Top             =   960
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   8070
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados do condomínio"
      TabPicture(0)   =   "Condominio.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cpCodigo"
      Tab(0).Control(1)=   "cpInscricao"
      Tab(0).Control(2)=   "cpCgc"
      Tab(0).Control(3)=   "cpFonePortaria"
      Tab(0).Control(4)=   "cpFone"
      Tab(0).Control(5)=   "cpCep"
      Tab(0).Control(6)=   "cpEstado"
      Tab(0).Control(7)=   "cpCidade"
      Tab(0).Control(8)=   "cpBairro"
      Tab(0).Control(9)=   "cpEndereco"
      Tab(0).Control(10)=   "cpNome"
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(15)=   "Label5"
      Tab(0).Control(16)=   "Label6"
      Tab(0).Control(17)=   "Label7"
      Tab(0).Control(18)=   "Label8"
      Tab(0).Control(19)=   "Label17"
      Tab(0).Control(20)=   "Label18"
      Tab(0).Control(21)=   "Label19"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Despesas"
      TabPicture(1)   =   "Condominio.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "cpJuros"
      Tab(1).Control(2)=   "cpMulta"
      Tab(1).Control(3)=   "cpFrase"
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(6)=   "Label22"
      Tab(1).Control(7)=   "Label21"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Dados para boletos"
      TabPicture(2)   =   "Condominio.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lbCedente"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label27"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label15"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label25"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label24"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lbCarteira"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label9"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label14"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cpCarteira"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cpDiasRec"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cpDiaVence"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cpAgCedente"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cpOperacao"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cpConta"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "cpCdAgencia"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "cpBanco"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "cpTipoCobranca"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Frame1"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "cpDiasProtesto"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).ControlCount=   23
      Begin rdActiveText.ActiveText cpDiasProtesto 
         Height          =   315
         Left            =   1500
         TabIndex        =   26
         Top             =   2100
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
      Begin VB.Frame Frame2 
         Caption         =   "Divisão das despesas"
         Height          =   2235
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   7155
         Begin VB.OptionButton OpValorFixo 
            Caption         =   "Valor fixo idenpendente das despesas"
            Height          =   195
            Left            =   300
            TabIndex        =   15
            Top             =   1800
            Width           =   3015
         End
         Begin VB.OptionButton OpCotaSindico 
            Caption         =   "Divisão pela cota dos condôminos, menos para o síndico (seu valor de despesas entra como despesas para os demais.)"
            Height          =   375
            Left            =   300
            TabIndex        =   14
            Top             =   1320
            Width           =   5355
         End
         Begin VB.OptionButton OpCotaGeral 
            Caption         =   "Divisão por cotas dos condôminos"
            Height          =   195
            Left            =   300
            TabIndex        =   13
            Top             =   1020
            Width           =   2775
         End
         Begin VB.OptionButton OpSindico 
            Caption         =   $"Condominio.frx":0060
            Height          =   435
            Left            =   300
            TabIndex        =   12
            Top             =   540
            Width           =   5685
         End
         Begin VB.OptionButton OpGeral 
            Caption         =   "Divisão pela fração ideal entre todos os condôminos."
            Height          =   195
            Left            =   300
            TabIndex        =   11
            Top             =   300
            Value           =   -1  'True
            Width           =   4050
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Titular apresentado no boleto"
         Height          =   1395
         Left            =   120
         TabIndex        =   62
         Top             =   3060
         Width           =   7275
         Begin rdActiveText.ActiveText cpCnpj 
            Height          =   315
            Left            =   1260
            TabIndex        =   31
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
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
            MaxLength       =   18
            TextMask        =   8
            RawText         =   8
            Mask            =   "##.###.###/####-##"
            FontName        =   "MS Sans Serif"
            FontSize        =   8,25
         End
         Begin VB.CheckBox chTitular 
            Caption         =   "Próprio condomínio"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   300
            Width           =   1695
         End
         Begin rdActiveText.ActiveText cpRazaoSocial 
            Height          =   315
            Left            =   1260
            TabIndex        =   30
            Top             =   600
            Width           =   5775
            _ExtentX        =   10186
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
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   180
            TabIndex        =   64
            Top             =   1020
            Width           =   405
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   660
            Width           =   945
         End
      End
      Begin VB.ComboBox cpTipoCobranca 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   900
         Width           =   4275
      End
      Begin VB.ComboBox cpBanco 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   4275
      End
      Begin rdActiveText.ActiveText cpJuros 
         Height          =   315
         Left            =   -70980
         TabIndex        =   18
         Top             =   3705
         Width           =   1095
         _ExtentX        =   1931
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
         MaxLength       =   15
         TextMask        =   4
         RawText         =   4
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText cpMulta 
         Height          =   315
         Left            =   -73260
         TabIndex        =   17
         Top             =   3690
         Width           =   1005
         _ExtentX        =   1773
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
         MaxLength       =   15
         TextMask        =   4
         RawText         =   4
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpFrase 
         Height          =   315
         Left            =   -74700
         TabIndex        =   16
         Top             =   3240
         Width           =   6675
         _ExtentX        =   11774
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
         MaxLength       =   100
         RawText         =   0
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpCodigo 
         Height          =   315
         Left            =   -73740
         TabIndex        =   0
         Top             =   600
         Width           =   1020
         _ExtentX        =   1799
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
         TextMask        =   3
         RawText         =   3
         FontName        =   "Arial"
         FontSize        =   9,75
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpInscricao 
         Height          =   315
         Left            =   -70770
         TabIndex        =   8
         Top             =   2640
         Width           =   1950
         _ExtentX        =   3440
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
         TextCase        =   1
         RawText         =   0
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpCgc 
         Height          =   315
         Left            =   -73755
         TabIndex        =   7
         Top             =   2640
         Width           =   1950
         _ExtentX        =   3440
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
         MaxLength       =   18
         TextMask        =   9
         RawText         =   9
         Mask            =   "##.###.###/####-##"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpFonePortaria 
         Height          =   315
         Left            =   -70770
         TabIndex        =   10
         Top             =   2970
         Width           =   1950
         _ExtentX        =   3440
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
         MaxLength       =   17
         TextMask        =   9
         RawText         =   9
         Mask            =   "(#xx##) ####-####"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpFone 
         Height          =   315
         Left            =   -73770
         TabIndex        =   9
         Top             =   2985
         Width           =   2190
         _ExtentX        =   3863
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
         MaxLength       =   17
         TextMask        =   9
         RawText         =   9
         Mask            =   "(#xx##) ####-####"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpCep 
         Height          =   315
         Left            =   -72525
         TabIndex        =   6
         Top             =   2310
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   -73755
         TabIndex        =   5
         Top             =   2310
         Width           =   585
         _ExtentX        =   1032
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
         Left            =   -73755
         TabIndex        =   4
         Top             =   1965
         Width           =   4755
         _ExtentX        =   8387
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
      Begin rdActiveText.ActiveText cpBairro 
         Height          =   315
         Left            =   -73755
         TabIndex        =   3
         Top             =   1635
         Width           =   4755
         _ExtentX        =   8387
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
      Begin rdActiveText.ActiveText cpEndereco 
         Height          =   315
         Left            =   -73755
         TabIndex        =   2
         Top             =   1305
         Width           =   4755
         _ExtentX        =   8387
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
      Begin rdActiveText.ActiveText cpNome 
         Height          =   315
         Left            =   -73755
         TabIndex        =   1
         Top             =   960
         Width           =   4755
         _ExtentX        =   8387
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
      Begin rdActiveText.ActiveText cpCdAgencia 
         Height          =   315
         Left            =   5880
         TabIndex        =   20
         Top             =   480
         Width           =   750
         _ExtentX        =   1323
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
         MaxLength       =   3
         TextMask        =   9
         RawText         =   9
         Mask            =   "###"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpConta 
         Height          =   315
         Left            =   5595
         TabIndex        =   25
         Top             =   1680
         Width           =   1470
         _ExtentX        =   2593
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
         MaxLength       =   8
         TextMask        =   9
         RawText         =   9
         Mask            =   "########"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpOperacao 
         Height          =   315
         Left            =   1500
         TabIndex        =   23
         Top             =   1680
         Width           =   705
         _ExtentX        =   1244
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
         MaxLength       =   3
         TextMask        =   9
         RawText         =   9
         Mask            =   "###"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpAgCedente 
         Height          =   315
         Left            =   3180
         TabIndex        =   24
         Top             =   1680
         Width           =   1125
         _ExtentX        =   1984
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
         MaxLength       =   4
         TextMask        =   9
         RawText         =   9
         Mask            =   "####"
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpDiaVence 
         Height          =   315
         Left            =   3060
         TabIndex        =   27
         Top             =   2580
         Width           =   630
         _ExtentX        =   1111
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
         TextMask        =   3
         RawText         =   3
         FontName        =   "Arial"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText cpDiasRec 
         Height          =   315
         Left            =   5700
         TabIndex        =   28
         Top             =   2580
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
         MaxLength       =   2
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin MSForms.ComboBox cpCarteira 
         Height          =   315
         Left            =   1500
         TabIndex        =   22
         Top             =   1260
         Width           =   990
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1746;556"
         ListWidth       =   10583
         ColumnCount     =   2
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "dias de atrazo (0 conforme contrato)"
         Height          =   195
         Left            =   2220
         TabIndex        =   71
         Top             =   2160
         Width           =   2550
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Protestar após"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label lbCarteira 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2580
         TabIndex        =   69
         Top             =   1260
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Não receber boleto após"
         Height          =   195
         Left            =   3840
         TabIndex        =   68
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "de vencido"
         Height          =   195
         Left            =   6360
         TabIndex        =   67
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dia para o vencimento do condomínio"
         Height          =   195
         Left            =   240
         TabIndex        =   66
         Top             =   2640
         Width           =   2715
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tipo cobrança"
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   600
         Width           =   465
      End
      Begin VB.Label lbCedente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cedente s/díg."
         Height          =   195
         Left            =   4440
         TabIndex        =   59
         Top             =   1770
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Operação"
         Height          =   195
         Left            =   225
         TabIndex        =   58
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   2460
         TabIndex        =   57
         Top             =   1740
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Carteira"
         Height          =   195
         Left            =   240
         TabIndex        =   56
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   -69840
         TabIndex        =   55
         Top             =   3780
         Width           =   120
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "%  Juros p/mês"
         Height          =   195
         Left            =   -72180
         TabIndex        =   54
         Top             =   3780
         Width           =   1080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Multa sobre atraso"
         Height          =   195
         Left            =   -74715
         TabIndex        =   53
         Top             =   3765
         Width           =   1305
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Frase de informação sobre a cobrança sobre atrasados."
         Height          =   195
         Left            =   -74715
         TabIndex        =   52
         Top             =   3045
         Width           =   3945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Condomínio"
         Height          =   195
         Left            =   -74700
         TabIndex        =   51
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   -74535
         TabIndex        =   50
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   -74250
         TabIndex        =   49
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   195
         Left            =   -74340
         TabIndex        =   48
         Top             =   2010
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   -74340
         TabIndex        =   47
         Top             =   2355
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   -72885
         TabIndex        =   46
         Top             =   2400
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fone 1"
         Height          =   195
         Left            =   -74340
         TabIndex        =   45
         Top             =   3075
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fone 2"
         Height          =   195
         Left            =   -71385
         TabIndex        =   44
         Top             =   3045
         Width           =   495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   -74280
         TabIndex        =   43
         Top             =   2745
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição"
         Height          =   195
         Left            =   -71535
         TabIndex        =   42
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Código interno"
         Height          =   195
         Left            =   -74865
         TabIndex        =   41
         Top             =   660
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   60
      ScaleHeight     =   0
      ScaleWidth      =   7395
      TabIndex        =   39
      Top             =   960
      Width           =   7455
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   825
      Left            =   6345
      Picture         =   "Condominio.frx":00E9
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "&Localizar"
      Height          =   825
      Left            =   4995
      Picture         =   "Condominio.frx":03F3
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   825
      Left            =   4005
      Picture         =   "Condominio.frx":06FD
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdDesfazer 
      Caption         =   "&Desfazer"
      Height          =   825
      Left            =   3015
      Picture         =   "Condominio.frx":0A07
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   825
      Left            =   2025
      Picture         =   "Condominio.frx":0D11
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   825
      Left            =   1035
      Picture         =   "Condominio.frx":101B
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Novo"
      Height          =   825
      Left            =   45
      Picture         =   "Condominio.frx":1325
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   60
      Width           =   990
   End
End
Attribute VB_Name = "Condominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type nEndereco
  endereco As String
  Numero As String
  bairro As String
  estado As String
  cep As String
End Type

Dim sTemp As String
Private vbGrava As Integer
Private vbBook  As Variant
Private vbIndex As String
Private vbRetg  As Boolean
Private vbSoundex As New Soundex
Private Mudou As nEndereco
Private tbCondominio As Recordset

Const msgbal As String = "Se '0' o boleto não tem limite para pagamento após vencido, se informado um número de dias o boleto não poderá ser pago após este número de dias de vencido, se um número maior que 20 for informado não permite acumular débito(s) anterior(es)."

Private Sub chTitular_Click()
  If chTitular.Value = 1 Then
    cpRazaoSocial.Text = cpNome.Text
    cpCnpj.Text = cpCgc.Text
  Else
    cpRazaoSocial.Text = ""
    cpCnpj.Text = ""
  End If
End Sub

Private Sub chTitular_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpRazaoSocial.SetFocus
  End If
End Sub

Private Sub cmdAlterar_Click()
  Botoes (False)
  vbGrava = 2
  If Not (cpCodigo.Text = "") Then
    vbBook = tbCondominio.Bookmark
  Else
    vbBook = Null
  End If
  Travar (False)
  cpNome.SetFocus
End Sub

Private Sub cmdDesfazer_Click()
  Resp = MsgBox("Cancelar as alterações?", vbQuestion + vbYesNo + vbDefaultButton2, Caption)
  If Resp = vbYes Then
    Limpar
    Botoes (True)
    If Not IsNull(vbBook) Then
      If Not IsEmpty(vbBook) Then
        With tbCondominio
          .MoveFirst
          LerDados
        End With
      End If
    Else
      With tbCondominio
        .MoveFirst
        LerDados
      End With
    End If
    Travar (True)
  End If
End Sub

Private Sub cmdExcluir_Click()
  Resp = MsgBox("Confirma a exclusão de '" + cpNome.Text + "' do cadastro?", vbQuestion + vbYesNo + vbDefaultButton2, Caption)
  If Resp = vbYes Then
    Resp = MsgBox("A exclusão deste condominio irá excluir todos os dados realacionados. Continuar a exclusão?", vbQuestion + vbYesNo + vbDefaultButton2, Caption)
    If Resp = vbYes Then
      db.Execute "Delete From Associados Where Condominio = " + cpCodigo.Text
      db.Execute "Delete From Mensalidades Where Condominio = " + cpCodigo.Text
      With tbCondominio
        .Delete
        DoEvents
        If .RecordCount > 0 Then
          .MovePrevious
          If Not .BOF Then
            LerDados
          Else
            .MoveNext
            If Not .EOF Then
              LerDados
            Else
              Limpar
            End If
          End If
        Else
          Limpar
        End If
      End With
    End If
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdIncluir_Click()
  Limpar
  cpCodigo.Text = ProximoCodigo("condominio", "codigo")
  cpNome.SetFocus
  vbGrava = 1
  Botoes (False)
  Travar (False)
End Sub

Private Sub cmdLocalizar_Click()
  RetCodigo = 0
  l_Condominio.Show 1
  If RetCodigo > 0 Then
    With tbCondominio
      .Index = "codigoid"
      .Seek "=", RetCodigo
      If Not .NoMatch Then
        LerDados
      End If
    End With
  End If
End Sub

Private Sub cmdSalvar_Click()

  Dim rsCart As Recordset
  Dim sMensagem As String
  
  If cpBanco.ListIndex < 0 Then
    MsgBox "Favor escolher um banco.", vbInformation + vbOKOnly, "Aviso"
    cpBanco.SetFocus
    Exit Sub
  End If
  
  If cpTipoCobranca.ListIndex < 0 Then
    MsgBox "Favor escolher o tipo de cobrança.", vbInformation + vbOKOnly, "Aviso"
    cpTipoCobranca.SetFocus
    Exit Sub
  End If
  
  If lbCarteira.Caption = "" Then
    sMensagem = ""
    Set rsCart = db.OpenRecordset("select * from carteira where id_bancosub = " & cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex) & ";", dbOpenDynaset)
    With rsCart
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          sMensagem = sMensagem & vbCrLf & !carteira & ": " & !Descricao
          .MoveNext
        Loop
      End If
    End With
    Set rsCart = Nothing
    If sMensagem <> "" Then
      MsgBox "Carteira inválida ou não informada!" & sMensagem, vbCritical + vbOKOnly, "Aviso"
      cpCarteira.SetFocus
      Exit Sub
    End If
  End If
  
  If Trim(cpRazaoSocial.Text) = "" Then
    MsgBox "Favor preencher o titular do titular boleto.", vbInformation + vbOKOnly, "Aviso"
    cpRazaoSocial.SetFocus
    Exit Sub
  End If
  
  If Trim(cpCnpj.Text) = "" Then
    MsgBox "Favor preencher o CNPJ do titular do boleto.", vbInformation + vbOKOnly, "Aviso"
    cpRazaoSocial.SetFocus
    Exit Sub
  End If
  
  If Trim(cpNome.Text) = "" Then
    MsgBox "Informe o nome do condomínio!", vbCritical + vbOKOnly, "Aviso"
    cpNome.SetFocus
    Exit Sub
  End If
  
  If vbGrava = 1 Then
    vbRetg = GravaDados(vbGrava)
    If vbRetg Then
      vbGrava = 0
      Botoes (True)
      Travar (True)
    End If
  Else
    vbRetg = GravaDados(vbGrava, vbBook)
    If vbRetg Then
      If Val(cpDiasRec.Text) > 20 Then
        db.Execute "update associados set acumulado = 'N' where condominio = " & cpCodigo.Text & ";"
      End If
      vbGrava = 0
      Botoes (True)
      Travar (True)
    End If
  End If
End Sub

Private Sub cpAgCedente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpConta.SetFocus
  End If
End Sub

Private Sub cpBairro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCidade.SetFocus
  End If
End Sub

Private Sub cpBanco_Click()
  If cpBanco.ListIndex >= 0 Then
    cpTipoCobranca.Clear
    Call PreencheCombo(cpTipoCobranca, "bancosub", "id", "descricao", "WHERE BANCO = " & cpBanco.ItemData(cpBanco.ListIndex))
  End If
End Sub

Private Sub cpBanco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpTipoCobranca.SetFocus
  End If
End Sub

Private Sub cpBanco_LostFocus()
  If Not cpBanco.Locked And cpBanco.ListIndex > -1 Then
    If InStr(cpBanco.Text, "SICOOB") > 0 Then
      lbCedente.Caption = "Cedente c/dig"
    Else
      lbCedente.Caption = "Cedente s/dig"
    End If
    lbCedente.Refresh
    cpCdAgencia.Text = RetornaNumeroBanco(cpBanco.ItemData(cpBanco.ListIndex))
    cpTipoCobranca.Clear
    Call PreencheCombo(cpTipoCobranca, "bancosub", "id", "descricao", "WHERE BANCO = " & cpBanco.ItemData(cpBanco.ListIndex))
  End If
End Sub

Private Sub cpCarteira_Change()
  lbCarteira.Caption = TipoCarteira(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex), cpCarteira.Value)
End Sub

Private Sub cpCarteira_KeyPress(KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpOperacao.SetFocus
  End If
End Sub

Private Sub cpCarteira_LostFocus()
  If cpTipoCobranca.ListIndex >= 0 Then
    lbCarteira.Caption = TipoCarteira(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex), cpCarteira.Text)
    If lbCarteira.Caption = "" Then
      Principal.balao.ShowBalloonTip "Carteira não informada ou inválida.", beError, "Aviso", 7000
    Else
      Principal.balao.ShowBalloonTip "", beNoSound, ".", 300
    End If
  Else
    Principal.balao.ShowBalloonTip "", beNoSound, ".", 300
  End If
End Sub

Private Sub cpCdAgencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDiaVence.SetFocus
  End If
End Sub

Private Sub cpCep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCgc.SetFocus
  End If
End Sub

Private Sub cpCgc_Change()
  If chTitular.Value = 1 Then
    cpRazaoSocial.Text = cpNome.Text
    cpCnpj.Text = cpCgc.Text
  End If
End Sub

Private Sub cpCgc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpInscricao.SetFocus
  End If
End Sub

Private Sub cpCidade_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEstado.SetFocus
  End If
End Sub

Private Sub cpCnpj_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpBanco.SetFocus
  End If
End Sub

Private Sub cpCodigo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpNome.SetFocus
  End If
End Sub

Private Sub cpConta_GotFocus()
  Principal.balao.ShowBalloonTip "Verifique com o banco '" & cpBanco.Text & "' sobre " & lbCedente.Caption & " e o tipo de convênio de cobrança!", beInformation, "Tipo de Cobrança", 3000
End Sub

Private Sub cpConta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDiasProtesto.SetFocus
  End If
End Sub

Private Sub cpConta_LostFocus()
  cpConta.Text = PadLeft(cpConta.Text, "0", 6)
  Principal.balao.ShowBalloonTip "", beNoSound, ".", 300
End Sub

Private Sub cpDiasProtesto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDiaVence.SetFocus
  End If
End Sub

Private Sub cpDiasRec_GotFocus()
  Principal.balao.ShowBalloonTip msgbal, beInformation, "Condomínio", 10000
End Sub

Private Sub cpDiasRec_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    chTitular.SetFocus
  End If
End Sub

Private Sub cpDiaVence_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDiasRec.SetFocus
  End If
End Sub

Private Sub cpEndereco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpBairro.SetFocus
  End If
End Sub

Private Sub cpEstado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCep.SetFocus
  End If
End Sub

Private Sub cpFone_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFonePortaria.SetFocus
  End If
End Sub

Private Sub cpFonePortaria_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpNome.SetFocus
  End If
End Sub

Private Sub cpFrase_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpMulta.SetFocus
  End If
End Sub

Private Sub cpInscricao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFone.SetFocus
  End If
End Sub

Private Sub cpJuros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpBanco.SetFocus
  End If
End Sub

Private Sub cpMulta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpJuros.SetFocus
  End If
End Sub

Private Sub cpNome_Change()
  If chTitular.Value = 1 Then
    cpRazaoSocial.Text = cpNome.Text
    cpCnpj.Text = cpCgc.Text
  End If
End Sub

Private Sub cpNome_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpEndereco.SetFocus
  End If
End Sub

Private Sub cpOperacao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpAgCedente.SetFocus
  End If
End Sub

Private Sub cpRazaoSocial_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCnpj.SetFocus
  End If
End Sub

Private Sub cpTipoCobranca_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpCarteira.SetFocus
  End If
End Sub

Private Sub cpTipoCobranca_LostFocus()
  If cpTipoCobranca.ListIndex >= 0 Then
    Call PreencheCarteira(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex))
    sTemp = cpConta.Text
    cpConta.Mask = PadLeft("#", "#", DigitosCedente(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex)))
    If Len(sTemp) > 0 And Len(sTemp) > cpConta.MaxLength Then
      cpConta.Text = Right(sTemp, cpConta.MaxLength)
    ElseIf Len(sTemp) > 0 And Len(sTemp) < cpConta.MaxLength Then
      cpConta.Text = PadLeft(sTemp, "0", cpConta.MaxLength)
    Else
      cpConta.Text = sTemp
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If cmdSalvar.Enabled = False Then
    If Shift = 0 Then
      Select Case KeyCode
        Case 33
          With tbCondominio
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
          With tbCondominio
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
          With tbCondominio
            If .RecordCount > 0 Then
              .MoveFirst
              LerDados
            End If
          End With
        Case 34
          With tbCondominio
            If .RecordCount > 0 Then
              .MoveLast
              LerDados
            End If
          End With
      End Select
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    cmdFechar = True
  End If
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Refresh
  KeyPreview = True
  Call PreencheCombo(cpBanco, "bancos", "codigo", "nomebanco")
  With tbCondominio
    .Index = "nomeid"
    If .RecordCount > 0 Then
      .MoveFirst
      LerDados
    End If
  End With
  Botoes (True)
  Travar (True)
  Abas.Tab = 0
End Sub

Private Sub LerDados()
  Dim iCont As Integer
  With tbCondominio
    If .RecordCount > 0 Then
      cpAgCedente.Text = IIf(IsNull(!agcedente), "", !agcedente)
      cpBairro.Text = IIf(IsNull(!bairro), "", !bairro)
      Mudou.bairro = cpBairro.Text
      For iCont = 0 To cpBanco.ListCount - 1
        If cpBanco.ItemData(iCont) = !Banco Then
          cpBanco.ListIndex = iCont
          Exit For
        End If
      Next iCont
      If InStr(cpBanco.Text, "SICOOB") > 0 Then
        lbCedente.Caption = "Cedente c/dig"
      Else
        lbCedente.Caption = "Cedente s/dig"
      End If
      lbCedente.Refresh
      cpTipoCobranca.Clear
      Call PreencheCombo(cpTipoCobranca, "bancosub", "id", "descricao", "WHERE BANCO = " & !Banco)
      For iCont = 0 To cpTipoCobranca.ListCount - 1
        If cpTipoCobranca.ItemData(iCont) = !tipoboleto Then
          cpTipoCobranca.ListIndex = iCont
          Exit For
        End If
      Next iCont
      Call PreencheCarteira(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex))
      For iCont = 0 To cpCarteira.ListCount - 1
        If !carteira & "" = cpCarteira.List(iCont, 0) Then
          cpCarteira.ListIndex = iCont
          Exit For
        End If
      Next iCont
      cpCdAgencia.Text = IIf(IsNull(!CdAgencia), "", !CdAgencia)
      cpCep.Text = IIf(IsNull(!cep), "", !cep)
      cpCgc.Text = IIf(IsNull(!CGC), "", !CGC)
      cpCidade.Text = IIf(IsNull(!Cidade), "", !Cidade)
      cpDiasProtesto.Text = !diasprotesto
      sTemp = !conta & ""
      If Len(sTemp) > 6 Then
        sTemp = Right(sTemp, 6)
      End If
      cpConta.Mask = PadLeft("#", "#", DigitosCedente(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex)))
      If Len(sTemp) > 0 And Len(sTemp) > cpConta.MaxLength Then
        cpConta.Text = Right(sTemp, cpConta.MaxLength)
      ElseIf Len(sTemp) > 0 And Len(sTemp) < cpConta.MaxLength Then
        cpConta.Text = PadLeft(sTemp, "0", cpConta.MaxLength)
      Else
        cpConta.Text = sTemp
      End If
      cpDiaVence.Text = IIf(IsNull(!Vencimento), "", !Vencimento)
      cpEndereco.Text = IIf(IsNull(!endereco), "", !endereco)
      cpEstado.Text = IIf(IsNull(!estado), "", !estado)
      cpFone.Text = IIf(IsNull(!Fone), "", !Fone)
      cpCodigo = IIf(IsNull(!Codigo), "0", !Codigo)
      cpFonePortaria.Text = IIf(IsNull(!FonePortaria), "", !FonePortaria)
      cpInscricao.Text = IIf(IsNull(!Inscricao), "", !Inscricao)
      cpNome.Text = IIf(IsNull(!Nome), "", !Nome)
      cpOperacao.Text = IIf(IsNull(!Operacao), "", !Operacao)
      If Not IsNull(!geral) Then
        Select Case !geral
          Case 1
            OpSindico.Value = True
          Case 2
            OpCotaGeral.Value = True
          Case 3
            OpCotaSindico.Value = True
          Case 4
            OpValorFixo.Value = True
          Case Else
            OpGeral.Value = True
        End Select
      End If
      cpMulta.Text = IIf(IsNull(!Multa), "", !Multa)
      cpJuros.Text = IIf(IsNull(!juros), "", !juros)
      cpFrase.Text = IIf(IsNull(!Frase), "", !Frase)
      cpDiasRec.Text = !dias & ""
      If Not IsNull(!titularboleto) Then
        chTitular.Value = !titularboleto
      Else
        chTitular.Value = 0
      End If
      If cpTipoCobranca.ListIndex >= 0 Then
        lbCarteira.Caption = TipoCarteira(cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex), cpCarteira.Text)
      End If
      cpRazaoSocial.Text = !razaoboleto & ""
      cpCnpj.Text = !cnpjboleto & ""
      Me.Caption = "Cadastro de condominio: " & !Nome
    End If
  End With
End Sub

Private Sub Limpar()
  cpAgCedente.Text = ""
  cpDiasRec.Text = 0
  cpBairro.Text = ""
  cpBanco.ListIndex = -1
  cpCarteira.ListIndex = -1
  cpCdAgencia.Text = ""
  cpCep.Text = "36570-000"
  cpCgc.Text = ""
  cpConta.Text = ""
  cpDiaVence.Text = "5"
  cpEndereco.Text = ""
  cpEstado.Text = "MG"
  cpFone.Text = ""
  cpCidade.Text = ""
  cpFonePortaria.Text = ""
  cpInscricao.Text = ""
  cpNome.Text = ""
  cpOperacao.Text = ""
  OpGeral.Value = True
  cpMulta.Text = ""
  cpJuros.Text = ""
  cpFrase.Text = ""
  cpRazaoSocial.Text = ""
  cpCnpj.Text = ""
  cpTipoCobranca.ListIndex = -1
  chTitular.Value = 1
  cpTipoCobranca.Clear
  cpCdAgencia.Text = ""
  cpDiasProtesto.Text = "0"
End Sub

Private Function GravaDados(nTipo As Integer, Optional nBook As Variant) As Boolean
  With tbCondominio
    If nTipo = 1 Then
      .AddNew
      !Codigo = cpCodigo.Text
    ElseIf nTipo = 2 Then
      .Edit
    End If
    !agcedente = cpAgCedente.Text
    !bairro = cpBairro.Text
    !Banco = cpBanco.ItemData(cpBanco.ListIndex)
    !carteira = cpCarteira.Value
    !Busca = vbSoundex.Soundex(cpNome.Text)
    !CdAgencia = cpCdAgencia.Text
    !cep = cpCep.Text
    !CGC = cpCgc.Text
    !Cidade = cpCidade.Text
    !conta = PadLeft(cpConta.Text, "0", 6)
    !Vencimento = cpDiaVence.Text
    !endereco = cpEndereco.Text
    !estado = cpEstado.Text
    !Fone = cpFone.Text
    !FonePortaria = cpFonePortaria.Text
    !Inscricao = cpInscricao.Text
    !diasprotesto = cpDiasProtesto.Text
    !Nome = cpNome.Text
    !Operacao = cpOperacao.Text
    If OpCotaSindico.Value Then
      !geral = 3
    ElseIf OpSindico.Value Then
      !geral = 1
    ElseIf OpCotaGeral.Value Then
      !geral = 2
    ElseIf OpValorFixo.Value Then
      !geral = 4
    Else
      !geral = 0
    End If
    !Multa = cpMulta.Text
    !juros = cpJuros.Text
    !Frase = cpFrase.Text
    !dias = cpDiasRec.Text
    !razaoboleto = cpRazaoSocial.Text
    !cnpjboleto = cpCnpj.Text
    !tipoboleto = cpTipoCobranca.ItemData(cpTipoCobranca.ListIndex)
    !titularboleto = chTitular.Value
    .Update
  End With
  GravaDados = True
End Function

Private Sub Botoes(Tipo As Boolean)
  cmdDesfazer.Enabled = Not Tipo
  cmdSalvar.Enabled = Not Tipo
  cmdAlterar.Enabled = Tipo
  cmdExcluir.Enabled = Tipo
  cmdFechar.Enabled = Tipo
  cmdIncluir.Enabled = Tipo
  cmdLocalizar.Enabled = Tipo
End Sub

Private Sub Travar(Tipo As Boolean)
  cpRazaoSocial.Locked = Tipo
  cpCnpj.Locked = Tipo
  cpDiasProtesto.Locked = Tipo
  cpTipoCobranca.Locked = Tipo
  chTitular.Enabled = Not Tipo
  cpAgCedente.Locked = Tipo
  cpBairro.Locked = Tipo
  cpBanco.Locked = Tipo
  cpCarteira.Locked = Tipo
  cpCep.Locked = Tipo
  cpCgc.Locked = Tipo
  cpConta.Locked = Tipo
  cpDiaVence.Locked = Tipo
  cpEndereco.Locked = Tipo
  cpEstado.Locked = Tipo
  cpFone.Locked = Tipo
  cpCidade.Locked = Tipo
  cpFonePortaria.Locked = Tipo
  cpInscricao.Locked = Tipo
  cpNome.Locked = Tipo
  cpOperacao.Locked = Tipo
  Frame2.Enabled = Not Tipo
  cpMulta.Locked = Tipo
  cpJuros.Locked = Tipo
  cpFrase.Locked = Tipo
  cpDiasRec.Locked = Tipo
  cpCdAgencia.Locked = Tipo
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set vbSoundex = Nothing
  Set tbCondominio = Nothing
End Sub

Private Sub PreencheCarteira(ByVal BancoSub As Long)
  Dim rs As Recordset
  cpCarteira.Clear
  Set rs = db.OpenRecordset("Select * from carteira where id_bancosub = " & BancoSub & " order by carteira;", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        cpCarteira.AddItem
        cpCarteira.List(cpCarteira.ListCount - 1, 0) = .Fields("carteira").Value
        cpCarteira.List(cpCarteira.ListCount - 1, 1) = .Fields("descricao").Value
        .MoveNext
      Loop
      cpCarteira.ColumnWidths = "20 pt; 100 pt"
    End If
  End With
  Set rs = Nothing
End Sub
