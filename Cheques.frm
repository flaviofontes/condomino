VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Cheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de cheques"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   Icon            =   "Cheques.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4275
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7541
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Incluir"
      TabPicture(0)   =   "Cheques.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(8)=   "Label17"
      Tab(0).Control(9)=   "cpData"
      Tab(0).Control(10)=   "cpPrePara"
      Tab(0).Control(11)=   "cpNumero"
      Tab(0).Control(12)=   "chPreDatado"
      Tab(0).Control(13)=   "cpAgencia"
      Tab(0).Control(14)=   "cpConta"
      Tab(0).Control(15)=   "cpValor"
      Tab(0).Control(16)=   "cpExtenso"
      Tab(0).Control(17)=   "cmdSalvar"
      Tab(0).Control(18)=   "List1"
      Tab(0).Control(19)=   "cpFavorecido"
      Tab(0).Control(20)=   "cpObsInc"
      Tab(0).Control(21)=   "cpCondominio"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Alterar"
      TabPicture(1)   =   "Cheques.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cpCondominioAlt"
      Tab(1).Control(1)=   "cpObsAlt"
      Tab(1).Control(2)=   "cpFavorecidoAlt"
      Tab(1).Control(3)=   "cmdSalvarAlt"
      Tab(1).Control(4)=   "cpExtensoAlt"
      Tab(1).Control(5)=   "cpValorAlt"
      Tab(1).Control(6)=   "cpContaAlt"
      Tab(1).Control(7)=   "cpAgenciaAlt"
      Tab(1).Control(8)=   "chPreDatadoAlt"
      Tab(1).Control(9)=   "cpNumeroAlt"
      Tab(1).Control(10)=   "cmdLocalizar"
      Tab(1).Control(11)=   "cpPreParaAlt"
      Tab(1).Control(12)=   "cpDataAlt"
      Tab(1).Control(13)=   "Label18"
      Tab(1).Control(14)=   "Label16"
      Tab(1).Control(15)=   "Label14"
      Tab(1).Control(16)=   "Label13"
      Tab(1).Control(17)=   "Label12"
      Tab(1).Control(18)=   "Label11"
      Tab(1).Control(19)=   "Label10"
      Tab(1).Control(20)=   "Label9"
      Tab(1).Control(21)=   "Label8"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Excluir/Imprimir"
      TabPicture(2)   =   "Cheques.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCancelar"
      Tab(2).Control(1)=   "cmdExcluir"
      Tab(2).Control(2)=   "cmdLocalizarPrn"
      Tab(2).Control(3)=   "cpNumeroPrn"
      Tab(2).Control(4)=   "chPreDatadoPrn"
      Tab(2).Control(5)=   "cpAgenciaPrn"
      Tab(2).Control(6)=   "cpContaPrn"
      Tab(2).Control(7)=   "cpValorPrn"
      Tab(2).Control(8)=   "cpExtensoPrn"
      Tab(2).Control(9)=   "cpFavorecidoPrn"
      Tab(2).Control(10)=   "cmdPrint"
      Tab(2).Control(11)=   "cpPreParaPrn"
      Tab(2).Control(12)=   "cpDataPrn"
      Tab(2).Control(13)=   "Label8Prn"
      Tab(2).Control(14)=   "Label9Prn"
      Tab(2).Control(15)=   "Label10Prn"
      Tab(2).Control(16)=   "Label11Prn"
      Tab(2).Control(17)=   "Label12Prn"
      Tab(2).Control(18)=   "Label13Prn"
      Tab(2).Control(19)=   "Label14Prn"
      Tab(2).ControlCount=   20
      TabCaption(3)   =   "Baixa"
      TabPicture(3)   =   "Cheques.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cpDataBaixa"
      Tab(3).Control(1)=   "cmdBaixar"
      Tab(3).Control(2)=   "cpDataDesc"
      Tab(3).Control(3)=   "cpFavorBaixa"
      Tab(3).Control(4)=   "cpValorBaixa"
      Tab(3).Control(5)=   "cpContaBaixa"
      Tab(3).Control(6)=   "cpAgenciaBaixa"
      Tab(3).Control(7)=   "cpNumeroBaixa"
      Tab(3).Control(8)=   "cmdLocalizarBaixa"
      Tab(3).Control(9)=   "Data"
      Tab(3).Control(10)=   "lbInfoBaixa"
      Tab(3).Control(11)=   "Label24"
      Tab(3).Control(12)=   "Line1"
      Tab(3).Control(13)=   "Label23"
      Tab(3).Control(14)=   "Label22"
      Tab(3).Control(15)=   "Label21"
      Tab(3).Control(16)=   "Label20"
      Tab(3).Control(17)=   "Label19"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "Imprimir vários"
      TabPicture(4)   =   "Cheques.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label25"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cpCodigo"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cpNome"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdLocVarios"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Data1"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdLimparFiltro"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdAplicarFiltro"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdImprimirTodos"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "opNaoImpressos"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "lsCheques"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      Begin MSComctlLib.ListView lsCheques 
         Height          =   2820
         Left            =   90
         TabIndex        =   90
         Top             =   855
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4974
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Agência"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Conta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Data"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Data Pre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Favorecido"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "N. Imp."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox opNaoImpressos 
         Caption         =   "Ainda não impresso"
         Height          =   195
         Left            =   7125
         TabIndex        =   89
         Top             =   480
         Width           =   1680
      End
      Begin VB.CommandButton cmdImprimirTodos 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   7860
         TabIndex        =   88
         Top             =   3750
         Width           =   1125
      End
      Begin VB.CommandButton cmdAplicarFiltro 
         Caption         =   "Aplicar Filtro"
         Height          =   375
         Left            =   6645
         TabIndex        =   87
         Top             =   3765
         Width           =   1125
      End
      Begin VB.CommandButton cmdLimparFiltro 
         Caption         =   "Limpar Filtro"
         Height          =   375
         Left            =   5430
         TabIndex        =   86
         Top             =   3780
         Width           =   1125
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "\\VBOXSVR\programacao\Porto real\Fontes\porto-real\dados.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1095
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3780
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdLocVarios 
         Caption         =   "..."
         Height          =   315
         Left            =   1905
         TabIndex        =   83
         Top             =   435
         Width           =   435
      End
      Begin VB.ComboBox cpCondominioAlt 
         Height          =   315
         Left            =   -74820
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3780
         Width           =   7035
      End
      Begin VB.ComboBox cpCondominio 
         Height          =   315
         Left            =   -74820
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   3840
         Width           =   7035
      End
      Begin VB.TextBox cpObsAlt 
         Height          =   855
         Left            =   -70080
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox cpObsInc 
         Height          =   855
         Left            =   -70620
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   75
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox cpFavorecido 
         Height          =   315
         Left            =   -74820
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   3180
         Width           =   5835
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -70440
         TabIndex        =   73
         Top             =   3720
         Width           =   1155
      End
      Begin rdActiveText.ActiveText cpDataBaixa 
         Height          =   315
         Left            =   -74820
         TabIndex        =   72
         Top             =   2100
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
      Begin VB.CommandButton cmdBaixar 
         Caption         =   "Baixar"
         Height          =   375
         Left            =   -71220
         TabIndex        =   69
         Top             =   3300
         Width           =   1035
      End
      Begin rdActiveText.ActiveText cpDataDesc 
         Height          =   315
         Left            =   -73560
         TabIndex        =   68
         Top             =   3300
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.TextBox cpFavorBaixa 
         Height          =   315
         Left            =   -74820
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   61
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox cpValorBaixa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -70680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   60
         Top             =   780
         Width           =   1635
      End
      Begin VB.TextBox cpContaBaixa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -72420
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   59
         Top             =   780
         Width           =   1575
      End
      Begin VB.TextBox cpAgenciaBaixa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73620
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   58
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox cpNumeroBaixa 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74820
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   57
         Top             =   780
         Width           =   1035
      End
      Begin VB.CommandButton cmdLocalizarBaixa 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   -74820
         TabIndex        =   56
         Top             =   2640
         Width           =   1155
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   -71640
         TabIndex        =   55
         Top             =   3720
         Width           =   1155
      End
      Begin VB.CommandButton cmdLocalizarPrn 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   -74040
         TabIndex        =   31
         Top             =   3720
         Width           =   1155
      End
      Begin VB.TextBox cpNumeroPrn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74820
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   22
         Top             =   660
         Width           =   1035
      End
      Begin VB.CheckBox chPreDatadoPrn 
         Caption         =   "Pre-Datado para"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -73140
         TabIndex        =   28
         Top             =   2580
         Width           =   1515
      End
      Begin VB.TextBox cpAgenciaPrn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73620
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   23
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox cpContaPrn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -72420
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   24
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox cpValorPrn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -70680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   25
         Top             =   660
         Width           =   1635
      End
      Begin VB.TextBox cpExtensoPrn 
         Height          =   855
         Left            =   -74820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox cpFavorecidoPrn 
         Height          =   315
         Left            =   -74820
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3180
         Width           =   5775
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -72840
         TabIndex        =   32
         Top             =   3720
         Width           =   1155
      End
      Begin VB.TextBox cpFavorecidoAlt 
         Height          =   315
         Left            =   -74820
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3180
         Width           =   5775
      End
      Begin VB.CommandButton cmdSalvarAlt 
         Caption         =   "Salvar"
         Height          =   315
         Left            =   -67260
         TabIndex        =   21
         Top             =   3720
         Width           =   1155
      End
      Begin VB.TextBox cpExtensoAlt 
         Height          =   855
         Left            =   -74820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   4635
      End
      Begin VB.TextBox cpValorAlt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -70680
         MaxLength       =   15
         TabIndex        =   14
         Top             =   660
         Width           =   1635
      End
      Begin VB.TextBox cpContaAlt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -72420
         MaxLength       =   12
         TabIndex        =   13
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox cpAgenciaAlt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73620
         MaxLength       =   6
         TabIndex        =   12
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox chPreDatadoAlt 
         Caption         =   "Pre-Datado para"
         Height          =   195
         Left            =   -73140
         TabIndex        =   17
         Top             =   2580
         Width           =   1515
      End
      Begin VB.TextBox cpNumeroAlt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74820
         MaxLength       =   8
         TabIndex        =   11
         Top             =   660
         Width           =   1035
      End
      Begin VB.CommandButton cmdLocalizar 
         Caption         =   "Localizar"
         Height          =   315
         Left            =   -67260
         TabIndex        =   20
         Top             =   3360
         Width           =   1155
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   -66960
         TabIndex        =   33
         Top             =   2700
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   -67200
         TabIndex        =   10
         Top             =   3720
         Width           =   1155
      End
      Begin VB.TextBox cpExtenso 
         Height          =   855
         Left            =   -74820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox cpValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -70680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   660
         Width           =   1635
      End
      Begin VB.TextBox cpConta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -72420
         MaxLength       =   12
         TabIndex        =   3
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox cpAgencia 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73620
         MaxLength       =   6
         TabIndex        =   2
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox chPreDatado 
         Caption         =   "Pre-Datado para"
         Height          =   195
         Left            =   -73140
         TabIndex        =   7
         Top             =   2580
         Width           =   1515
      End
      Begin VB.TextBox cpNumero 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74820
         MaxLength       =   8
         TabIndex        =   0
         Top             =   660
         Width           =   1035
      End
      Begin rdActiveText.ActiveText cpPrePara 
         Height          =   315
         Left            =   -71460
         TabIndex        =   8
         Top             =   2520
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
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpData 
         Height          =   315
         Left            =   -74820
         TabIndex        =   6
         Top             =   2520
         Width           =   1395
         _ExtentX        =   2461
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
      Begin rdActiveText.ActiveText cpPreParaAlt 
         Height          =   315
         Left            =   -71460
         TabIndex        =   18
         Top             =   2520
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
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpDataAlt 
         Height          =   315
         Left            =   -74820
         TabIndex        =   16
         Top             =   2520
         Width           =   1395
         _ExtentX        =   2461
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
      Begin rdActiveText.ActiveText cpPreParaPrn 
         Height          =   315
         Left            =   -71460
         TabIndex        =   29
         Top             =   2520
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
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpDataPrn 
         Height          =   315
         Left            =   -74820
         TabIndex        =   27
         Top             =   2520
         Width           =   1395
         _ExtentX        =   2461
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
         Locked          =   -1  'True
      End
      Begin rdActiveText.ActiveText cpNome 
         Height          =   315
         Left            =   2385
         TabIndex        =   84
         Top             =   435
         Width           =   4575
         _ExtentX        =   8070
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
         Left            =   1125
         TabIndex        =   85
         Top             =   435
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
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Condomínio"
         Height          =   195
         Left            =   150
         TabIndex        =   82
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Condomínio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   80
         Top             =   3540
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Condomínio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   78
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
         Height          =   195
         Left            =   -70080
         TabIndex        =   76
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Obs:"
         Height          =   195
         Left            =   -70620
         TabIndex        =   74
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Data 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -74820
         TabIndex        =   71
         Top             =   1860
         Width           =   345
      End
      Begin VB.Label lbInfoBaixa 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   -74820
         TabIndex        =   70
         Top             =   3720
         Width           =   5835
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Descontado em"
         Height          =   195
         Left            =   -74760
         TabIndex        =   67
         Top             =   3360
         Width           =   1125
      End
      Begin VB.Line Line1 
         X1              =   -74820
         X2              =   -69060
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -70680
         TabIndex        =   66
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   -72420
         TabIndex        =   65
         Top             =   540
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   -73620
         TabIndex        =   64
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -74820
         TabIndex        =   63
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Favorecido"
         Height          =   195
         Left            =   -74820
         TabIndex        =   62
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label8Prn 
         AutoSize        =   -1  'True
         Caption         =   "Favorecido"
         Height          =   195
         Left            =   -74820
         TabIndex        =   54
         Top             =   2940
         Width           =   795
      End
      Begin VB.Label Label9Prn 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -74820
         TabIndex        =   53
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label10Prn 
         AutoSize        =   -1  'True
         Caption         =   "Extenso"
         Height          =   195
         Left            =   -74820
         TabIndex        =   52
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label11Prn 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -74820
         TabIndex        =   51
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label12Prn 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   -73620
         TabIndex        =   50
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label13Prn 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   -72420
         TabIndex        =   49
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label14Prn 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -70680
         TabIndex        =   48
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -70680
         TabIndex        =   47
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   -72420
         TabIndex        =   46
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   -73620
         TabIndex        =   45
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -74820
         TabIndex        =   44
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Extenso"
         Height          =   195
         Left            =   -74820
         TabIndex        =   43
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -74820
         TabIndex        =   42
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Favorecido"
         Height          =   195
         Left            =   -74820
         TabIndex        =   41
         Top             =   2940
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -70680
         TabIndex        =   40
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   -72420
         TabIndex        =   39
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   -73620
         TabIndex        =   38
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -74820
         TabIndex        =   37
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Extenso"
         Height          =   195
         Left            =   -74820
         TabIndex        =   36
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -74820
         TabIndex        =   35
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Favorecido"
         Height          =   195
         Left            =   -74820
         TabIndex        =   34
         Top             =   2940
         Width           =   795
      End
   End
End
Attribute VB_Name = "Cheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim nConta As String
Dim nTecla As Integer
Dim rsCheque As Recordset
Dim rsChequeAlt As Recordset
Dim tbUsuarios As Recordset
Dim tbCondominio As Recordset
Dim i As Integer
Dim sNumero As String, sAgencia As String, sConta As String
Dim cdFornec As Long
Dim lpt As Boolean

Private Sub chPreDatado_Click()
  cpPrePara.Text = ""
  If chPreDatado.Value = 1 Then
    cpPrePara.Locked = False
    cpPrePara.SetFocus
  Else
    cpPrePara.Locked = True
  End If
End Sub

Private Sub chPreDatadoAlt_Click()
  cpPreParaAlt.Text = ""
  If chPreDatadoAlt.Value = 1 Then
    cpPreParaAlt.Locked = False
    cpPreParaAlt.SetFocus
  Else
    cpPreParaAlt.Locked = True
  End If
End Sub

Private Sub cmdAplicarFiltro_Click()
  If Val(cpCodigo.Text) > 0 Then
    If opNaoImpressos.Value = 1 Then
      Data1.RecordSource = "select * from cheques where condominio = " _
          & cpCodigo.Text & " and cancelado = false and baixado = false and nimpressoes < 1 order by numero;"
      Data1.Refresh
    Else
      Data1.RecordSource = "select * from cheques where condominio = " _
          & cpCodigo.Text & " and cancelado = false and baixado = false order by numero;"
      Data1.Refresh
    End If
  Else
    If opNaoImpressos.Value = 1 Then
      Data1.RecordSource = "select * from cheques where cancelado = false and baixado = false and nimpressoes < 1 order by numero;"
      Data1.Refresh
    Else
      Data1.RecordSource = "select * from cheques where cancelado = false and baixado = false order by numero;"
      Data1.Refresh
    End If
  End If
  
  lsCheques.ListItems.Clear
  Dim nItem As MSComctlLib.ListItem
  
  With Data1.Recordset
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        Set nItem = lsCheques.ListItems.Add(, "C" & !Numero & !Conta)
        nItem.Text = !Numero
        nItem.SubItems(1) = !agencia
        nItem.SubItems(2) = !Conta
        nItem.SubItems(3) = Format$(!Data, "dd/MM/yyyy")
        nItem.SubItems(4) = Format$(!Datapre, "dd/MM/yyyy")
        nItem.SubItems(5) = !favorecido
        nItem.SubItems(6) = Format$(!valor, "#,##0.00")
        nItem.SubItems(7) = !nimpressoes
        .MoveNext
      Loop
    End If
  End With
  
End Sub

Private Sub cmdBaixar_Click()
  If cpNumeroBaixa.Text = "" Then
    MsgBox "Localize o cheque primeiro.", vbInformation + vbOKOnly, "Aviso"
    cmdLocalizarBaixa.SetFocus
    Exit Sub
  End If
  If rsCheque!baixado Then
    MsgBox "Cheque já baixado.", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  If Not IsDate(cpDataDesc.Text) Then
    MsgBox "Informe a data da baixa.", vbInformation + vbOKOnly, "Aviso"
    cpDataDesc.SetFocus
    Exit Sub
  End If
  If CDate(cpDataDesc.Text) < CDate(cpDataBaixa.Text) Then
    MsgBox "A data da baixa não pode ser menor que a data do cheque.", vbInformation + vbOKOnly, "Aviso"
    cpDataDesc.SetFocus
    Exit Sub
  End If
  Resp = MsgBox("Confirma que o cheque '" & cpNumeroBaixa.Text & "' foi descontado em '" & Format$(cpDataDesc.Text, "dd/MM/yyyy") & "'?", vbQuestion + vbYesNo, "Baixa")
  If Resp = vbYes Then
    
    rsCheque.MoveFirst
    rsCheque.Edit
    rsCheque!datadesconto = cpDataDesc.Text
    rsCheque!baixado = True
    rsCheque.Update
    
    lbInfoBaixa.Caption = "1 registro alterado." & vbCrLf & _
    "Último cheque baixado: " & cpNumeroBaixa.Text
    
    LimpaBaixa
  End If
End Sub

Private Sub cmdCancelar_Click()
  Resp = MsgBox("Cancelar este cheque?", vbQuestion + vbYesNo, "Cancelar")
  If Resp = vbYes Then
    rsCheque.MoveFirst
    rsCheque.Edit
    rsCheque!CANCELADO = True
    rsCheque.Update
    LimparPrn
  End If
End Sub

Private Sub cmdExcluir_Click()
  Dim rsSaldo As Recordset
  Dim sdlValor As Double
  Dim sSinal As String
  Dim nData As Date
  
  Resp = MsgBox("Excluir este cheque?", vbQuestion + vbYesNo, "Excluir")
  If Resp = vbYes Then
    If cpNumeroPrn.Text <> "" Then
      With rsCheque
        If !impresso Then
          Dim sUsuario As String
          Dim sTexto As String
          
Dinovo:
          Supervisor.Caption = "Senha do usuário - Cheuqe já impresso"
          Supervisor.Show vbModal
          If sSuper <> "" Then
            If sSuper <> "meunome" Then
              With tbUsuarios
                .MoveFirst
                Do While Not .EOF
                  If sSuper = Decodifica(!Senha) Then
                    sUsuario = Decodifica(!Nome)
                    Exit Do
                  End If
                  .MoveNext
                Loop
                If .EOF Then
                  MsgBox "Esta não é a senha correta!", vbCritical + vbOKOnly, "Aviso"
                  GoTo Dinovo
                End If
              End With
            Else
              sUsuario = "MESTRE"
            End If
          Else
            Exit Sub
          End If
        End If
      End With
      
      db.Execute "delete from cheques where numero = '" & cpNumeroPrn.Text & "' and agencia = '" & _
        cpAgenciaPrn.Text & "' and conta = '" & cpContaPrn.Text & "' and valor = " & Replace(cpValorPrn.Text, ",", ".") & ";"
      CriarArquivoLog "Eclusão de cheque número '" & cpNumeroPrn.Text & "' para " & cpFavorecidoPrn.Text, sUsuario
      LimparPrn
    End If
  End If
End Sub

Private Sub cmdImprimirTodos_Click()
  
  Dim iCounter As Integer
  Dim iFree As Integer
  Dim sPrnNome As String
  Dim Item As MSComctlLib.ListItem
  
  iCounter = 0
  For Each Item In lsCheques.ListItems
    If Item.Checked = True Then
      iCounter = iCounter + 1
    End If
  Next
  
  If iCounter = 0 Then
    MsgBox "Nenhum item foi selecionado para impressão!", vbInformation + vbOKOnly, "Aviso"
    Exit Sub
  End If
  
  sPrnNome = Printer.DeviceName
  Set Extensos = List1
  
  If Left$(sPrnNome, 2) = "\\" Then
    If lpt = False Then
      Shell "net use LPT2: /DEL", vbHide
      DoEvents
      Shell "net use LPT2: """ & sPrnNome & """", vbHide
      DoEvents
      lpt = True
    End If
  End If
  
  Resp = MsgBox("Verifique a impressora '" & sPrnNome & "' e clique em ok.", vbInformation + vbOKCancel, "Imprimir")
  If Resp = vbOK Then
    iFree = FreeFile
    If Left$(sPrnNome, 2) = "\\" Then
      Open "lpt2" For Output As #iFree
    Else
      If Left$(Printer.Port, 3) = "LPT" Then
        Open Printer.Port For Output As #iFree
      Else
        MsgBox "A impressora '" & sPrnNome & "' não pode ser utilizada."
        Exit Sub
      End If
    End If
    
    Dim sRetorno As String

    If (lsCheques.ListItems.Count > 0) Then
      Print #iFree, Chr(27) & Chr(70);
      For Each Item In lsCheques.ListItems
        If Item.Checked = True Then
          Print #iFree, Tab(68 - Len(Item.SubItems(6))); Format(Item.SubItems(6), "#0.00")
          Print #iFree, ""
          Print #iFree, ""
          sRetorno = Space(512)
          Call extenso(Item.SubItems(6), sRetorno)
          DividirExtenso AcertaLetras(sRetorno), 55
          If Extensos.ListCount > 1 Then
            Print #iFree, Tab(20); "(" & Extensos.List(0)
            Print #iFree, Tab(10); Extensos.List(1) & ")"
          Else
            Print #iFree, Tab(20); "(" & Extensos.List(0) & ")"
            Print #iFree, ""
          End If
          Print #iFree, ""
          Print #iFree, Tab(10); AcertaLetras(Trim(Item.SubItems(5)))
          Print #iFree, ""
          Print #iFree, Tab(39); AcertaLetras(vbEmpresa.Cidade) & ",";
          Print #iFree, Tab(46); Day(Item.SubItems(3));
          Print #iFree, Tab(57); NomeDoMes(Month(Item.SubItems(3)));
          Print #iFree, Tab(73); Right(Year(Item.SubItems(3)), 2)
          Print #iFree, ""
          Print #iFree, ""
          Print #iFree, ""
          Print #iFree, ""
          Print #iFree, ""
          If IsDate(Item.SubItems(4)) Then
            Print #iFree, Tab(50); "BOM P/ " & Item.SubItems(4)
          Else
            Print #iFree, ""
          End If
          Print #iFree, ""
          Print #iFree, ""
          Print #iFree, ""
          db.Execute "update cheques set impresso = True, nimpressoes = " & (Val(Item.SubItems(7)) + 1) _
                     & " where numero = '" & Item.Text & "' and conta = '" & Item.SubItems(2) & "';"
        End If
        DoEvents
      Next
    End If
    Close #iFree
  End If
  cmdAplicarFiltro_Click
End Sub

Private Sub cmdLimparFiltro_Click()
  cpCodigo.Text = 0
  cpNome.Text = ""
  cmdAplicarFiltro_Click
End Sub

Private Sub cmdLocalizar_Click()
  Dim rsSaldo As Recordset
  Dim sdlValor As Double
  Dim sSinal As String
  Dim nData As Date
  Dim i As Integer
  
  LocCheques.Show vbModal
  If sLocCheque(0) <> "" Then
    Set rsCheque = db.OpenRecordset("select * from cheques where numero = '" & sLocCheque(0) & "' and agencia = '" & _
      sLocCheque(1) & "' and conta = '" & sLocCheque(2) & "';", dbOpenDynaset)
    With rsCheque
      If .RecordCount > 0 Then
        Limpar
        .MoveFirst
        If !impresso Then
          If Not !baixado Then
            Resp = MsgBox("Cheque '" & sLocCheque(0) & "' já impresso não pode ser alterado! Gostaria de cancelar?", vbExclamation + vbYesNo, "Aviso")
            If Resp = vbYes Then
              .Edit
              !CANCELADO = True
              .Update
              MsgBox "Cheque '" & sLocCheque(0) & "' cancelado.", vbInformation + vbOKOnly, "Aviso"
            End If
          Else
pBaixado:
            MsgBox "Cheque '" & sLocCheque(0) & "' baixado, somente a conta do financeiro pode ser alterada!", vbExclamation + vbOKOnly, "Aviso"
            sNumero = !Numero
            sAgencia = !agencia
            sConta = !Conta
            
            chPreDatadoAlt.Enabled = False
            cpNumeroAlt.Locked = True
            cpPreParaAlt.Enabled = False
            cpAgenciaAlt.Locked = True
            cpContaAlt.Locked = True
            cpValorAlt.Locked = True
            cpExtensoAlt.Locked = True
            cpDataAlt.Locked = True
            cpFavorecidoAlt.Locked = True
            
            cpNumeroAlt.Text = !Numero
            cpAgenciaAlt.Text = !agencia
            cpContaAlt.Text = !Conta
            cpValorAlt.Text = Format$(!valor, "#0.00")
            cpExtensoAlt.Text = !extenso
            cpFavorecidoAlt.Text = !favorecido
            cpDataAlt.Text = !Data
            cpObsAlt.Text = !obs & ""
            For i = 0 To cpCondominioAlt.ListCount - 1
              If cpCondominioAlt.ItemData(i) = !Condominio Then
                cpCondominioAlt.ListIndex = i
                Exit For
              End If
            Next i
            If !predatado Then
              chPreDatadoAlt.Value = 1
              cpPreParaAlt.Text = !Datapre
            Else
              chPreDatadoAlt.Value = 0
              cpPreParaAlt.Text = ""
            End If
          End If
        Else
          If !CANCELADO Then
            MsgBox "Cheque '" & sLocCheque(0) & "' cancelado, não pode ser alterado!", vbExclamation + vbOKOnly, "Aviso"
          Else
            If !baixado Then
              GoTo pBaixado
            Else
              sNumero = !Numero
              sAgencia = !agencia
              sConta = !Conta
              
              chPreDatadoAlt.Enabled = True
              cpNumeroAlt.Locked = False
              cpPreParaAlt.Enabled = True
              cpAgenciaAlt.Locked = False
              cpContaAlt.Locked = False
              cpValorAlt.Locked = False
              cpExtensoAlt.Locked = False
              cpDataAlt.Locked = False
              cpFavorecidoAlt.Locked = False
              
              cpNumeroAlt.Text = !Numero
              cpAgenciaAlt.Text = !agencia
              cpContaAlt.Text = !Conta
              cpValorAlt.Text = Format$(!valor, "#0.00")
              cpExtensoAlt.Text = !extenso
              cpFavorecidoAlt.Text = !favorecido
              cpObsAlt.Text = !obs & ""
              cpDataAlt.Text = !Data
              For i = 0 To cpCondominioAlt.ListCount - 1
                If cpCondominioAlt.ItemData(i) = !Condominio Then
                  cpCondominioAlt.ListIndex = i
                  Exit For
                End If
              Next i
              If !predatado Then
                chPreDatadoAlt.Value = 1
                cpPreParaAlt.Text = !Datapre
              Else
                chPreDatadoAlt.Value = 0
                cpPreParaAlt.Text = ""
              End If
            End If
          End If
        End If
      End If
    End With
  End If
End Sub

Private Sub LimpaBaixa()
  cpNumeroBaixa.Text = ""
  cpAgenciaBaixa.Text = ""
  cpContaBaixa.Text = ""
  cpValorBaixa.Text = 0
  cpFavorBaixa.Text = ""
  cpDataBaixa.Text = ""
End Sub

Private Sub cmdLocalizarBaixa_Click()
  LimpaBaixa
  LocCheques.Show vbModal
  If sLocCheque(0) <> "" Then
    Set rsCheque = db.OpenRecordset("select * from cheques where numero = '" & sLocCheque(0) & "' and agencia = '" & _
      sLocCheque(1) & "' and conta = '" & sLocCheque(2) & "';", dbOpenDynaset)
    With rsCheque
      If .RecordCount > 0 Then
        LimparPrn
        If !CANCELADO Then
          MsgBox "Cheque '" & sLocCheque(0) & "' cancelado, não pode ser baixado!", vbExclamation + vbOKOnly, "Aviso"
        ElseIf Not !impresso Then
          MsgBox "Cheque '" & sLocCheque(0) & "' ainda não impresso, não pode ser baixado!", vbExclamation + vbOKOnly, "Aviso"
        ElseIf !baixado Then
          MsgBox "Cheque '" & sLocCheque(0) & "' já baixado!", vbExclamation + vbOKOnly, "Aviso"
        Else
          cpNumeroBaixa.Text = !Numero
          cpAgenciaBaixa.Text = !agencia
          cpContaBaixa.Text = !Conta
          cpValorBaixa.Text = Format$(!valor, "#0.00")
          cpFavorBaixa.Text = !favorecido
          cpDataBaixa.Text = !Data
        End If
        cpDataDesc.SetFocus
      End If
    End With
  End If
End Sub

Private Sub cmdLocalizarPrn_Click()
  Dim i As Integer
  LocCheques.Show vbModal
  If sLocCheque(0) <> "" Then
    Set rsCheque = db.OpenRecordset("select * from cheques where numero = '" & sLocCheque(0) & "' and agencia = '" & _
      sLocCheque(1) & "' and conta = '" & sLocCheque(2) & "' and valor = " & Replace(sLocCheque(3), ",", ".") & ";", dbOpenDynaset)
    With rsCheque
      If .RecordCount > 0 Then
        LimparPrn
        If !baixado Then
          MsgBox "Cheque '" & sLocCheque(0) & "' baixado, não pode ser impresso/excluido!", vbExclamation + vbOKOnly, "Aviso"
        Else
          If !CANCELADO Then
            cmdPrint.Enabled = False
            cmdCancelar.Enabled = False
          Else
            cmdPrint.Enabled = True
            cmdCancelar.Enabled = True
          End If
          cpNumeroPrn.Text = !Numero
          cpAgenciaPrn.Text = !agencia
          cpContaPrn.Text = !Conta
          cpValorPrn.Text = Format$(!valor, "#0.00")
          cpExtensoPrn.Text = !extenso
          cpFavorecidoPrn.Text = !favorecido
          cpDataPrn.Text = !Data
          If !predatado Then
            chPreDatadoPrn.Value = 1
            cpPreParaPrn.Text = !Datapre
          Else
            chPreDatadoPrn.Value = 0
            cpPreParaPrn.Text = ""
          End If
        End If
      End If
    End With
  End If
End Sub

Private Sub cmdLocVarios_Click()
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
  
  Dim iCounter As Integer
  Dim iFree As Integer
  Dim sPrnNome As String
  
  If cpNumeroPrn.Text = Empty Then
    MsgBox "Informe o número do cheque.", vbCritical + vbOKOnly, "Aviso"
    cpNumeroPrn.SetFocus
    GoTo Sair
  End If
  
  If cpValorPrn.Text = Empty Then
    MsgBox "Informe o valor do cheque.", vbCritical + vbOKOnly, "Aviso"
    cpValorPrn.SetFocus
    GoTo Sair
  Else
    If CDbl(cpValorPrn.Text) <= 0 Then
      MsgBox "Informe o valor positivo.", vbCritical + vbOKOnly, "Aviso"
      cpValorPrn.SetFocus
      GoTo Sair
    End If
  End If
  If Not IsDate(cpDataPrn.Text) Then
    MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
    cpDataPrn.SetFocus
    GoTo Sair
  End If
  
  If chPreDatadoPrn.Value = 1 Then
    If Not IsDate(cpPreParaPrn.Text) Then
      MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
      cpPreParaPrn.SetFocus
      GoTo Sair
    End If
  End If
  
  sPrnNome = Printer.DeviceName
  Set Extensos = List1
  
  DividirExtenso AcertaLetras(cpExtensoPrn.Text), 55
  
  If Left$(sPrnNome, 2) = "\\" Then
    If lpt = False Then
      Shell "net use LPT2: /DEL", vbHide
      DoEvents
      Shell "net use LPT2: """ & sPrnNome & """", vbHide
      DoEvents
      lpt = True
    End If
  End If
  
  Resp = MsgBox("Verifique a impressora '" & sPrnNome & "' e clique em ok.", vbInformation + vbOKCancel, "Imprimir")
  If Resp = vbOK Then
    iFree = FreeFile
    If Left$(sPrnNome, 2) = "\\" Then
      Open "lpt2" For Output As #iFree
    Else
      If Left$(Printer.Port, 3) = "LPT" Then
        Open Printer.Port For Output As #iFree
      Else
        MsgBox "A impressora '" & sPrnNome & "' não pode ser utilizada."
        Exit Sub
      End If
    End If
    Print #iFree, Chr(27) & Chr(70); Tab(68 - Len(cpValorPrn.Text)); cpValorPrn.Text
    Print #iFree, ""
    Print #iFree, ""
    If Extensos.ListCount > 1 Then
      Print #iFree, Tab(20); "(" & Extensos.List(0)
      Print #iFree, Tab(10); Extensos.List(1) & ")"
    Else
      Print #iFree, Tab(20); "(" & Extensos.List(0) & ")"
      Print #iFree, ""
    End If
    Print #iFree, ""
    Print #iFree, Tab(10); AcertaLetras(cpFavorecidoPrn.Text)
    Print #iFree, ""
    Print #iFree, Tab(39); AcertaLetras(vbEmpresa.Cidade) & ",";
    Print #iFree, Tab(46); Day(cpDataPrn.Text);
    Print #iFree, Tab(57); NomeDoMes(Month(cpDataPrn.Text));
    Print #iFree, Tab(73); Right(Year(cpDataPrn.Text), 2)
    Print #iFree, ""
    Print #iFree, ""
    Print #iFree, ""
    Print #iFree, ""
    Print #iFree, ""
    If chPreDatadoPrn.Value = 1 Then
      Print #iFree, Tab(50); "BOM P/ " & cpPreParaPrn.Text
    Else
      Print #iFree, ""
    End If
    Print #iFree, ""
    Print #iFree, ""
    Print #iFree, ""
    Close #iFree
  End If
  rsCheque.Edit
  rsCheque!impresso = True
  If Not IsNull(rsCheque!nimpressoes) Then
    rsCheque!nimpressoes = rsCheque!nimpressoes + 1
  Else
    rsCheque!nimpressoes = 1
  End If
  rsCheque.Update
Sair:
  Exit Sub
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Errado
  
  Dim nValor As String
  
  If cpNumero.Text = Empty Then
    MsgBox "Informe o número do cheque.", vbCritical + vbOKOnly, "Aviso"
    cpNumero.SetFocus
    GoTo Sair
  End If
  If cpCondominio.ListIndex < 0 Then
    MsgBox "Selecione um condomínio.", vbCritical + vbOKOnly, "Aviso"
    cpCondominio.SetFocus
    GoTo Sair
  End If
  
  If cpAgencia.Text = Empty Then
    cpAgencia.Text = "-"
  End If

  If cpConta.Text = Empty Then
    cpConta.Text = "-"
  End If
  
  If cpValor.Text = Empty Then
    MsgBox "Informe o valor do cheque.", vbCritical + vbOKOnly, "Aviso"
    cpValor.SetFocus
    GoTo Sair
  Else
    If CDbl(cpValor.Text) <= 0 Then
      MsgBox "Informe o valor positivo.", vbCritical + vbOKOnly, "Aviso"
      cpValor.SetFocus
      GoTo Sair
    End If
  End If
  If Not IsDate(cpData.Text) Then
    MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
    cpData.SetFocus
    GoTo Sair
  End If
  If chPreDatado.Value = 1 Then
    If cpData.Text = Empty Then
      MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
      cpData.SetFocus
      GoTo Sair
    End If
  End If
  
  Set rsCheque = db.OpenRecordset("select * from cheques where numero = '" & cpNumero.Text & "' and agencia = '" & _
    cpAgencia.Text & "' and conta = '" & cpConta.Text & "' and cancelado = false;", dbOpenDynaset)
  
  If rsCheque.RecordCount > 0 Then
    MsgBox "Já existe um cheque lançado com mesmo número para esta conta e agência.", vbCritical + vbOKOnly, "Aviso"
    cpNumero.SetFocus
    GoTo Sair
  End If
  Set rsCheque = Nothing
  If Trim(cpObsInc.Text) = "" Then
    cpObsInc.Text = "."
  End If
  If Trim(cpFavorecido.Text) = "" Then
    cpFavorecido.Text = "."
  End If
  nValor = Replace(cpValor.Text, ",", ".")
  If (chPreDatado.Value = 1) Then
    db.Execute "insert into cheques(numero, agencia, conta, valor, extenso, favorecido, data, predatado, datapre, plconta, cdfornecedor, obs, condominio) values ('" & _
        cpNumero.Text & "', '" & cpAgencia.Text & "', '" & cpConta.Text & "', " & nValor & ", '" & cpExtenso.Text & _
        "', '" & cpFavorecido.Text & "', #" & Format$(cpData.Text, "mm/dd/yyyy") & "#, True, #" & Format$(cpPrePara.Text, "mm/dd/yyyy") & "#, '" & _
        "000000" & "', " & cdFornec & ", '" & cpObsInc.Text & "', " & cpCondominio.ItemData(cpCondominio.ListIndex) & ");"
  Else
    db.Execute "insert into cheques(numero, agencia, conta, valor, extenso, favorecido, data, predatado, datapre, plconta, cdfornecedor, obs, condominio) values ('" & _
        cpNumero.Text & "', '" & cpAgencia.Text & "', '" & cpConta.Text & "', " & nValor & ", '" & cpExtenso.Text & _
        "', '" & cpFavorecido.Text & "', #" & Format$(cpData.Text, "mm/dd/yyyy") & "#, false, null, '" & _
        "000000" & "', " & cdFornec & ", '" & cpObsInc.Text & "', " & cpCondominio.ItemData(cpCondominio.ListIndex) & ");"
  End If
  If db.RecordsAffected Then
    Limpar
    cpNumero.SetFocus
  Else
    MsgBox "Erro ao gravar o registro.", vbCritical + vbOKOnly, "Erro"
  End If
Sair:
  Exit Sub
Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Mensagem " & Err.Description & vbCrLf & "Origem " & Err.Source, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Sub

Private Sub cmdSalvarAlt_Click()
On Error GoTo Errado

  Dim sContaOld As String
  
  If cpCondominioAlt.ListIndex = -1 Then
    MsgBox "Selecione um condominio.", vbCritical + vbOKOnly, "Aviso"
    cpCondominioAlt.SetFocus
    GoTo Sair
  End If
  If cpNumeroAlt.Text = Empty Then
    MsgBox "Informe o número do cheque.", vbCritical + vbOKOnly, "Aviso"
    cpNumeroAlt.SetFocus
    GoTo Sair
  End If
  
  If cpAgenciaAlt.Text = Empty Then
    cpAgenciaAlt.Text = "-"
  End If
  If cpContaAlt.Text = Empty Then
    cpContaAlt.Text = "-"
  End If
  
  If cpValorAlt.Text = Empty Then
    MsgBox "Informe o valor do cheque.", vbCritical + vbOKOnly, "Aviso"
    cpValorAlt.SetFocus
    GoTo Sair
  Else
    If CDbl(cpValorAlt.Text) <= 0 Then
      MsgBox "Informe o valor positivo.", vbCritical + vbOKOnly, "Aviso"
      cpValorAlt.SetFocus
      GoTo Sair
    End If
  End If
  If Not IsDate(cpDataAlt.Text) Then
    MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
    cpDataAlt.SetFocus
    GoTo Sair
  End If
  If chPreDatadoAlt.Value = 1 Then
    If cpDataAlt.Text = Empty Then
      MsgBox "Informe uma data válida.", vbCritical + vbOKOnly, "Aviso"
      cpDataAlt.SetFocus
      GoTo Sair
    End If
  End If
  
  If (sNumero <> cpNumeroAlt.Text) Or (sAgencia <> cpAgenciaAlt.Text) Or (sConta <> cpContaAlt.Text) Then
    Set rsChequeAlt = db.OpenRecordset("select * from cheques where numero = '" & cpNumeroAlt.Text & "' and agencia = '" & _
      cpAgenciaAlt.Text & "' and conta = '" & cpContaAlt.Text & "';", dbOpenDynaset)
    If rsCheque.RecordCount > 0 Then
      MsgBox "Já existe um cheque lançado com mesmo número para esta conta e agência.", vbCritical + vbOKOnly, "Aviso"
      cpNumeroAlt.SetFocus
      GoTo Sair
    End If
  End If
  If Trim(cpObsAlt.Text) = "" Then
    cpObsAlt.Text = "."
  End If
  If Trim(cpFavorecidoAlt.Text) = "" Then
    cpFavorecidoAlt.Text = "."
  End If
  
  With rsCheque
    .Edit
    !Numero = cpNumeroAlt.Text
    !agencia = cpAgenciaAlt.Text
    !Conta = cpContaAlt.Text
    !valor = cpValorAlt.Text
    !extenso = cpExtensoAlt.Text
    !favorecido = cpFavorecidoAlt.Text
    !Data = cpDataAlt.Text
    !cdfornecedor = cdFornec
    !plconta = "000000"
    !Condominio = cpCondominioAlt.ItemData(cpCondominioAlt.ListIndex)
    !obs = cpObsAlt.Text
    If (chPreDatadoAlt.Value = 1) Then
      !predatado = True
      !Datapre = cpPreParaAlt.Text
    Else
      !predatado = False
      !Datapre = Null
    End If
    .Update
  End With
  
  MsgBox "Alterações salvas com sucesso.", vbInformation + vbOKOnly, "Aviso"
  LimparAlt
  cpNumero.SetFocus
Sair:
  Set rsChequeAlt = Nothing
  Exit Sub
  
Errado:
  MsgBox "Erro n. " & Err.Number & vbCrLf & "Mensagem " & Err.Description & vbCrLf & "Origem " & Err.Source, vbCritical + vbOKOnly, "Erro"
  Resume Sair
End Sub

Private Sub cpAgencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpConta.SetFocus
  Else
    KeyAscii = vMask(KeyAscii, cpAgencia)
  End If
End Sub

Private Sub cpAgenciaAlt_GotFocus()
  cpAgenciaAlt.SelStart = Len(cpAgenciaAlt.Text)
End Sub

Private Sub cpAgenciaAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpContaAlt.SetFocus
  Else
    KeyAscii = vMask(KeyAscii, cpAgenciaAlt)
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

Private Sub cpConta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValor.SetFocus
  Else
    KeyAscii = vMask(KeyAscii, cpConta)
  End If
End Sub

Private Sub cpContaAlt_GotFocus()
  cpContaAlt.SelStart = Len(cpContaAlt.Text)
End Sub

Private Sub cpContaAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpValorAlt.SetFocus
  Else
    KeyAscii = vMask(KeyAscii, cpContaAlt)
  End If
End Sub

Private Sub cpData_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If chPreDatado.Value = 1 Then
      cpData.SetFocus
    Else
      cpFavorecido.SetFocus
    End If
  End If
End Sub

Private Sub cpDataAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    If chPreDatadoAlt.Value = 1 Then
      cpDataAlt.SetFocus
    Else
      cpFavorecidoAlt.SetFocus
    End If
  End If
End Sub

Private Sub cpDataDesc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdBaixar.SetFocus
  End If
End Sub

Private Sub cpExtenso_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpData.SetFocus
  End If
End Sub

Private Sub cpExtensoAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDataAlt.SetFocus
  End If
End Sub

Private Sub cpFavorecido_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSalvar.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpFavorecidoAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSalvarAlt.SetFocus
  Else
    KeyAscii = vTexto(KeyAscii)
  End If
End Sub

Private Sub cpNumero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpAgencia.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpNumero_LostFocus()
  If Len(cpNumero.Text) < 6 Then
    cpNumero.Text = Format$(cpNumero.Text, "000000")
  End If
End Sub

Private Sub cpNumeroAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpAgenciaAlt.SetFocus
  Else
    KeyAscii = vNumero(KeyAscii)
  End If
End Sub

Private Sub cpNumeroAlt_LostFocus()
  If Len(cpNumeroAlt.Text) < 6 Then
    cpNumeroAlt.Text = Format$(cpNumeroAlt.Text, "000000")
  End If
End Sub

Private Sub cpPreParaAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFavorecidoAlt.SetFocus
  End If
End Sub

Private Sub cpValorAlt_GotFocus()
  cpValorAlt.SelStart = Len(cpValorAlt.Text)
End Sub

Private Sub cpValorAlt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpDataAlt.SetFocus
  Else
    KeyAscii = vMaskValor(KeyAscii, cpValorAlt)
  End If
End Sub

Private Sub cpPrePara_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpFavorecido.SetFocus
  End If
End Sub

Private Sub cpValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cpData.SetFocus
  Else
    KeyAscii = vMaskValor(KeyAscii, cpValor)
  End If
End Sub

Private Sub cpValor_LostFocus()
  Dim sRetorno As String
  sRetorno = Space(512)
  If cpValor.Text <> Empty Then
    Call extenso(cpValor.Text, sRetorno)
    cpExtenso.Text = Trim$(UCase(sRetorno))
  Else
    cpExtenso.Text = Empty
  End If
End Sub

Private Sub cpValorAlt_LostFocus()
  If cpValorAlt.Text <> Empty Then
    cpExtensoAlt.Text = MinhaDll.NumeroPorExtenso(cpValorAlt.Text, , , sdUcase)
  Else
    cpExtensoAlt.Text = Empty
  End If
End Sub

Private Sub Limpar()
  cpNumero.Text = ""
  cpCondominio.ListIndex = -1
  cpObsInc.Text = ""
  cpValor.Text = ""
  cpExtenso.Text = ""
  cpFavorecido.Text = ""
  cpData.Text = ""
  chPreDatado.Value = 0
  cpPrePara.Text = ""
  cdFornec = -1
End Sub

Private Sub LimparAlt()
  cpNumeroAlt.Text = ""
  cpCondominioAlt.ListIndex = -1
  cpAgenciaAlt.Text = ""
  cpContaAlt.Text = ""
  cpValorAlt.Text = ""
  cpExtensoAlt.Text = ""
  cpFavorecidoAlt.Text = ""
  cpDataAlt.Text = ""
  chPreDatadoAlt.Value = 0
  cpPreParaAlt.Text = ""
  cpObsAlt.Text = ""
  cdFornec = -1
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Call PreencheCombo(cpCondominio, "condominio", "codigo", "nome")
  Call PreencheCombo(cpCondominioAlt, "condominio", "codigo", "nome")
  Set tbCondominio = db.OpenRecordset("CONDOMINIO", dbOpenTable)
  Data1.DatabaseName = Parametros.dados
  Set tbUsuarios = db.OpenRecordset("Usuarios", dbOpenTable)
  SSTab1.Tab = 0
  lpt = False
End Sub

Private Sub LimparPrn()
  cpNumeroPrn.Text = ""
  cpAgenciaPrn.Text = ""
  cpContaPrn.Text = ""
  cpValorPrn.Text = ""
  cpExtensoPrn.Text = ""
  cpFavorecidoPrn.Text = ""
  cpDataPrn.Text = ""
  chPreDatadoPrn.Value = 0
  cpPreParaPrn.Text = ""
End Sub

Private Function DividirExtenso(ByVal StrExtenso As String, Optional ByVal nMaximo As Integer = 45)
  Dim nPos  As Integer
  Dim resto As String
  Extensos.Clear
  If StrExtenso = "" Or nMaximo = 0 Then
    Exit Function
  End If
  resto = StrExtenso
  Do While True
    resto = Trim(resto)
    If Len(resto) > (nMaximo) Then
      nPos = InStrRev(resto, " ", nMaximo)
      Extensos.AddItem Left$(resto, nPos - 1)
      If nPos > 0 Then
        resto = Mid$(resto, nPos + 1)
      Else
        resto = ""
      End If
    Else
      Extensos.AddItem resto
      Exit Do
    End If
  Loop
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set tbUsuarios = Nothing
  Set tbCondominio = Nothing
  Set rsCheque = Nothing
  Set rsChequeAlt = Nothing
End Sub
