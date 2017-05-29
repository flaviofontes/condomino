VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form NivelDoUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração dos Níveis de Usuários"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   Icon            =   "NivelDoUsuario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   7500
      TabIndex        =   47
      ToolTipText     =   "Fecha sem salvar as alterações."
      Top             =   4200
      Width           =   1410
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   315
      Left            =   5910
      TabIndex        =   46
      ToolTipText     =   "Salva as configurações."
      Top             =   4200
      Width           =   1410
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4065
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Menus do Sistema"
      TabPicture(0)   =   "NivelDoUsuario.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line19"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line21"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line7"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Op7000"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Op4000"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Op3000"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "OpK003"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "OpK002"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "OpK001"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "OpK000"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Op2003"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Op2002"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Op2001"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "OpG003"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "OpG002"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "OpG001"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Op5000"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Op2004"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Op2000"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "OpG000"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Op1003"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Op1002"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Op1001"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Op1000"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "OpP000"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Op8000"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Op8001"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Op8002"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Op8003"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Op8004"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Op9000"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Op9001"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Op9002"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Op9003"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Op9004"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "OpJ000"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Op6000"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "OpH000"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "OpI000"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "OpL000"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Op3001"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Op5001"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Op5002"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Op4001"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Op4003"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Op4002"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Op3002"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "OpL002"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Op7001"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).ControlCount=   67
      TabCaption(1)   =   "Relatórios do Sistema"
      TabPicture(1)   =   "NivelDoUsuario.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Op3003"
      Tab(1).Control(1)=   "OpR007"
      Tab(1).Control(2)=   "OpR006"
      Tab(1).Control(3)=   "OpR005"
      Tab(1).Control(4)=   "OpR004"
      Tab(1).Control(5)=   "OpR003"
      Tab(1).Control(6)=   "OpR002"
      Tab(1).Control(7)=   "OpR001"
      Tab(1).ControlCount=   8
      Begin VB.CheckBox Op3003 
         Caption         =   "Etiquetas"
         Height          =   195
         Left            =   -71550
         TabIndex        =   57
         Top             =   945
         Width           =   960
      End
      Begin VB.CheckBox Op7001 
         Caption         =   "Visitantes"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4365
         TabIndex        =   56
         Top             =   3495
         Width           =   1005
      End
      Begin VB.CheckBox OpL002 
         Caption         =   "Boleto Bancário (Imprimir)"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   55
         Top             =   3780
         Width           =   2205
      End
      Begin VB.CheckBox Op3002 
         Caption         =   "Estorno de Pagamento"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3090
         TabIndex        =   54
         Top             =   3780
         Width           =   1980
      End
      Begin VB.CheckBox Op4002 
         Caption         =   "Modificar Dependentes"
         Height          =   195
         Left            =   105
         TabIndex        =   53
         Top             =   1830
         Width           =   1980
      End
      Begin VB.CheckBox Op4003 
         Caption         =   "Excluir Dependentes"
         Height          =   195
         Left            =   105
         TabIndex        =   52
         Top             =   2085
         Width           =   1875
      End
      Begin VB.CheckBox Op4001 
         Caption         =   "Alterar Dependentes"
         Height          =   195
         Left            =   105
         TabIndex        =   51
         Top             =   1575
         Width           =   1905
      End
      Begin VB.CheckBox Op5002 
         Caption         =   "Trocar Foto"
         Height          =   195
         Left            =   105
         TabIndex        =   50
         Top             =   1305
         Width           =   1245
      End
      Begin VB.CheckBox Op5001 
         Caption         =   "Relacionar Foto"
         Height          =   195
         Left            =   105
         TabIndex        =   49
         Top             =   1065
         Width           =   1575
      End
      Begin VB.CheckBox Op3001 
         Caption         =   "Corrigir Lançam. em Atrazo"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   48
         Top             =   2460
         Width           =   2280
      End
      Begin VB.CheckBox OpR007 
         Caption         =   "Listagem Bancária"
         Height          =   225
         Left            =   -74670
         TabIndex        =   44
         Top             =   3105
         Width           =   1890
      End
      Begin VB.CheckBox OpR006 
         Caption         =   "Relação de Miores de 21 Anos"
         Height          =   255
         Left            =   -74670
         TabIndex        =   43
         Top             =   2730
         Width           =   2715
      End
      Begin VB.CheckBox OpR005 
         Caption         =   "Caixa"
         Height          =   225
         Left            =   -74670
         TabIndex        =   42
         Top             =   2385
         Width           =   810
      End
      Begin VB.CheckBox OpR004 
         Caption         =   "Extrato do Sócio"
         Height          =   195
         Left            =   -74670
         TabIndex        =   41
         Top             =   2025
         Width           =   1575
      End
      Begin VB.CheckBox OpR003 
         Caption         =   "Contas a Receber"
         Height          =   225
         Left            =   -74670
         TabIndex        =   40
         Top             =   1650
         Width           =   1755
      End
      Begin VB.CheckBox OpR002 
         Caption         =   "Contas a Pagar"
         Height          =   225
         Left            =   -74655
         TabIndex        =   39
         Top             =   1305
         Width           =   1485
      End
      Begin VB.CheckBox OpR001 
         Caption         =   "Sócios"
         Height          =   210
         Left            =   -74655
         TabIndex        =   38
         Top             =   975
         Width           =   990
      End
      Begin VB.CheckBox OpL000 
         Caption         =   "Reimprime Recibo"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   37
         Top             =   2100
         Width           =   1740
      End
      Begin VB.CheckBox OpI000 
         Caption         =   "Recebimento Antecipado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6180
         TabIndex        =   36
         Top             =   1710
         Width           =   2220
      End
      Begin VB.CheckBox OpH000 
         Caption         =   "Recibo de Mensalidade"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   35
         Top             =   1335
         Width           =   2325
      End
      Begin VB.CheckBox Op6000 
         Caption         =   "Mensalidades Extras"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   34
         Top             =   975
         Width           =   2265
      End
      Begin VB.CheckBox OpJ000 
         Caption         =   "Cancelamento de Recibo"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   33
         Top             =   615
         Width           =   2265
      End
      Begin VB.CheckBox Op9004 
         Caption         =   "Quitar"
         Height          =   195
         Left            =   4365
         TabIndex        =   32
         Top             =   3105
         Width           =   765
      End
      Begin VB.CheckBox Op9003 
         Caption         =   "Excluir"
         Height          =   195
         Left            =   3090
         TabIndex        =   31
         Top             =   3105
         Width           =   765
      End
      Begin VB.CheckBox Op9002 
         Caption         =   "Modificar"
         Height          =   195
         Left            =   4365
         TabIndex        =   30
         Top             =   2850
         Width           =   990
      End
      Begin VB.CheckBox Op9001 
         Caption         =   "Adicionar"
         Height          =   195
         Left            =   3090
         TabIndex        =   29
         Top             =   2850
         Width           =   960
      End
      Begin VB.CheckBox Op9000 
         Caption         =   "Contas Pagar"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3090
         TabIndex        =   28
         Top             =   2565
         Width           =   1440
      End
      Begin VB.CheckBox Op8004 
         Caption         =   "Quitar"
         Height          =   195
         Left            =   4350
         TabIndex        =   27
         Top             =   2190
         Width           =   870
      End
      Begin VB.CheckBox Op8003 
         Caption         =   "Excluir"
         Height          =   195
         Left            =   3090
         TabIndex        =   26
         Top             =   2190
         Width           =   810
      End
      Begin VB.CheckBox Op8002 
         Caption         =   "Modificar"
         Height          =   195
         Left            =   4350
         TabIndex        =   25
         Top             =   1950
         Width           =   945
      End
      Begin VB.CheckBox Op8001 
         Caption         =   "Adicionar"
         Height          =   195
         Left            =   3090
         TabIndex        =   24
         Top             =   1950
         Width           =   960
      End
      Begin VB.CheckBox Op8000 
         Caption         =   "Contas a Receber"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3090
         TabIndex        =   23
         Top             =   1710
         Width           =   1740
      End
      Begin VB.CheckBox OpP000 
         Caption         =   "Gerar Mensalidades de Sócios"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3090
         TabIndex        =   22
         Top             =   1350
         Width           =   2805
      End
      Begin VB.CheckBox Op1000 
         Caption         =   "Clientes"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   105
         TabIndex        =   21
         Top             =   585
         Width           =   975
      End
      Begin VB.CheckBox Op1001 
         Caption         =   "Adicionar"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   825
         Width           =   960
      End
      Begin VB.CheckBox Op1002 
         Caption         =   "Modificar"
         Height          =   195
         Left            =   1095
         TabIndex        =   19
         Top             =   825
         Width           =   945
      End
      Begin VB.CheckBox Op1003 
         Caption         =   "Excluir"
         Height          =   195
         Left            =   2115
         TabIndex        =   18
         Top             =   825
         Width           =   765
      End
      Begin VB.CheckBox OpG000 
         Caption         =   "Bancos"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6180
         TabIndex        =   17
         Top             =   3180
         Width           =   1035
      End
      Begin VB.CheckBox Op2000 
         Caption         =   "Usuários"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   2475
         Width           =   1125
      End
      Begin VB.CheckBox Op2004 
         Caption         =   "Alterar Sua Senha"
         Height          =   195
         Left            =   1245
         TabIndex        =   15
         Top             =   2970
         Width           =   1590
      End
      Begin VB.CheckBox Op5000 
         Caption         =   "Parâmetros"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3090
         TabIndex        =   14
         Top             =   3495
         Width           =   1095
      End
      Begin VB.CheckBox OpG001 
         Caption         =   "Adicionar"
         Height          =   195
         Left            =   6180
         TabIndex        =   13
         Top             =   3435
         Width           =   960
      End
      Begin VB.CheckBox OpG002 
         Caption         =   "Modificar"
         Height          =   195
         Left            =   7185
         TabIndex        =   12
         Top             =   3435
         Width           =   945
      End
      Begin VB.CheckBox OpG003 
         Caption         =   "Excluir"
         Height          =   195
         Left            =   8190
         TabIndex        =   11
         Top             =   3435
         Width           =   765
      End
      Begin VB.CheckBox Op2001 
         Caption         =   "Adicionar"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   2715
         Width           =   960
      End
      Begin VB.CheckBox Op2002 
         Caption         =   "Modificar"
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   2955
         Width           =   945
      End
      Begin VB.CheckBox Op2003 
         Caption         =   "Excluir"
         Height          =   195
         Left            =   1245
         TabIndex        =   8
         Top             =   2730
         Width           =   1080
      End
      Begin VB.CheckBox OpK000 
         Caption         =   "Tipo De Cliente"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   3315
         Width           =   1515
      End
      Begin VB.CheckBox OpK001 
         Caption         =   "Adicionar"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   3600
         Width           =   960
      End
      Begin VB.CheckBox OpK002 
         Caption         =   "Modificar"
         Height          =   195
         Left            =   1110
         TabIndex        =   5
         Top             =   3600
         Width           =   960
      End
      Begin VB.CheckBox OpK003 
         Caption         =   "Excluir"
         Height          =   195
         Left            =   2115
         TabIndex        =   4
         Top             =   3600
         Width           =   765
      End
      Begin VB.CheckBox Op3000 
         Caption         =   "Autorização de Isenção"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6195
         TabIndex        =   3
         Top             =   2820
         Width           =   2085
      End
      Begin VB.CheckBox Op4000 
         Caption         =   "Autoriz. Débito em Conta"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3090
         TabIndex        =   2
         Top             =   615
         Width           =   2145
      End
      Begin VB.CheckBox Op7000 
         Caption         =   "Atualizar Valor das Mensalidades"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3090
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.Line Line7 
         X1              =   6090
         X2              =   9090
         Y1              =   3705
         Y2              =   3705
      End
      Begin VB.Line Line3 
         X1              =   3000
         X2              =   6090
         Y1              =   3735
         Y2              =   3735
      End
      Begin VB.Line Line21 
         X1              =   6090
         X2              =   9060
         Y1              =   2745
         Y2              =   2745
      End
      Begin VB.Line Line20 
         X1              =   6090
         X2              =   9090
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line19 
         X1              =   6090
         X2              =   9030
         Y1              =   2370
         Y2              =   2370
      End
      Begin VB.Line Line18 
         X1              =   6105
         X2              =   9030
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line17 
         X1              =   6090
         X2              =   9045
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line16 
         X1              =   6090
         X2              =   9060
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line15 
         X1              =   6090
         X2              =   9075
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line14 
         X1              =   6090
         X2              =   9045
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line13 
         X1              =   3000
         X2              =   6090
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Line Line12 
         X1              =   3000
         X2              =   6090
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line11 
         X1              =   3000
         X2              =   6090
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line10 
         X1              =   3000
         X2              =   6090
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line9 
         X1              =   3000
         X2              =   6090
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   3000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line6 
         X1              =   3015
         X2              =   6090
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   3000
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   3015
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         X1              =   6090
         X2              =   6090
         Y1              =   525
         Y2              =   5040
      End
      Begin VB.Line Line1 
         X1              =   3000
         X2              =   3000
         Y1              =   525
         Y2              =   5040
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   45
      Top             =   4230
      Width           =   585
   End
End
Attribute VB_Name = "NivelDoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tbUsuarios As Recordset

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdSalvar_Click()
  With tbUsuarios
      .Edit
      !Op1000 = Op1000.Value
      !Op1001 = Op1001.Value
      !Op1002 = Op1002.Value
      !Op1003 = Op1003.Value
      !Op2000 = Op2000.Value
      !Op2001 = Op2001.Value
      !Op2002 = Op2002.Value
      !Op2003 = Op2003.Value
      !Op2004 = Op2004.Value
      !Op3000 = Op3000.Value
      !Op3001 = Op3001.Value
      !Op3002 = Op3002.Value
      !Op3003 = Op3003.Value
      !Op4000 = Op4000.Value
      !Op4001 = Op4001.Value
      !Op4002 = Op4002.Value
      !Op4003 = Op4003.Value
      !Op5000 = Op5000.Value
      !Op5001 = Op5001.Value
      !Op5002 = Op5002.Value
      '!Op5003 = Op5003.Value
      !Op6000 = Op6000.Value
      !Op7000 = Op7000.Value
      !Op7001 = Op7001.Value
      '!Op7002 = Op7002.Value
      '!Op7003 = Op7003.Value
      !Op8000 = Op8000.Value
      !Op8001 = Op8001.Value
      !Op8002 = Op8002.Value
      !Op8003 = Op8003.Value
      !Op8004 = Op8004.Value
      !Op9000 = Op9000.Value
      !Op9001 = Op9001.Value
      !Op9002 = Op9002.Value
      !Op9003 = Op9003.Value
      !Op9004 = Op9004.Value
      !OpG000 = OpG000.Value
      !OpG001 = OpG001.Value
      !OpG002 = OpG002.Value
      !OpG003 = OpG003.Value
      !OpJ000 = OpJ000.Value
      '!OpJ001 = OpJ001.Value
      '!OpJ002 = OpJ002.Value
      '!OpJ003 = OpJ003.Value
      '!OpJ004 = OpJ004.Value
      '!OpJ005 = OpJ005.Value
      !OpK000 = OpK000.Value
      !OpK001 = OpK001.Value
      !OpK002 = OpK002.Value
      !OpK003 = OpK003.Value
      !OpL000 = OpL000.Value
      !OpL002 = OpL002.Value
      !OpP000 = OpP000.Value
      '!OpP001 = OpP001.Value
      '!OpP002 = OpP002.Value
      '!OpP003 = OpP003.Value
      '!OpP004 = OpP004.Value
      '!OpP005 = OpP005.Value
      !OpH000 = OpH000.Value
      !OpI000 = OpI000.Value
      !OpR001 = OpR001.Value
      !OpR002 = OpR002.Value
      !OpR003 = OpR003.Value
      !OpR004 = OpR004.Value
      !OpR005 = OpR005.Value
      !OpR006 = OpR006.Value
      !OpR007 = OpR007.Value
      '!OpR008 = OpR008.Value
      '!OpR009 = OpR009.Value
      '!OpR010 = OpR010.Value
      '!OpR011 = OpR011.Value
      '!OpR012 = OpR012.Value
      '!OpR013 = OpR013.Value
      '!OpR014 = OpR014.Value
      '!OpR015 = OpR015.Value
      '!OpR016 = OpR016.Value
      '!OpR017 = OpR017.Value
      '!OpR018 = OpR018.Value
      '!OpR019 = OpR019.Value
      '!OpR020 = OpR020.Value
      '!OpR021 = OpR021.Value
      .Update
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Call Centraliza(Me)
  Set tbUsuarios = db.OpenRecordset("Usuarios", dbOpenTable)
  Me.Refresh
  DoEvents
  With tbUsuarios
    If Not .NoMatch Then
      Op1000.Value = !Op1000
      Op1001.Value = !Op1001
      Op1002.Value = !Op1002
      Op1003.Value = !Op1003
      Op2000.Value = !Op2000
      Op2001.Value = !Op2001
      Op2002.Value = !Op2002
      Op2003.Value = !Op2003
      Op2004.Value = !Op2004
      Op3000.Value = !Op3000
      Op3001.Value = !Op3001
      Op3002.Value = !Op3002
      Op3003.Value = !Op3003
      Op4000.Value = !Op4000
      Op4001.Value = !Op4001
      Op4002.Value = !Op4002
      Op4003.Value = !Op4003
      Op5000.Value = !Op5000
      Op5001.Value = !Op5001
      Op5002.Value = !Op5002
      'Op5003.Value = !Op5003
      Op6000.Value = !Op6000
      Op7000.Value = !Op7000
      Op7001.Value = !Op7001
      'Op7002.Value = !Op7002
      'Op7003.Value = !Op7003
      Op8000.Value = !Op8000
      Op8001.Value = !Op8001
      Op8002.Value = !Op8002
      Op8003.Value = !Op8003
      Op8004.Value = !Op8004
      Op9000.Value = !Op9000
      Op9001.Value = !Op9001
      Op9002.Value = !Op9002
      Op9003.Value = !Op9003
      Op9004.Value = !Op9004
      OpG000.Value = !OpG000
      OpG001.Value = !OpG001
      OpG002.Value = !OpG002
      OpG003.Value = !OpG003
      OpJ000.Value = !OpJ000
      'OpJ001.Value = !OpJ001
      'OpJ002.Value = !OpJ002
      'OpJ003.Value = !OpJ003
      'OpJ004.Value = !OpJ004
      'OpJ005.Value = !OpJ005
      OpK000.Value = !OpK000
      OpK001.Value = !OpK001
      OpK002.Value = !OpK002
      OpK003.Value = !OpK003
      OpL000.Value = !OpL000
      OpL002.Value = !OpL002
      OpP000.Value = !OpP000
      'OpP001.Value = !OpP001
      'OpP002.Value = !OpP002
      'OpP003.Value = !OpP003
      'OpP004.Value = !OpP004
      'OpP005.Value = !OpP005
      OpH000.Value = !OpH000
      OpI000.Value = !OpI000
      OpR001.Value = !OpR001
      OpR002.Value = !OpR002
      OpR003.Value = !OpR003
      OpR004.Value = !OpR004
      OpR005.Value = !OpR005
      OpR006.Value = !OpR006
      OpR007.Value = !OpR007
      'OpR008.Value = !OpR008
      'OpR009.Value = !OpR009
      'OpR010.Value = !OpR010
      'OpR011.Value = !OpR011
      'OpR012.Value = !OpR012
      'OpR013.Value = !OpR013
      'OpR014.Value = !OpR014
      'OpR015.Value = !OpR015
      'OpR016.Value = !OpR016
      'OpR017.Value = !OpR017
      'OpR018.Value = !OpR018
      'OpR019.Value = !OpR019
      'OpR020.Value = !OpR020
      'OpR021.Value = !OpR021
    End If
  End With
  
  If Op1000.Value = 0 Then
    Op1001.Enabled = False
    Op1002.Enabled = False
    Op1003.Enabled = False
    Op5001.Enabled = False
    Op5002.Enabled = False
    Op4001.Enabled = False
    Op4002.Enabled = False
    Op4003.Enabled = False
  End If
  If Op2000.Value = 0 Then
    Op2001.Enabled = False
    Op2002.Enabled = False
    Op2003.Enabled = False
    Op2004.Enabled = False
  End If
  If Op8000.Value = 0 Then
    Op8001.Enabled = False
    Op8002.Enabled = False
    Op8003.Enabled = False
    Op8004.Enabled = False
  End If
  If Op9000.Value = 0 Then
    Op9001.Enabled = False
    Op9002.Enabled = False
    Op9003.Enabled = False
    Op9004.Enabled = False
  End If
  If OpG000.Value = 0 Then
    OpG001.Enabled = False
    OpG002.Enabled = False
    OpG003.Enabled = False
  End If
  If OpK000.Value = 0 Then
    OpK001.Enabled = False
    OpK002.Enabled = False
    OpK003.Enabled = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set tbUsuarios = Nothing
End Sub

Private Sub Op1000_Click()
  If Op1000.Value = 1 Then
    Op1001.Enabled = True
    Op1002.Enabled = True
    Op1003.Enabled = True
    Op5002.Enabled = True
    Op5001.Enabled = True
    Op4001.Enabled = True
    Op4002.Enabled = True
    Op4003.Enabled = True
  Else
    Op1001.Enabled = False
    Op1002.Enabled = False
    Op1003.Enabled = False
    Op5002.Enabled = False
    Op5001.Enabled = False
    Op4001.Enabled = False
    Op4002.Enabled = False
    Op4003.Enabled = False
    Op1001.Value = 0
    Op1002.Value = 0
    Op1003.Value = 0
    Op5002.Value = 0
    Op5001.Value = 0
    Op4001.Value = 0
    Op4002.Value = 0
    Op4003.Value = 0
  End If
End Sub

Private Sub Op2000_Click()
  If Op2000.Value = 1 Then
    Op2001.Enabled = True
    Op2002.Enabled = True
    Op2003.Enabled = True
    Op2004.Enabled = True
  Else
    Op2001.Enabled = False
    Op2002.Enabled = False
    Op2003.Enabled = False
    Op2004.Enabled = False
    Op2001.Value = 0
    Op2002.Value = 0
    Op2003.Value = 0
    Op2004.Value = 0
  End If
End Sub

Private Sub Op8000_Click()
  If Op8000.Value = 1 Then
    Op8001.Enabled = True
    Op8002.Enabled = True
    Op8003.Enabled = True
    Op8004.Enabled = True
  Else
    Op8001.Enabled = False
    Op8002.Enabled = False
    Op8003.Enabled = False
    Op8004.Enabled = False
    Op8001.Value = 0
    Op8002.Value = 0
    Op8003.Value = 0
    Op8004.Value = 0
  End If
End Sub

Private Sub Op9000_Click()
  If Op9000.Value = 1 Then
    Op9001.Enabled = True
    Op9002.Enabled = True
    Op9003.Enabled = True
    Op9004.Enabled = True
  Else
    Op9001.Enabled = False
    Op9002.Enabled = False
    Op9003.Enabled = False
    Op9004.Enabled = False
    Op9001.Value = 0
    Op9002.Value = 0
    Op9003.Value = 0
    Op9004.Value = 0
  End If
End Sub

Private Sub OpG000_Click()
  If OpG000.Value = 1 Then
    OpG001.Enabled = True
    OpG002.Enabled = True
    OpG003.Enabled = True
  Else
    OpG001.Enabled = False
    OpG002.Enabled = False
    OpG003.Enabled = False
    OpG001.Value = 0
    OpG002.Value = 0
    OpG003.Value = 0
  End If
End Sub

Private Sub OpK000_Click()
  If OpK000.Value = 1 Then
    OpK001.Enabled = True
    OpK002.Enabled = True
    OpK003.Enabled = True
  Else
    OpK001.Enabled = False
    OpK002.Enabled = False
    OpK003.Enabled = False
    OpK001.Value = 0
    OpK002.Value = 0
    OpK003.Value = 0
  End If
End Sub


