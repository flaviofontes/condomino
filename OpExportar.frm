VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form OpExportar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportação"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   Icon            =   "OpExportar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3180
      TabIndex        =   9
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1860
      Width           =   1095
   End
   Begin rdActiveText.ActiveText ActiveText3 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   1140
      Width           =   615
      _ExtentX        =   1085
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
      Text            =   "ActiveText3"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText ActiveText2 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1140
      Width           =   675
      _ExtentX        =   1191
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
      Text            =   "ActiveText2"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Páginas"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Todas as paginas"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "..."
      Height          =   315
      Left            =   5220
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin rdActiveText.ActiveText ActiveText1 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
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
      Text            =   "ActiveText1"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3660
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1980
      TabIndex        =   6
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "OpExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

