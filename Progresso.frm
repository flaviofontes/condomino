VERSION 5.00
Object = "{CCF682E3-14F1-11D7-80E2-10E08C102140}#1.0#0"; "ZProgBar.ocx"
Begin VB.Form Progresso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   HelpContextID   =   4240
   Icon            =   "Progresso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ZealProgressBar.ProgressBar Barra 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5295
      _ExtentX        =   9340
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
   Begin VB.Label Label1 
      Height          =   615
      Left            =   75
      TabIndex        =   1
      Top             =   405
      Width           =   5280
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Progresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

