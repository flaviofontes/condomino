VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form LocCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Cheques"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   ControlBox      =   0   'False
   Icon            =   "LocCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2580
      Top             =   4860
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   315
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4860
      Width           =   975
   End
   Begin VB.TextBox cpNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   4860
      Width           =   1275
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "LocCheques.frx":000C
      Height          =   4635
      Left            =   60
      OleObjectBlob   =   "LocCheques.frx":0020
      TabIndex        =   1
      Top             =   60
      Width           =   7215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\programacao\Porto real\Fontes\porto-real\dados.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3060
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CHEQUES"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   555
   End
End
Attribute VB_Name = "LocCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
  sLocCheque(0) = ""
  sLocCheque(1) = ""
  sLocCheque(2) = ""
  sLocCheque(3) = ""
  Unload Me
End Sub

Private Sub cmdOk_Click()
  If DBGrid1.ApproxCount > 0 Then
    sLocCheque(0) = DBGrid1.Columns(0).Text
    sLocCheque(1) = DBGrid1.Columns(1).Text
    sLocCheque(2) = DBGrid1.Columns(2).Text
    sLocCheque(3) = DBGrid1.Columns(3).Text
  End If
  Unload Me
End Sub

Private Sub cpNumero_Change()
  With Data1.Recordset
    .Seek "=", Format$(cpNumero.Text, "000000")
  End With
End Sub

Private Sub Form_Load()
  Data1.DatabaseName = Parametros.dados
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Data1.Recordset.Index = "numero"
  If Data1.Recordset.RecordCount > 0 Then
    Data1.Recordset.MoveLast
  End If
End Sub
