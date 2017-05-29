VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form LocHistoricos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de históricos"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "LocHistoricos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   435
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1455
   End
   Begin ComctlLib.TreeView tvHistoricos 
      Height          =   6495
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   11456
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "LocHistoricos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As Recordset
Dim nNode As node

Private Sub cmdCancelar_Click()
  RetNome = ""
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim i As Integer
  
  For i = 1 To tvHistoricos.Nodes.Count
    If tvHistoricos.Nodes.Item(i).Selected Then
      RetNome = tvHistoricos.Nodes.Item(i).Text
      Exit For
    End If
  Next i
  Unload Me
End Sub

Private Sub Form_Load()
  Set rs = db.OpenRecordset("select * from historicos order by historico;", dbOpenDynaset)
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        Set nNode = tvHistoricos.Nodes.Add()
        nNode.Text = !Historico
        .MoveNext
      Loop
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub
