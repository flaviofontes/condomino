VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form BoletosAchados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "BoletosAchados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView ListView1 
      Height          =   1875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Inquilino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Condomínio"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "BoletosAchados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim node As ListItem
Dim hw As Long

Public Sub Carregar(ByRef dados As Recordset)
  With dados
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        Set node = ListView1.ListItems.Add(, "A" + Str(!id))
        node.Text = NomeCompleto(!cdsc)
        node.SubItems(1) = Format(!valr, "#,##0.00")
        node.SubItems(2) = !Condominio
        .MoveNext
      Loop
    End If
  End With
  Call AutoAjusteListView(ListView1, 0)
  Call AutoAjusteListView(ListView1, 2)
  hw = ListView1.ListItems(1).Width
  ListView1.Width = hw + 500
  Me.Width = hw + 600 + (ListView1.Left * 2)
  Me.Show 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 27
      KeyAscii = 0
      RetCodigo = 0
      Unload Me
    Case 13
      KeyAscii = 0
      RetCodigo = Val(Mid(ListView1.SelectedItem.Key, 2))
      Unload Me
  End Select
End Sub

Private Sub ListView1_DblClick()
  RetCodigo = Val(Mid(ListView1.SelectedItem.Key, 2))
  Unload Me
End Sub
