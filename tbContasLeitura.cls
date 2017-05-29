VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbContasLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_leitura As Integer
Dim data_leitura As Date
Dim ant_leitura As Long
Dim atual_leitura As Long
Dim total_leitura As Long
Dim valor_leitura As Double
Dim cond_leitura As Long
Dim asso_leitura As Long
Dim codigo_condominio As Long
Dim tipo_leitura As Long

Public Property Let TipoLeitura(ByVal iCod As Long)
  tipo_leitura = iCod
End Property

Public Property Get TipoLeitura() As Long
  TipoLeitura = tipo_leitura
End Property

Public Property Get idLeitura() As Integer
  idLeitura = id_leitura
End Property

Public Property Let dataLeitura(ByVal Valor As Date)
  data_leitura = Valor
End Property

Public Property Get dataLeitura() As Date
  dataLeitura = data_leitura
End Property

Public Property Let leituraAnterior(ByVal Valor As Long)
  ant_leitura = Valor
End Property

Public Property Get leituraAnterior() As Long
  leituraAnterior = ant_leitura
End Property

Public Property Let leituraAtual(ByVal Valor As Long)
  atual_leitura = Valor
End Property

Public Property Get leituraAtual() As Long
  leituraAtual = atual_leitura
End Property

Public Property Get totalGasto() As Long
  totalGasto = atual_leitura - ant_leitura
End Property

Public Property Let valorGasto(ByVal Valor As Double)
  valor_leitura = Valor
End Property

Public Property Get valorGasto() As Double
  valorGasto = valor_leitura
End Property

Public Property Let Associado(ByVal asso As Long)
  asso_leitura = asso
End Property

Public Property Get Associado() As Long
  Associado = asso_leitura
End Property

Public Property Let CodigoCondominio(ByVal Codigo As Long)
  codigo_condominio = Codigo
End Property

Public Property Get CodigoCondominio() As Long
  CodigoCondominio = codigo_condominio
End Property

Public Function Incluir() As Long
  Dim rs As Recordset
  Dim Ret As Long
  
  Ret = 0
  total_leitura = atual_leitura - ant_leitura
  Set rs = db.OpenRecordset("CONTAS_LEITURA", dbOpenTable)
  With rs
    .AddNew
    .Fields("DATA_LEITURA").Value = data_leitura
    .Fields("ANT_LEITURA").Value = ant_leitura
    .Fields("ATUAL_LEITURA").Value = atual_leitura
    .Fields("TOTAL_LEITURA").Value = total_leitura
    .Fields("VALOR_LEITURA").Value = valor_leitura
    .Fields("ASSO_LEITURA").Value = asso_leitura
    .Fields("CODIGO_CONDOMINIO").Value = codigo_condominio
    .Fields("TIPO_LEITURA").Value = tipo_leitura
    .Update
    .Bookmark = .LastModified
    Ret = .Fields("ID_LEITURA").Value
  End With
  Incluir = Ret
End Function

Public Function Alterar(ByVal id As Long) As Boolean
  Dim Ret As Boolean
  Dim rs As Recordset
  
  Ret = False
  Set rs = db.OpenRecordset("select * from CONTAS_LEITURA where ID_LEITURA = " & id & ";", dbOpenDynaset)
  total_leitura = atual_leitura - ant_leitura
  
  With rs
    If .RecordCount = 1 Then
      .MoveFirst
      .Edit
      !data_leitura = data_leitura
      !ant_leitura = ant_leitura
      !atual_leitura = atual_leitura
      !total_leitura = total_leitura
      !valor_leitura = valor_leitura
      !asso_leitura = asso_leitura
      !codigo_condominio = codigo_condominio
      !tipo_leitura = tipo_leitura
      .Update
      Ret = True
    End If
  End With
  Set rs = Nothing
  Alterar = Ret
End Function

Public Function Excluir(ByVal id As Long) As Boolean
  Dim sql As String
  Dim Ret As Boolean
  
  sql = "delete from CONTAS_LEITURA where ID_LEITURA = " & id & ";"
  
  db.Execute sql
  If db.RecordsAffected > 1 Then
    Ret = True
  Else
    Ret = False
  End If
End Function

Public Sub Selecionar(ByVal id As Long)
  Dim rs As Recordset
  
  data_leitura = Null
  ant_leitura = 0
  atual_leitura = 0
  total_leitura = 0
  valor_leitura = 0
  asso_leitura = 0
  codigo_condominio = -1
  
  Set rs = db.OpenRecordset("select * from CONTAS_LEITURA where ID_LEITURA = " & id & ";", dbOpenDynaset)
  
  With rs
    If .RecordCount = 1 Then
      .MoveFirst
      data_leitura = !data_leitura
      ant_leitura = !ant_leitura
      atual_leitura = !atual_leitura
      total_leitura = !total_leitura
      valor_leitura = !valor_leitura
      asso_leitura = !asso_leitura
      codigo_condominio = !codigo_condominio
      tipo_leitura = !tipo_leitura
    End If
  End With
  Set rs = Nothing
End Sub

Public Function Tabela() As Recordset
  Dim rs As Recordset
  Set rs = db.OpenRecordset("CONTAS_LEITURA", dbOpenTable)
  Set Tabela = rs
End Function
