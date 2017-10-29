VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsESTADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public mTIMEOUT As String
Public mSTRING_CONEXAO As String


Public Function RecuperaEstados() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = "SELECT * FROM ESTADO WITH (NOLOCK) "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperaEstados = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperaNomeEstado(ByVal idEstado As Long) As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = "SELECT DESCRICAO FROM estado WITH(NOLOCK) WHERE ID_ESTADO = " & idEstado

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    RecuperaNomeEstado = mRs(0)
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function



