VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRAMOATIVIDADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_RAMO As Long
Public m_02_DESCRICAO As String
'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public mTIMEOUT As String
Public mSTRING_CONEXAO As String

Public Function crudInsert() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO RAMOATIVIDADE                                " & vbCrLf
    strSql = strSql & "            (DESCRICAO)                                  " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        ('" & m_02_DESCRICAO & "')                       " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudInsert = m_01_ID_RAMO

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIDRamo() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('RAMOATIVIDADE') + IDENT_INCR('RAMOATIVIDADE')" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDRamo = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDRamo = 0
End Function
Public Function crudUpdate() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                     " & vbCrLf
    strSql = strSql & "UPDATE RAMOATIVIDADE                                   " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   DESCRICAO = '" & m_02_DESCRICAO & "'                " & vbCrLf
    strSql = strSql & " WHERE ID_RAMO = " & m_01_ID_RAMO
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudUpdate = m_01_ID_RAMO
    
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
Public Function crudDelete() As Long
Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open
    
    strSql = ""
    strSql = strSql & "DELETE                              " & vbCrLf
    strSql = strSql & "  FROM RAMOATIVIDADE                " & vbCrLf
    strSql = strSql & " WHERE ID_RAMO = " & m_01_ID_RAMO & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudDelete = m_01_ID_RAMO

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function Consulta_By_Codigo(ByVal lngCodRamo As Long, _
                                   ByVal strConexao As String, _
                                   ByVal lngTimeOut As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = lngTimeOut
    mCn.ConnectionString = strConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM RAMOATIVIDADE WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_RAMO = '" & lngCodRamo & "' " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set Consulta_By_Codigo = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarRamos() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = "SELECT * FROM RAMOATIVIDADE WITH (NOLOCK) ORDER BY DESCRICAO"

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarRamos = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function



