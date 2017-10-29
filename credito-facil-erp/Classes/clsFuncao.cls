VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFUNCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_FUNCAO As Long
Public m_02_DESCRICAO_FUNCAO As String
Public m_03_STATUS_FUNCAO As String
Public m_04_SIGLA_FUNCAO As String
Public m_05_USUARIO_INCLUSAO As String
Public m_06_DATA_INCLUSAO As String
Public m_07_USUARIO_ALTERACAO As String
Public m_08_DATA_ALTERACAO As String
Public m_09_CT_LOCK As Long
'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public mTIMEOUT As String
Public mSTRING_CONEXAO As String

Public Function crudInsert() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo TrataErro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO FUNCAO                                       " & vbCrLf
   'strSql = strSql & "             (ID_FUNCAO,                                 " & vbCrLf
    strSql = strSql & "            (DESCRICAO_FUNCAO,                           " & vbCrLf
    strSql = strSql & "             STATUS_FUNCAO,                              " & vbCrLf
    strSql = strSql & "             SIGLA_FUNCAO,                               " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                           " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                              " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                          " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                             " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                    " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
   'strSql = strSql & "        (" & m_01_ID_FUNCAO & vbCrLf
    strSql = strSql & "        ('" & m_02_DESCRICAO_FUNCAO & "'                 " & vbCrLf
    strSql = strSql & "        ,'" & m_03_STATUS_FUNCAO & "'                    " & vbCrLf
    strSql = strSql & "        ,'" & m_04_SIGLA_FUNCAO & "'                     " & vbCrLf
    strSql = strSql & "        ,'" & m_05_USUARIO_INCLUSAO & "'                 " & vbCrLf
    strSql = strSql & "        ,'" & m_06_DATA_INCLUSAO & "'                    " & vbCrLf
    strSql = strSql & "        ,'" & m_07_USUARIO_ALTERACAO & "'                " & vbCrLf
    strSql = strSql & "        ,'" & m_08_DATA_ALTERACAO & "'                   " & vbCrLf
    strSql = strSql & "        ," & m_09_CT_LOCK & ")                           " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudInsert = m_01_ID_FUNCAO

Exit Function
TrataErro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIDFuncao() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo TrataErro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('FUNCAO') + IDENT_INCR('FUNCAO')" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDFuncao = mRs(0)

Exit Function
TrataErro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDFuncao = 0
End Function
Public Function crudUpdate() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo TrataErro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                     " & vbCrLf
    strSql = strSql & "UPDATE FUNCAO                                          " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   DESCRICAO_FUNCAO = '" & m_02_DESCRICAO_FUNCAO & "'  " & vbCrLf
    strSql = strSql & " , STATUS_FUNCAO = '" & m_03_STATUS_FUNCAO & "'        " & vbCrLf
    strSql = strSql & " , SIGLA_FUNCAO = '" & m_04_SIGLA_FUNCAO & "'          " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_07_USUARIO_ALTERACAO & "'" & vbCrLf
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                          " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_09_CT_LOCK & "                      " & vbCrLf
    strSql = strSql & " WHERE ID_FUNCAO = " & m_01_ID_FUNCAO
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudUpdate = m_01_ID_FUNCAO
    
Exit Function
TrataErro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
Public Function crudDelete() As Long
Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo TrataErro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open
    
    strSql = ""
    strSql = strSql & "DELETE                              " & vbCrLf
    strSql = strSql & "  FROM FUNCAO                       " & vbCrLf
    strSql = strSql & " WHERE ID_FUNCAO = " & m_01_ID_FUNCAO & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudDelete = m_01_ID_FUNCAO

Exit Function
TrataErro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function Consulta_By_Codigo(ByVal lngCodFuncao As Long, _
                                   ByVal strConexao As String, _
                                   ByVal lngTimeOut As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo TrataErro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = lngTimeOut
    mCn.ConnectionString = strConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM FUNCAO WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_FUNCAO = '" & lngCodFuncao & "' " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set Consulta_By_Codigo = mRs
  
    Set mCn = Nothing

Exit Function
TrataErro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarFuncoes(ByVal strConexao As String, ByVal lngTimeOut As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo TrataErro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = lngTimeOut
    mCn.ConnectionString = strConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT * FROM FUNCAO WITH (NOLOCK) "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarFuncoes = mRs
  
    Set mCn = Nothing

Exit Function
TrataErro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

