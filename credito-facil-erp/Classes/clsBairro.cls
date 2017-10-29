VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBAIRRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_BAIRRO As Long
Public m_02_DESCRICAO_BAIRRO As String
Public m_03_ID_MUNICIPIO As String
Public m_04_ID_ESTADO As String
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

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO BAIRRO                                       " & vbCrLf
    strSql = strSql & "            (DESCRICAO_BAIRRO,                           " & vbCrLf
    strSql = strSql & "             ID_MUNICIPIO,                               " & vbCrLf
    strSql = strSql & "             ID_ESTADO,                                  " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                           " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                              " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                          " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                             " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                    " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        ('" & m_02_DESCRICAO_BAIRRO & "'                 " & vbCrLf
    strSql = strSql & "        ,'" & m_03_ID_MUNICIPIO & "'                     " & vbCrLf
    strSql = strSql & "        ,'" & m_04_ID_ESTADO & "'                        " & vbCrLf
    strSql = strSql & "        ,'" & m_05_USUARIO_INCLUSAO & "'                 " & vbCrLf
    strSql = strSql & "        ,'" & m_06_DATA_INCLUSAO & "'                    " & vbCrLf
    strSql = strSql & "        ,'" & m_07_USUARIO_ALTERACAO & "'                " & vbCrLf
    strSql = strSql & "        ,'" & m_08_DATA_ALTERACAO & "'                   " & vbCrLf
    strSql = strSql & "        , " & m_09_CT_LOCK & ")                          " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudInsert = m_01_ID_BAIRRO

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIDBairro() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = "SELECT isnull(IDENT_CURRENT('BAIRRO') + IDENT_INCR('BAIRRO'),1)" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDBairro = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDBairro = 0
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
    strSql = strSql & "UPDATE BAIRRO                                          " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   DESCRICAO_BAIRRO = '" & m_02_DESCRICAO_BAIRRO & "'  " & vbCrLf
    strSql = strSql & " , ID_MUNICIPIO = '" & m_03_ID_MUNICIPIO & "'          " & vbCrLf
    strSql = strSql & " , ID_ESTADO = '" & m_04_ID_ESTADO & "'                " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_07_USUARIO_ALTERACAO & "'" & vbCrLf
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                          " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_09_CT_LOCK & "                      " & vbCrLf
    strSql = strSql & " WHERE ID_BAIRRO = " & m_01_ID_BAIRRO
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudUpdate = m_01_ID_BAIRRO
    
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
    strSql = strSql & "  FROM BAIRRO                       " & vbCrLf
    strSql = strSql & " WHERE ID_BAIRRO = " & m_01_ID_BAIRRO & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudDelete = m_01_ID_BAIRRO

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function Consulta(ByVal idBairro As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM BAIRRO WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_BAIRRO = '" & idBairro & "' " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set Consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperaNomeBairro(ByVal idBairro As Long) As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT DESCRICAO_BAIRRO FROM bairro WITH(NOLOCK) WHERE ID_BAIRRO = " & idBairro

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    RecuperaNomeBairro = mRs(0)
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarBairros(ByVal strConexao As String, ByVal lngTimeOut As Long, ByVal lngIdEstado As Long, lngIdMunicipio As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = lngTimeOut
    mCn.ConnectionString = strConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & "SELECT * FROM BAIRRO WITH (NOLOCK) " & vbCrLf
    ssql = ssql & "WHERE ID_ESTADO = " & lngIdEstado & vbCrLf
    ssql = ssql & "  AND ID_MUNICIPIO = " & lngIdMunicipio & vbCrLf
    ssql = ssql & "order by descricao_bairro"
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarBairros = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function



