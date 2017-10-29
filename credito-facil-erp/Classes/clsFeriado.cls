VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFERIADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_DATA_FERIADO As Date
Public m_02_FERIADO As String
Public m_03_EXCEPCIONALMENTE As String

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
    strSql = strSql & "INSERT INTO FERIADO                                      " & vbCrLf
    strSql = strSql & "            (DATA_FERIADO,                               " & vbCrLf
    strSql = strSql & "             FERIADO,                                    " & vbCrLf
    strSql = strSql & "             EXCEPCIONALMENTE)                           " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        ('" & m_01_DATA_FERIADO & "'                     " & vbCrLf
    strSql = strSql & "        ,'" & m_02_FERIADO & "'                          " & vbCrLf
    strSql = strSql & "        ,'" & m_03_EXCEPCIONALMENTE & "')                " & vbCrLf
        
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudInsert = m_01_DATA_FERIADO

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
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
    strSql = strSql & "UPDATE FERIADO                                         " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   FERIADO = '" & m_02_FERIADO & "'                    " & vbCrLf
    strSql = strSql & " , EXCEPCIONALMENTE = '" & m_03_EXCEPCIONALMENTE & "'  " & vbCrLf
    strSql = strSql & " WHERE datepart(dd,DATA_FERIADO) = " & Mid(m_01_DATA_FERIADO, 1, 2) & vbCrLf
    strSql = strSql & "   AND datepart(mm,DATA_FERIADO) = " & Mid(m_01_DATA_FERIADO, 4, 2) & vbCrLf
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudUpdate = m_01_DATA_FERIADO
    
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
    strSql = strSql & "SET DATEFORMAT DMY                               " & vbCrLf
    strSql = strSql & "DELETE                                           " & vbCrLf
    strSql = strSql & "  FROM FERIADO                                   " & vbCrLf
    strSql = strSql & "  WHERE DATEPART(DD,DATA_FERIADO) = DATEPART(DD,'" & m_01_DATA_FERIADO & "') AND" & vbCrLf
    strSql = strSql & "        DATEPART(MM,DATA_FERIADO) = DATEPART(MM,'" & m_01_DATA_FERIADO & "')"
    'strSql = strSql & " WHERE DATA_FERIADO = '" & m_01_DATA_FERIADO & "'" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudDelete = m_01_DATA_FERIADO

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function Consulta_By_Data(ByVal dtData As Date, _
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
    ssql = ssql & " SET DATEFORMAT DMY" & vbCrLf
    ssql = ssql & " SELECT * FROM FERIADO WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE datepart(dd,DATA_FERIADO) = " & Mid(CStr(dtData), 1, 2) & vbCrLf
    ssql = ssql & "   AND datepart(mm,DATA_FERIADO) = " & Mid(CStr(dtData), 4, 2) & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set Consulta_By_Data = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarFeriados(ByVal strConexao As String, ByVal lngTimeOut As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = lngTimeOut
    mCn.ConnectionString = strConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SET DATEFORMAT DMY" & vbCrLf
    ssql = ssql & " SELECT" & vbCrLf
    ssql = ssql & " '(' + RIGHT ('0' + CONVERT(VARCHAR(10),DATEPART (DD ,DATA_FERIADO)),2) + '/'" & vbCrLf
    ssql = ssql & "     + RIGHT ('0' + CONVERT(VARCHAR(10),DATEPART (MM ,DATA_FERIADO)),2) + ')'" & vbCrLf
    ssql = ssql & " + ' - ' + FERIADO AS DESCRICAO,* FROM feriado WITH(NOLOCK)" & vbCrLf
    ssql = ssql & " ORDER BY datepart(mm,DATA_FERIADO),datepart(dd,DATA_FERIADO)" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarFeriados = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function DataFeriado(ByVal dtData As Date) As Boolean

    Dim mCn       As ADODB.Connection
    Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SET DATEFORMAT DMY" & vbCrLf
    ssql = ssql & " SELECT * FROM FERIADO WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE datepart(dd,DATA_FERIADO) = substring('" & dtData & "',1,2)" & vbCrLf
    ssql = ssql & "  AND datepart(mm,DATA_FERIADO) = substring('" & dtData & "',4,2)" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
       
    If Not mRs.EOF Then
        DataFeriado = True
    End If
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
