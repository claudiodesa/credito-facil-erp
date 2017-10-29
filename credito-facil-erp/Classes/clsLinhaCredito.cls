VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinhaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_idLinhaCredito As Long
Public m_02_idEmpresaCliente As Long
Public m_03_limite As Double
Public m_04_aprovado As String
Public m_05_usuarioInclusao As String
Public m_06_dataInclusao As String
Public m_07_usuarioAlteracao As String
Public m_08_dataAlteracao As String
Public m_09_ctLock As Long

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
    cnConexao.BeginTrans

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO linhaCredito                                 " & vbCrLf
    strSql = strSql & " ( ID_EMPRESACLIENTE                                     " & vbCrLf
    strSql = strSql & "  ,LIMITE                                                " & vbCrLf
    strSql = strSql & "  ,APROVADO                                              " & vbCrLf
    strSql = strSql & "  ,USUARIO_INCLUSAO                                      " & vbCrLf
    strSql = strSql & "  ,DATA_INCLUSAO                                         " & vbCrLf
    strSql = strSql & "  ,USUARIO_ALTERACAO                                     " & vbCrLf
    strSql = strSql & "  ,DATA_ALTERACAO                                        " & vbCrLf
    strSql = strSql & "  ,CT_LOCK)                                              " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & " ( " & m_02_idEmpresaCliente & "                         " & vbCrLf
    strSql = strSql & " , " & Replace(m_03_limite, ",", ".") & "                " & vbCrLf
    strSql = strSql & " ,'" & m_04_aprovado & "'                                " & vbCrLf
    strSql = strSql & " ,'" & m_05_usuarioInclusao & "'                         " & vbCrLf
    strSql = strSql & " , GetDate()                                             " & vbCrLf
    strSql = strSql & " , '" & m_07_usuarioAlteracao & "'                       " & vbCrLf
    strSql = strSql & " , GetDate()                                             " & vbCrLf
    strSql = strSql & " , " & m_09_ctLock & ")                                  " & vbCrLf
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
    crudInsert = m_01_idLinhaCredito

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIDLinhaCredito() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('linhaCredito') + IDENT_INCR('linhaCredito')" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDLinhaCredito = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDLinhaCredito = 0
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
    strSql = strSql & "UPDATE linhaCredito                                    " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   LIMITE = '" & Replace(m_03_limite, ",", ".") & "'   " & vbCrLf
    strSql = strSql & "  ,APROVADO = '" & m_04_aprovado & "'                  " & vbCrLf
    strSql = strSql & "  ,USUARIO_ALTERACAO = '" & m_05_usuarioInclusao & "'  " & vbCrLf
    strSql = strSql & "  ,DATA_ALTERACAO = Getdate()                          " & vbCrLf
    strSql = strSql & "  ,CT_LOCK = " & m_09_ctLock & "                       " & vbCrLf
    strSql = strSql & " WHERE ID_LINHACREDITO = " & m_01_idLinhaCredito
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudUpdate = m_01_idLinhaCredito
    
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
    cnConexao.BeginTrans
    
    strSql = ""
    strSql = strSql & "DELETE                                         " & vbCrLf
    strSql = strSql & "  FROM linhaCredito                            " & vbCrLf
    strSql = strSql & " WHERE ID_LINHACREDITO = " & m_01_idLinhaCredito & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
    crudDelete = m_01_idLinhaCredito

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consulta(ByVal idEmpresaCliente As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM linhaCredito WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_EMPRESACLIENTE = " & idEmpresaCliente

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function recuperarEmpresasComLinhadeCreditoAprovada() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " Select" & vbCrLf
    ssql = ssql & " Case empresaCliente.TIPO" & vbCrLf
    ssql = ssql & "     when 'F' then empresaCliente.NOME_PESSOA_FISICA" & vbCrLf
    ssql = ssql & "     when 'J' then empresaCliente.NOME_FANTASIA" & vbCrLf
    ssql = ssql & " end as NOME,* from linhaCredito,empresaCliente with(nolock)" & vbCrLf
    ssql = ssql & " Where empresaCliente.ID_EMPRESACLIENTE = linhaCredito.ID_EMPRESACLIENTE" & vbCrLf
    ssql = ssql & " and APROVADO = 'S'" & vbCrLf
    ssql = ssql & " order by NOME" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperarEmpresasComLinhadeCreditoAprovada = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
