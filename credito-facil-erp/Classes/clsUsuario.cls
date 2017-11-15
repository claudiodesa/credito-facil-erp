VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_idUsuario As Long
Public m_02_idFuncionario As Long
Public m_03_status As String
Public m_04_login As String
Public m_05_senha As String
Public m_06_usuarioInclusao As String
Public m_07_dataInclusao As String
Public m_08_usuarioAlteracao As String
Public m_09_dataAlteracao As String
Public m_10_ctLock As Long
'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String


Public Function crudInsert() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO usuario                                      " & vbCrLf
    strSql = strSql & "            (ID_FUNCIONARIO,                             " & vbCrLf
    strSql = strSql & "             STATUS,                                     " & vbCrLf
    strSql = strSql & "             LOGIN,                                      " & vbCrLf
    strSql = strSql & "             SENHA,                                      " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                           " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                              " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                          " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                             " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                    " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        ( " & m_02_idFuncionario & "                     " & vbCrLf
    strSql = strSql & "        ,'" & m_03_status & "'                           " & vbCrLf
    strSql = strSql & "        ,'" & m_04_login & "'                            " & vbCrLf
    strSql = strSql & "        ,'" & m_05_senha & "'                            " & vbCrLf
    strSql = strSql & "        ,'" & m_06_usuarioInclusao & "'                  " & vbCrLf
    strSql = strSql & "        ,GetDate()                                       " & vbCrLf
    strSql = strSql & "        ,'" & m_08_usuarioAlteracao & "'                 " & vbCrLf
    strSql = strSql & "        ,GetDate()                                       " & vbCrLf
    strSql = strSql & "        ," & m_10_ctLock & ")                            " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
    crudInsert = m_01_idUsuario

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIdUsuario() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('usuario') + IDENT_INCR('usuario')" & vbCrLf
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIdUsuario = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIdUsuario = 0
End Function
Public Function crudUpdate() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                     " & vbCrLf
    strSql = strSql & "UPDATE usuario                                         " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   STATUS = '" & m_03_status & "'                      " & vbCrLf
    strSql = strSql & " , LOGIN = '" & m_04_login & "'                        " & vbCrLf
    strSql = strSql & " , SENHA = '" & m_05_senha & "'                        " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_08_usuarioAlteracao & "' " & vbCrLf
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                          " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_10_ctLock & "                       " & vbCrLf
    strSql = strSql & " WHERE ID_USUARIO = " & m_01_idUsuario
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
    crudUpdate = m_01_idUsuario
    
Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
Public Function crudDelete() As Long
Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    strSql = ""
    strSql = strSql & "DELETE                               " & vbCrLf
    strSql = strSql & "  FROM usuario                       " & vbCrLf
    strSql = strSql & " WHERE ID_USUARIO = " & m_01_idUsuario & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
    crudDelete = m_01_idUsuario

Exit Function
trataerro:
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consulta(ByVal idFuncionario As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM usuario WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_FUNCIONARIO = '" & idFuncionario & "' " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function ConsultaLogin(ByVal strLogin As String) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM usuario WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE LOGIN = '" & strLogin & "'    " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set ConsultaLogin = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function ConsultaIdFuncionario(ByVal idFuncionario As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM usuario WITH (NOLOCK)           " & vbCrLf
    ssql = ssql & " WHERE ID_FUNCIONARIO = '" & idFuncionario & "'" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set ConsultaIdFuncionario = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarUsuarios() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT * FROM usuario WITH (NOLOCK) "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarUsuarios = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function SenhaUnica(ByVal strSenha As String) As Boolean

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM usuario WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE SENHA = '" & strSenha & "'    " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    If mRs.EOF Then
        SenhaUnica = True
    Else
        SenhaUnica = False
    End If
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

