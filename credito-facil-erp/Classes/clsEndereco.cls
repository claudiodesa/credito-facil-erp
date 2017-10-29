VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsENDERECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_idSysObject As Long
Public m_02_IdEntidade As Long
Public m_03_idEndereco As Long
Public m_04_tipoLogradouro As String
Public m_05_logradouro As String
Public m_06_numero As String
Public m_07_complemento As String
Public m_08_pontoReferencia As String
Public m_09_cep As String
Public m_10_idBairro As Long
Public m_11_idMunicipio As Long
Public m_12_idEstado As Long
Public m_13_usuarioInclusao As String
Public m_14_dataInclusao As String
Public m_15_usuarioAlteracao As String
Public m_16_dataAlteracao As String
Public m_17_ctLock As Long

'Array de endereco
Public rsEndereco As ADODB.Recordset

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String
Public Sub inicializaEndereco()
  Set rsEndereco = New ADODB.Recordset
  rsEndereco.Fields.Append "ID_OBJECT_ENTIDADE", adBigInt
  rsEndereco.Fields.Append "ID_ENTIDADE", adBigInt
  rsEndereco.Fields.Append "ID_ENDERECO", adBigInt
  rsEndereco.Fields.Append "TIPO_LOGRADOURO", adVarChar, 10
  rsEndereco.Fields.Append "LOGRADOURO", adVarChar, 50
  rsEndereco.Fields.Append "NUMERO", adVarChar, 12
  rsEndereco.Fields.Append "COMPLEMENTO", adVarChar, 25
  rsEndereco.Fields.Append "PONTO_REFERENCIA", adVarChar, 25
  rsEndereco.Fields.Append "CEP", adVarChar, 9
  rsEndereco.Fields.Append "ID_BAIRRO", adBigInt
  rsEndereco.Fields.Append "ID_MUNICIPIO", adBigInt
  rsEndereco.Fields.Append "ID_ESTADO", adBigInt
  rsEndereco.Fields.Append "USUARIO_INCLUSAO", adVarChar, 20
  rsEndereco.Fields.Append "DATA_INCLUSAO", adVarChar, 10
  rsEndereco.Fields.Append "USUARIO_ALTERACAO", adVarChar, 20
  rsEndereco.Fields.Append "DATA_ALTERACAO", adVarChar, 20
  rsEndereco.Fields.Append "CT_LOCK", adInteger
End Sub
'Recebe uma conexão como parâmetro (com a transação aberta)
Public Function crudInsert(ByRef cnConexao As ADODB.Connection) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String

    On Error GoTo trataerro
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO ENDERECO                                     " & vbCrLf
    strSql = strSql & "            (ID_OBJECT_ENTIDADE,                         " & vbCrLf
    strSql = strSql & "             ID_ENTIDADE,                                " & vbCrLf
    strSql = strSql & "             TIPO_LOGRADOURO,                            " & vbCrLf
    strSql = strSql & "             LOGRADOURO,                                 " & vbCrLf
    strSql = strSql & "             NUMERO,                                     " & vbCrLf
    strSql = strSql & "             COMPLEMENTO,                                " & vbCrLf
    strSql = strSql & "             PONTO_REFERENCIA,                           " & vbCrLf
    strSql = strSql & "             CEP,                                        " & vbCrLf
    strSql = strSql & "             ID_BAIRRO,                                  " & vbCrLf
    strSql = strSql & "             ID_MUNICIPIO,                               " & vbCrLf
    strSql = strSql & "             ID_ESTADO,                                  " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                           " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                              " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                          " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                             " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                    " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        ( " & m_01_idSysObject & "                       " & vbCrLf
    strSql = strSql & "        , " & m_02_IdEntidade & "                        " & vbCrLf
    strSql = strSql & "        ,'" & m_04_tipoLogradouro & "'                   " & vbCrLf
    strSql = strSql & "        ,'" & m_05_logradouro & "'                       " & vbCrLf
    strSql = strSql & "        ,'" & m_06_numero & "'                           " & vbCrLf
    'Este campo pode ser null
    If Len(m_07_complemento) > 0 Then
    strSql = strSql & "        ,'" & m_07_complemento & "'                      " & vbCrLf
    Else
    strSql = strSql & "        ,null                                            " & vbCrLf
    End If
    'Este campo pode ser null
    If Len(m_08_pontoReferencia) > 0 Then
    strSql = strSql & "        ,'" & m_08_pontoReferencia & "'                  " & vbCrLf
    Else
    strSql = strSql & "        ,null                                            " & vbCrLf
    End If
    'Este campo pode ser null
    If Len(m_09_cep) > 0 Then
    strSql = strSql & "        ,'" & m_09_cep & "'                              " & vbCrLf
    Else
    strSql = strSql & "        ,null                                            " & vbCrLf
    End If
    strSql = strSql & "        , " & m_10_idBairro & "                          " & vbCrLf
    strSql = strSql & "        , " & m_11_idMunicipio & "                       " & vbCrLf
    strSql = strSql & "        , " & m_12_idEstado & "                          " & vbCrLf
    strSql = strSql & "        ,'" & m_13_usuarioInclusao & "'                  " & vbCrLf
    strSql = strSql & "        ,getDate()                                       " & vbCrLf
    strSql = strSql & "        ,'" & m_15_usuarioAlteracao & "'                 " & vbCrLf
    strSql = strSql & "        ,getDate()                                       " & vbCrLf
    strSql = strSql & "        , " & m_17_ctLock & ")                           " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        
    strSql = ""
    strSql = strSql & " SELECT ID_ENDERECO FROM ENDERECO WITH(NOLOCK) " & vbCrLf
    strSql = strSql & " Where ID_OBJECT_ENTIDADE = " & m_01_idSysObject & " And ID_ENTIDADE = " & m_02_IdEntidade
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        
    crudInsert = mRs("ID_ENDERECO")
    
    Set mRs.ActiveConnection = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
'Recebe uma conexão como parâmetro (com a transação aberta)
Public Function crudUpdate(ByRef cnConexao As ADODB.Connection) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String

    On Error GoTo trataerro
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                     " & vbCrLf
    strSql = strSql & "UPDATE ENDERECO                                        " & vbCrLf
    strSql = strSql & " SET                                                   " & vbCrLf
    strSql = strSql & "   TIPO_LOGRADOURO = '" & m_04_tipoLogradouro & "'     " & vbCrLf
    strSql = strSql & " , LOGRADOURO = '" & m_05_logradouro & "'              " & vbCrLf
    strSql = strSql & " , NUMERO = '" & m_06_numero & "'                      " & vbCrLf
    'Este campo pode ser null
    If Len(m_07_complemento) > 0 Then
    strSql = strSql & " , COMPLEMENTO = '" & m_07_complemento & "'            " & vbCrLf
    Else
    strSql = strSql & " , COMPLEMENTO = null                                  " & vbCrLf
    End If
    'Este campo pode ser null
    If Len(m_08_pontoReferencia) > 0 Then
    strSql = strSql & " , PONTO_REFERENCIA = '" & m_08_pontoReferencia & "'   " & vbCrLf
    Else
    strSql = strSql & " , PONTO_REFERENCIA = null                             " & vbCrLf
    End If
    'Este campo pode ser null
    If Len(m_09_cep) > 0 Then
    strSql = strSql & " , CEP = '" & m_09_cep & "'                            " & vbCrLf
    Else
    strSql = strSql & " , CEP = null                                          " & vbCrLf
    End If
    strSql = strSql & " , ID_BAIRRO = '" & m_10_idBairro & "'                 " & vbCrLf
    strSql = strSql & " , ID_MUNICIPIO = '" & m_11_idMunicipio & "'           " & vbCrLf
    strSql = strSql & " , ID_ESTADO = '" & m_12_idEstado & "'                 " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_15_usuarioAlteracao & "' " & vbCrLf
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                          " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_17_ctLock & "                       " & vbCrLf
    strSql = strSql & " WHERE ID_OBJECT_ENTIDADE = " & m_01_idSysObject & vbCrLf
    strSql = strSql & "   AND ID_ENTIDADE = " & m_02_IdEntidade & vbCrLf
    strSql = strSql & "   AND ID_ENDERECO = " & m_03_idEndereco & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudUpdate = m_03_idEndereco
    
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
'Recebe uma conexão como parâmetro (com a transação aberta)
Public Function crudDelete(ByRef cnConexao As ADODB.Connection) As Long

Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
    
    strSql = ""
    strSql = strSql & "DELETE                              " & vbCrLf
    strSql = strSql & "  FROM ENDERECO                     " & vbCrLf
    strSql = strSql & " WHERE ID_OBJECT_ENTIDADE = " & m_01_idSysObject & vbCrLf
    strSql = strSql & "   AND ID_ENTIDADE = " & m_02_IdEntidade & vbCrLf
    strSql = strSql & "   AND ID_ENDERECO = " & m_03_idEndereco & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = m_03_idEndereco

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consultaIdObjectEntidade(ByVal strTabela As String) As Long
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " select object_id from sys.objects where type='U' " & vbCrLf
    ssql = ssql & " and name = '" & strTabela & "'" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    consultaIdObjectEntidade = mRs(0)
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function recuperarEndereco(ByVal lngIdObjectID As Long, _
                                  ByVal lngIdEntidade As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM ENDERECO WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_OBJECT_ENTIDADE = " & lngIdObjectID & vbCrLf
    ssql = ssql & "   AND ID_ENTIDADE = " & lngIdEntidade & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperarEndereco = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

