VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFUNCIONARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_idFuncionario As Long
Public m_02_nome As String
Public m_03_cpf As String
Public m_04_rg As String
Public m_05_sexo As String
Public m_06_idEndereco As Long
Public m_07_telefone1 As String
Public m_08_telefone2 As String
Public m_09_telefone3 As String
Public m_10_blobFoto As String
Public m_11_idFuncao As Long
Public m_12_dataAdmissao As Date
Public m_13_situacao As String
Public m_14_observacoes As String
Public m_16_dataInclusao As String
Public m_17_usuarioInclusao As String
Public m_18_dataAlteracao As String
Public m_19_usuarioAlteracao As String
Public m_20_ctLock As Long
Public m_21_idEmpresa As Long
Public m_22_dataDemissao As Date

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String

Private oCreditoFacil As New ControladorCreditoFacil
'Recebe o recordset com endereço do funcionário
Public Function crudInsert(ByVal rsEndFun As ADODB.Recordset) As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndFun, "funcionario")

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                " & vbCrLf
    strSql = strSql & "INSERT INTO FUNCIONARIO                                                           " & vbCrLf
    strSql = strSql & "            (NOME,                                                                " & vbCrLf
    strSql = strSql & "             CPF,                                                                 " & vbCrLf
    strSql = strSql & "             RG,                                                                  " & vbCrLf
    strSql = strSql & "             SEXO,                                                                " & vbCrLf
    strSql = strSql & "             ID_ENDERECO,                                                         " & vbCrLf
    strSql = strSql & "             TELEFONE1,                                                           " & vbCrLf
    strSql = strSql & "             TELEFONE2,                                                           " & vbCrLf
    strSql = strSql & "             TELEFONE3,                                                           " & vbCrLf
    strSql = strSql & "             BLOB_FOTO,                                                           " & vbCrLf
    strSql = strSql & "             ID_FUNCAO,                                                           " & vbCrLf
    strSql = strSql & "             DATA_ADMISSAO,                                                       " & vbCrLf
    strSql = strSql & "             SITUACAO,                                                            " & vbCrLf
    strSql = strSql & "             OBSERVACOES,                                                         " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                                                       " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                                                    " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                                                      " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                                                   " & vbCrLf
    strSql = strSql & "             CT_LOCK,                                                             " & vbCrLf
    strSql = strSql & "             ID_EMPRESA)                                                          " & vbCrLf
    strSql = strSql & " VALUES                                                                           " & vbCrLf
    strSql = strSql & "        ('" & m_02_nome & "'                                                      " & vbCrLf
    strSql = strSql & "        ,'" & m_03_cpf & "'                                                       " & vbCrLf
    strSql = strSql & "        ,'" & m_04_rg & "'                                                        " & vbCrLf
    strSql = strSql & "        ,'" & m_05_sexo & "'                                                      " & vbCrLf
    strSql = strSql & "        ," & oCreditoFacil.oEndereco.crudInsert(cnConexao) & "                    " & vbCrLf
    strSql = strSql & "        ,'" & m_07_telefone1 & "'                                                 " & vbCrLf
    strSql = strSql & "        ,'" & m_08_telefone2 & "'                                                 " & vbCrLf
    strSql = strSql & "        ,'" & m_09_telefone3 & "'                                                 " & vbCrLf
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    strSql = strSql & "        ,'" & m_11_idFuncao & "'                                                  " & vbCrLf
    strSql = strSql & "        ,'" & m_12_dataAdmissao & "'                                              " & vbCrLf
    strSql = strSql & "        ,'" & m_13_situacao & "'                                                  " & vbCrLf
    strSql = strSql & "        ,'" & m_14_observacoes & "'                                               " & vbCrLf
    strSql = strSql & "        ,GetDate()                                                                " & vbCrLf
    strSql = strSql & "        ,'" & m_17_usuarioInclusao & "'                                           " & vbCrLf
    strSql = strSql & "        ,GetDate()                                                                " & vbCrLf
    strSql = strSql & "        ,'" & m_19_usuarioAlteracao & "'                                          " & vbCrLf
    strSql = strSql & "        , " & m_20_ctLock & "                                                     " & vbCrLf
    strSql = strSql & "        , " & m_21_idEmpresa & ")                                                 " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudInsert = m_01_idFuncionario
    'Se houver uma imagem, atualizar agora
    If m_10_blobFoto <> "" Then
        Call salvarImagem(m_10_blobFoto, m_01_idFuncionario, cnConexao)
    End If
    
    cnConexao.CommitTrans
    Set cnConexao = Nothing

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIDFuncionario() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT ISNULL(IDENT_CURRENT('FUNCIONARIO') + IDENT_INCR('FUNCIONARIO'),0)" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDFuncionario = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDFuncionario = 0
End Function
Public Function crudUpdate(ByVal rsEndFun As ADODB.Recordset) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndFun, "funcionario")
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                                         " & vbCrLf
    strSql = strSql & "UPDATE FUNCIONARIO                                                                                         " & vbCrLf
    strSql = strSql & " SET                                                                                                       " & vbCrLf
    strSql = strSql & "   NOME = '" & m_02_nome & "'                                                                              " & vbCrLf
    strSql = strSql & " , CPF = '" & m_03_cpf & "'                                                                                " & vbCrLf
    strSql = strSql & " , RG = '" & m_04_rg & "'                                                                                  " & vbCrLf
    strSql = strSql & " , SEXO = '" & m_05_sexo & "'                                                                              " & vbCrLf
    strSql = strSql & " , ID_ENDERECO = " & oCreditoFacil.oEndereco.crudUpdate(cnConexao) & "                                     " & vbCrLf
    strSql = strSql & " , TELEFONE1 = '" & m_07_telefone1 & "'                                                                    " & vbCrLf
    strSql = strSql & " , TELEFONE2 = '" & m_08_telefone2 & "'                                                                    " & vbCrLf
    strSql = strSql & " , TELEFONE3 = '" & m_09_telefone3 & "'                                                                    " & vbCrLf
    'Este campo pode ser null
    If m_10_blobFoto = "" Then
    strSql = strSql & " , BLOB_FOTO = null                                                                                        " & vbCrLf
    Else
        Call salvarImagem(m_10_blobFoto, m_01_idFuncionario, cnConexao)
    End If
    strSql = strSql & " , ID_FUNCAO = '" & m_11_idFuncao & "'                                                                    " & vbCrLf
    strSql = strSql & " , DATA_ADMISSAO = '" & m_12_dataAdmissao & "'                                                            " & vbCrLf
    strSql = strSql & " , SITUACAO = '" & m_13_situacao & "'                                                                     " & vbCrLf
    strSql = strSql & " , OBSERVACOES = '" & m_14_observacoes & "'                                                               " & vbCrLf
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                                                                             " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_19_usuarioAlteracao & "'                                                    " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_20_ctLock & "                                                                          " & vbCrLf
    strSql = strSql & " , ID_EMPRESA = " & m_21_idEmpresa & "                                                                    " & vbCrLf
    strSql = strSql & " , DATA_DEMISSAO = '" & m_22_dataDemissao & "'                                                            " & vbCrLf
    strSql = strSql & " WHERE ID_FUNCIONARIO = " & m_01_idFuncionario
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudUpdate = m_01_idFuncionario
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
Public Function crudDelete(ByVal rsEndFun As ADODB.Recordset) As Long

Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndFun, "funcionario")
    'Deleta o endereço do funcionário
    With oCreditoFacil.oEndereco
        .m_03_idEndereco = .crudDelete(cnConexao)
    End With
    
    strSql = ""
    strSql = strSql & "DELETE                                       " & vbCrLf
    strSql = strSql & "  FROM FUNCIONARIO                           " & vbCrLf
    strSql = strSql & " WHERE ID_FUNCIONARIO = " & m_01_idFuncionario & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = m_01_idFuncionario
    cnConexao.CommitTrans
    Set cnConexao = Nothing

Exit Function
trataerro:
  cnConexao.RollbackTrans
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
    ssql = ssql & " SELECT * FROM FUNCIONARIO WITH (NOLOCK)        " & vbCrLf
    ssql = ssql & " WHERE ID_FUNCIONARIO = '" & idFuncionario & "' " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarFuncionariosComLogin() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT F.* FROM FUNCIONARIO F WITH (NOLOCK) JOIN USUARIO U WITH (NOLOCK) ON F.ID_FUNCIONARIO = U.ID_FUNCIONARIO"

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarFuncionariosComLogin = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarFuncionariosSemLogin() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT F.* FROM FUNCIONARIO F WITH (NOLOCK) LEFT JOIN USUARIO U WITH (NOLOCK) ON F.ID_FUNCIONARIO = U.ID_FUNCIONARIO WHERE U.ID_FUNCIONARIO IS NULL"

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarFuncionariosSemLogin = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarFuncionarios() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT F.* FROM FUNCIONARIO F WITH (NOLOCK)"

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarFuncionarios = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function RecuperarNomeFuncionario(ByVal idFuncionario As Long) As String

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT NOME FROM funcionario WITH(NOLOCK) WHERE ID_FUNCIONARIO = " & idFuncionario

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    RecuperarNomeFuncionario = mRs(0)
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function RecuperarNomePeloLogin(ByVal login As String) As String

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & "select * from usuario,funcionario with(nolock)" & vbCrLf
    ssql = ssql & "Where usuario.ID_funcionario = funcionario.ID_funcionario" & vbCrLf
    ssql = ssql & "and usuario.LOGIN = '" & login & "'" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
     RecuperarNomePeloLogin = mRs("NOME")
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function FuncionarioPossuiRota(ByVal lngID_FUNCIONARIO As Long) As Boolean
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT * FROM ROTA WHERE ID_FUNCIONARIO = " & lngID_FUNCIONARIO & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    FuncionarioPossuiRota = IIf(mRs.EOF, False, True)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function carregarImagem(ByVal rs As ADODB.Recordset) As String

Dim lngImageSize As Long, lngOffset As Long, conChunkSize As Long
Dim bytChunk() As Byte
Dim strPathImagem As String

    carregarImagem = ""
    strPathImagem = App.Path & "\funcionario.jpg"
    
    If Dir(strPathImagem) <> "" Then
        Kill strPathImagem
    End If
    
    Open strPathImagem For Binary As #1
        lngImageSize = rs.Fields("BLOB_FOTO").ActualSize
        lngOffset = 0
    While lngOffset < lngImageSize
        bytChunk() = rs.Fields("BLOB_FOTO").GetChunk(lngImageSize)
        Put #1, , bytChunk()
        lngOffset = lngOffset + lngImageSize
    Wend
    
    Close #1
        
    carregarImagem = strPathImagem
    
End Function
Private Sub salvarImagem(ByVal strImg As String, ByVal lngId As Long, ByRef adoConn As ADODB.Connection)
    
    'Dim adoConn As New ADODB.Connection
    Dim adoStream As New ADODB.Stream
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    
    'adoConn.Open m_stringConexao
    
    'Adiciona img na stream
    adoStream.Type = adTypeBinary
    adoStream.Open
    adoStream.LoadFromFile strImg
    If InStr(1, strImg, App.Path) > 0 Then Kill strImg
    
    'Instrução que recupera o registro que conterá a imagem
    strSql = " SELECT * FROM funcionario WHERE ID_FUNCIONARIO = " & lngId
    
    rs.Open strSql, adoConn, adOpenKeyset, adLockOptimistic
    rs.Fields("BLOB_FOTO") = adoStream.Read
    rs.Update
    rs.Close
    adoStream.Close
    'adoConn.Close
    
End Sub
Private Sub moveEnderecoParaObjeto(ByVal rs As ADODB.Recordset, ByVal tabelaPai As String)

    With oCreditoFacil.oEndereco
        
        .m_stringConexao = m_stringConexao
        .m_timeOut = m_timeOut
        
        Do While Not rs.EOF
            'Instanciando Endereco
            .m_01_idSysObject = .consultaIdObjectEntidade(tabelaPai)
            .m_02_IdEntidade = m_01_idFuncionario
            .m_03_idEndereco = rs.Fields("ID_ENDERECO")
            .m_04_tipoLogradouro = rs.Fields("TIPO_LOGRADOURO")
            .m_05_logradouro = rs.Fields("LOGRADOURO")
            .m_06_numero = rs.Fields("NUMERO")
            .m_07_complemento = rs.Fields("COMPLEMENTO")
            .m_08_pontoReferencia = rs.Fields("PONTO_REFERENCIA")
            .m_09_cep = rs.Fields("CEP")
            .m_10_idBairro = rs.Fields("ID_BAIRRO")
            .m_11_idMunicipio = rs.Fields("ID_MUNICIPIO")
            .m_12_idEstado = rs.Fields("ID_ESTADO")
            .m_13_usuarioInclusao = rs.Fields("USUARIO_INCLUSAO")
            .m_14_dataInclusao = rs.Fields("DATA_INCLUSAO")
            .m_15_usuarioAlteracao = rs.Fields("USUARIO_ALTERACAO")
            .m_16_dataAlteracao = rs.Fields("DATA_ALTERACAO")
            .m_17_ctLock = rs.Fields("CT_LOCK")
            rs.MoveNext
        Loop
    End With

End Sub

