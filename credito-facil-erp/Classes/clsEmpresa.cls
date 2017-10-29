VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_idEmpresa As Long
Public m_02_cgcEmpresa As String
Public m_03_razaoSocial As String
Public m_04_nomeFantasia As String
Public m_05_idEndereco As Long
Public m_06_telefone1 As String
Public m_07_telefone2 As String
Public m_08_telefone3 As String
Public m_09_blobLogEmpresa As String
Public m_10_dataInclusao As String
Public m_11_usuarioInclusao As String
Public m_12_dataAlteracao As String
Public m_13_usuarioAlteracao As String
Public m_14_ctLock As Long

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String

Private oCreditoFacil As New ControladorCreditoFacil

Public Function crudInsert(ByVal rsEndEmp As ADODB.Recordset) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndEmp, "empresa")

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                " & vbCrLf
    strSql = strSql & "INSERT INTO EMPRESA                                                               " & vbCrLf
    strSql = strSql & "            (CGC_EMPRESA,                                                         " & vbCrLf
    strSql = strSql & "             RAZAO_SOCIAL,                                                        " & vbCrLf
    strSql = strSql & "             NOME_FANTASIA,                                                       " & vbCrLf
    strSql = strSql & "             ID_ENDERECO,                                                         " & vbCrLf
    strSql = strSql & "             TELEFONE1,                                                           " & vbCrLf
    strSql = strSql & "             TELEFONE2,                                                           " & vbCrLf
    strSql = strSql & "             TELEFONE3,                                                           " & vbCrLf
    strSql = strSql & "             BLOB_LOGO_EMPRESA,                                                   " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                                                       " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                                                    " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                                                      " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                                                   " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                                             " & vbCrLf
    strSql = strSql & " VALUES                                                                           " & vbCrLf
    strSql = strSql & "        ('" & m_02_cgcEmpresa & "'                                                " & vbCrLf
    strSql = strSql & "        ,'" & m_03_razaoSocial & "'                                               " & vbCrLf
    strSql = strSql & "        ,'" & m_04_nomeFantasia & "'                                              " & vbCrLf
    strSql = strSql & "        , " & oCreditoFacil.oEndereco.crudInsert(cnConexao) & "                   " & vbCrLf
    strSql = strSql & "        ,'" & m_06_telefone1 & "'                                                 " & vbCrLf
    strSql = strSql & "        ,'" & m_07_telefone2 & "'                                                 " & vbCrLf
    strSql = strSql & "        ,'" & m_08_telefone3 & "'                                                 " & vbCrLf
    strSql = strSql & "        , null                                                                    " & vbCrLf
    strSql = strSql & "        ,'" & m_10_dataInclusao & "'                                              " & vbCrLf
    strSql = strSql & "        ,'" & m_11_usuarioInclusao & "'                                           " & vbCrLf
    strSql = strSql & "        ,'" & m_12_dataAlteracao & "'                                             " & vbCrLf
    strSql = strSql & "        ,'" & m_13_usuarioAlteracao & "'                                          " & vbCrLf
    strSql = strSql & "        , " & m_14_ctLock & ")                                                    " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudInsert = m_01_idEmpresa
    'Se houver uma imagem, atualizar agora
    If m_09_blobLogEmpresa <> "" Then
        Call salvarImagem(m_09_blobLogEmpresa, m_01_idEmpresa, cnConexao)
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
Public Function getNovoIdEmpresa() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT ISNULL(IDENT_CURRENT('EMPRESA') + IDENT_INCR('EMPRESA'),0)" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    getNovoIdEmpresa = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  getNovoIdEmpresa = 0
End Function
Public Function crudUpdate(ByVal rsEndEmp As ADODB.Recordset) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndEmp, "empresa")
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                              " & vbCrLf
    strSql = strSql & "UPDATE EMPRESA                                                                                  " & vbCrLf
    strSql = strSql & " SET                                                                                            " & vbCrLf
    strSql = strSql & "   CGC_EMPRESA = '" & m_02_cgcEmpresa & "'                                                      " & vbCrLf
    strSql = strSql & " , RAZAO_SOCIAL = '" & m_03_razaoSocial & "'                                                    " & vbCrLf
    strSql = strSql & " , NOME_FANTASIA = '" & m_04_nomeFantasia & "'                                                  " & vbCrLf
    strSql = strSql & " , ID_ENDERECO = " & oCreditoFacil.oEndereco.crudUpdate(cnConexao) & "                          " & vbCrLf
    strSql = strSql & " , TELEFONE1 = '" & m_06_telefone1 & "'                                                         " & vbCrLf
    strSql = strSql & " , TELEFONE2 = '" & m_07_telefone2 & "'                                                         " & vbCrLf
    strSql = strSql & " , TELEFONE3 = '" & m_08_telefone3 & "'                                                         " & vbCrLf
    'Este campo pode ser null
    If m_09_blobLogEmpresa = "" Then
    strSql = strSql & " , BLOB_LOGO_EMPRESA = null                                                                     " & vbCrLf
    Else
        Call salvarImagem(m_09_blobLogEmpresa, m_01_idEmpresa, cnConexao)
    End If
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                                                                   " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_13_usuarioAlteracao & "'                                          " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_14_ctLock & "                                                                " & vbCrLf
    strSql = strSql & " WHERE ID_EMPRESA = " & m_01_idEmpresa
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudUpdate = m_01_idEmpresa
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
Public Function crudDelete(ByVal rsEndEmp As ADODB.Recordset) As Long

Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndEmp, "empresa")
    'Deleta o endereco da empresa
    With oCreditoFacil.oEndereco
        m_05_idEndereco = .crudDelete(cnConexao)
    End With
    
    strSql = ""
    strSql = strSql & "DELETE                               " & vbCrLf
    strSql = strSql & "  FROM EMPRESA                       " & vbCrLf
    strSql = strSql & " WHERE ID_EMPRESA = " & m_01_idEmpresa & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = m_01_idEmpresa
    cnConexao.CommitTrans
    Set cnConexao = Nothing

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consulta(ByVal idEmpresa As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM EMPRESA WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_EMPRESA = '" & idEmpresa & "' " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function recuperarEmpresas() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT * FROM EMPRESA WITH (NOLOCK) "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperarEmpresas = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function listagemEmpresasParaRelatorio() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & "select " & Chr(13)
    ssql = ssql & "case" & Chr(13)
    ssql = ssql & "   when eC.NOME_FANTASIA ='' then eC.NOME_PESSOA_FISICA else eC.NOME_FANTASIA" & Chr(13)
    ssql = ssql & "end as EMPRESA, rF.NOME, rF.TELEFONE1, isnull(rF.TELEFONE2,'') as TELEFONE2, isnull(rF.TELEFONE3,'') as TELEFONE3" & Chr(13)
    ssql = ssql & "from empresaCliente eC join responsavelFinanceiro rF on eC.ID_EMPRESACLIENTE = rF.ID_EMPRESACLIENTE" & Chr(13)
    ssql = ssql & "order by EMPRESA "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set listagemEmpresasParaRelatorio = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Private Sub moveEnderecoParaObjeto(ByVal rs As ADODB.Recordset, ByVal tabelaPai As String)

    With oCreditoFacil.oEndereco
        
        .m_stringConexao = m_stringConexao
        .m_timeOut = m_timeOut
        
        Do While Not rs.EOF
            'Instanciando Endereco
            .m_01_idSysObject = .consultaIdObjectEntidade(tabelaPai)
            .m_02_IdEntidade = m_01_idEmpresa
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
Public Function carregarImagem(ByVal rs As ADODB.Recordset) As String

Dim lngImageSize As Long, lngOffset As Long, conChunkSize As Long
Dim bytChunk() As Byte
Dim strPathImagem As String

    carregarImagem = ""
    strPathImagem = App.Path & "\empresa.jpg"
    
    If Dir(strPathImagem) <> "" Then
        Kill strPathImagem
    End If
    
    Open strPathImagem For Binary As #1
        lngImageSize = rs.Fields("BLOB_LOGO_EMPRESA").ActualSize
        lngOffset = 0
    While lngOffset < lngImageSize
        bytChunk() = rs.Fields("BLOB_LOGO_EMPRESA").GetChunk(lngImageSize)
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
    strSql = " SELECT * FROM EMPRESA WHERE ID_EMPRESA = " & lngId
    
    rs.Open strSql, adoConn, adOpenKeyset, adLockOptimistic
    rs.Fields("BLOB_LOGO_EMPRESA") = adoStream.Read
    rs.Update
    rs.Close
    adoStream.Close
    'adoConn.Close
    
End Sub

