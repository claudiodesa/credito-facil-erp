VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEMPRESACLIENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Atributos de Classe
Public m_01_idEmpresaCliente As Long
Public m_02_tipo As String
Public m_03_idRamo As Long
Public m_04_iniciouAtividade As Date
Public m_05_vendaDiaria As Double
Public m_06_idEndereco As Long
Public m_07_cgc As String
Public m_08_razaoSocial As String
Public m_09_nomeFantasia As String
Public m_10_cpf As String
Public m_11_nomePessoaFisica As String
Public m_12_blobLogoEmpresa As String
Public m_13_dataInclusao As String
Public m_14_usuarioInclusao As String
Public m_15_dataAlteracao As String
Public m_16_usuarioAlteracao As String
Public m_17_ctLock As Long

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String

Private oEnderecoEmpresa As New clsENDERECO
Private oResponsavel As New clsResponsavel

'Recebe:
'recordset contendo o endereço da empresaCliente
'recordset contendo o responsavelFinanceiro da empresaCliente
'recordset contendo o endereço do responsavelFinanceiro da empresaCliente
Public Function crudInsert(ByVal rsEndEmpCli, ByVal rsResFin, ByVal rsEndResFin As ADODB.Recordset) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndEmpCli, "empresaCliente")

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                " & vbCrLf
    strSql = strSql & "INSERT INTO EMPRESACLIENTE                                                        " & vbCrLf
    strSql = strSql & "            (TIPO,                                                                " & vbCrLf
    strSql = strSql & "             ID_RAMO,                                                             " & vbCrLf
    strSql = strSql & "             INICIOU_ATIVIDADE,                                                   " & vbCrLf
    strSql = strSql & "             VENDA_DIARIA,                                                        " & vbCrLf
    strSql = strSql & "             ID_ENDERECO,                                                         " & vbCrLf
    strSql = strSql & "             CGC,                                                                 " & vbCrLf
    strSql = strSql & "             RAZAO_SOCIAL,                                                        " & vbCrLf
    strSql = strSql & "             NOME_FANTASIA,                                                       " & vbCrLf
    strSql = strSql & "             CPF,                                                                 " & vbCrLf
    strSql = strSql & "             NOME_PESSOA_FISICA,                                                  " & vbCrLf
    strSql = strSql & "             BLOB_LOGO_EMPRESA,                                                   " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                                                       " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                                                    " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                                                      " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                                                   " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                                             " & vbCrLf
    strSql = strSql & " VALUES                                                                           " & vbCrLf
    strSql = strSql & "        ('" & m_02_tipo & "'                                                      " & vbCrLf
    strSql = strSql & "        ,'" & m_03_idRamo & "'                                                    " & vbCrLf
    strSql = strSql & "        ,'" & m_04_iniciouAtividade & "'                                          " & vbCrLf
    strSql = strSql & "        , " & Replace(m_05_vendaDiaria, ",", ".") & "                             " & vbCrLf
    strSql = strSql & "        , " & oEnderecoEmpresa.crudInsert(cnConexao) & "                          " & vbCrLf
    'Caso este campo venha a ser null
    If Len(m_07_cgc) > 0 Then
    strSql = strSql & "        ,'" & m_07_cgc & "'                                                       " & vbCrLf
    Else
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_08_razaoSocial) > 0 Then
    strSql = strSql & "        ,'" & m_08_razaoSocial & "'                                               " & vbCrLf
    Else
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_09_nomeFantasia) > 0 Then
    strSql = strSql & "        ,'" & m_09_nomeFantasia & "'                                              " & vbCrLf
    Else
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_10_cpf) > 0 Then
    strSql = strSql & "        ,'" & m_10_cpf & "'                                                       " & vbCrLf
    Else
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_11_nomePessoaFisica) > 0 Then
    strSql = strSql & "        ,'" & m_11_nomePessoaFisica & "'                                          " & vbCrLf
    Else
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    End If
    strSql = strSql & "        ,null                                                                     " & vbCrLf
    strSql = strSql & "        ,GetDate()                                                                " & vbCrLf
    strSql = strSql & "        ,'" & m_14_usuarioInclusao & "'                                           " & vbCrLf
    strSql = strSql & "        ,GetDate()                                                                " & vbCrLf
    strSql = strSql & "        ,'" & m_16_usuarioAlteracao & "'                                          " & vbCrLf
    strSql = strSql & "        , " & m_17_ctLock & ")                                                    " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
        
    'Insere o resposavelFinanceiro da empresaCliente
    Call moveResponsavelParaObjeto(rsResFin)
    With oResponsavel
        .m_01_idResponsavel = .crudInsert(rsEndResFin, cnConexao)
    End With
    
    crudInsert = m_01_idEmpresaCliente
    cnConexao.CommitTrans
    Set cnConexao = Nothing

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function getNovoIdEmpresaCliente() As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT ISNULL(IDENT_CURRENT('empresaCliente') + IDENT_INCR('empresaCliente'),0)" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    getNovoIdEmpresaCliente = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  getNovoIdEmpresaCliente = 0
End Function
'Recebe:
'recordset contendo o endereço da empresaCliente
'recordset contendo o responsavelFinanceiro da empresaCliente
'recordset contendo o endereço do responsavelFinanceiro da empresaCliente
Public Function crudUpdate(ByVal rsEndEmpCli, ByVal rsResFin, ByVal rsEndResFin As ADODB.Recordset) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndEmpCli, "empresaCliente")
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                              " & vbCrLf
    strSql = strSql & "UPDATE empresaCliente                                                                           " & vbCrLf
    strSql = strSql & " SET                                                                                            " & vbCrLf
    strSql = strSql & "   TIPO = '" & m_02_tipo & "'                                                                   " & vbCrLf
    strSql = strSql & " , ID_RAMO = '" & m_03_idRamo & "'                                                              " & vbCrLf
    strSql = strSql & " , INICIOU_ATIVIDADE = '" & m_04_iniciouAtividade & "'                                          " & vbCrLf
    strSql = strSql & " , VENDA_DIARIA = " & Replace(m_05_vendaDiaria, ",", ".") & "                                   " & vbCrLf
    strSql = strSql & " , ID_ENDERECO = '" & oEnderecoEmpresa.crudUpdate(cnConexao) & "'                               " & vbCrLf
    strSql = strSql & " , CGC = '" & m_07_cgc & "'                                                                     " & vbCrLf
    strSql = strSql & " , RAZAO_SOCIAL = '" & m_08_razaoSocial & "'                                                    " & vbCrLf
    strSql = strSql & " , NOME_FANTASIA = '" & m_09_nomeFantasia & "'                                                  " & vbCrLf
    strSql = strSql & " , CPF = '" & m_10_cpf & "'                                                                     " & vbCrLf
    strSql = strSql & " , NOME_PESSOA_FISICA = '" & m_11_nomePessoaFisica & "'                                         " & vbCrLf
    'Este campo pode ser null
    If m_12_blobLogoEmpresa = "" Then
    strSql = strSql & " , BLOB_LOGO_EMPRESA = null                                                                     " & vbCrLf
    Else
        Call salvarImagem(m_12_blobLogoEmpresa, m_01_idEmpresaCliente, cnConexao)
    End If
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                                                                   " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_16_usuarioAlteracao & "'                                          " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_17_ctLock & "                                                                " & vbCrLf
    strSql = strSql & " WHERE ID_EMPRESACLIENTE = " & m_01_idEmpresaCliente
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudUpdate = m_01_idEmpresaCliente
    
    'Atualiza o Responsavel
    Call moveResponsavelParaObjeto(rsResFin)
    With oResponsavel
        .m_01_idResponsavel = .crudUpdate(rsEndResFin, cnConexao)
    End With
    
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
'Recebe:
'recordset contendo o endereço da empresaCliente
'recordset contendo o responsavelFinanceiro da empresaCliente
'recordset contendo o endereço do responsavelFinanceiro da empresaCliente
Public Function crudDelete(ByVal rsEndEmpCli As ADODB.Recordset, _
                           ByVal rsResFin As ADODB.Recordset, _
                           ByVal rsEndResFin As ADODB.Recordset) As Long

Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    Call moveEnderecoParaObjeto(rsEndEmpCli, "empresaCliente")
    'Exclui o endereco da empresaCliente
    With oEnderecoEmpresa
        .m_03_idEndereco = .crudDelete(cnConexao)
    End With
    
    Call moveResponsavelParaObjeto(rsResFin)
    'Exclui o responsavelFinanceiro da empresaCliente e seu respectivo endereço
    With oResponsavel
        .m_01_idResponsavel = .crudDelete(rsEndResFin, cnConexao)
    End With
    
    strSql = ""
    strSql = strSql & "DELETE                                             " & vbCrLf
    strSql = strSql & "  FROM empresaCliente                              " & vbCrLf
    strSql = strSql & " WHERE ID_EMPRESACLIENTE = " & m_01_idEmpresaCliente & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = m_01_idEmpresaCliente
    cnConexao.CommitTrans
    Set cnConexao = Nothing

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
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " select" & vbCrLf
    ssql = ssql & " Case TIPO"
    ssql = ssql & "     when 'F' then NOME_PESSOA_FISICA"
    ssql = ssql & "     when 'J' then NOME_FANTASIA"
    ssql = ssql & " end as NOME,* FROM empresaCliente WITH (NOLOCK)"
    ssql = ssql & " WHERE ID_EMPRESACLIENTE = " & idEmpresaCliente & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function recuperarEmpresasCliente() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " select" & vbCrLf
    ssql = ssql & " Case TIPO"
    ssql = ssql & "     when 'F' then NOME_PESSOA_FISICA"
    ssql = ssql & "     when 'J' then NOME_FANTASIA"
    ssql = ssql & " end as NOME,*"
    ssql = ssql & " FROM empresaCliente WITH (NOLOCK) "
    ssql = ssql & " order by NOME "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperarEmpresasCliente = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function getNomeEmpresaCliente(ByVal idEmpresaCliente As Long) As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " select" & vbCrLf
    ssql = ssql & " Case TIPO"
    ssql = ssql & "     when 'F' then NOME_PESSOA_FISICA"
    ssql = ssql & "     when 'J' then NOME_FANTASIA"
    ssql = ssql & " end as NOME"
    ssql = ssql & " FROM empresaCliente WITH (NOLOCK) "
    ssql = ssql & " WHERE ID_EMPRESACLIENTE = " & idEmpresaCliente

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    getNomeEmpresaCliente = mRs("NOME")
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Private Sub moveEnderecoParaObjeto(ByVal rs As ADODB.Recordset, ByRef tabelaPai As String)

    With oEnderecoEmpresa
        
        .m_stringConexao = m_stringConexao
        .m_timeOut = m_timeOut
        
        Do While Not rs.EOF
            'Instanciando Endereco
            .m_01_idSysObject = .consultaIdObjectEntidade(tabelaPai)
            .m_02_IdEntidade = m_01_idEmpresaCliente
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
Private Sub moveResponsavelParaObjeto(ByVal rs As ADODB.Recordset)

    With oResponsavel
        .inicializaResFin
        .m_stringConexao = m_stringConexao
        .m_timeOut = m_timeOut
        
        Do While Not rs.EOF
            'Instanciando responsavelFinanceiro
            .m_01_idResponsavel = rs.Fields("ID_RESPONSAVEL")
            .m_02_idEmpresaCliente = m_01_idEmpresaCliente
            .m_03_situacao = rs.Fields("SITUACAO")
            .m_04_nome = rs.Fields("NOME")
            .m_05_cpf = rs.Fields("CPF")
            .m_06_rg = rs.Fields("RG")
            .m_07_orgaoEmissor = rs.Fields("ORGAO_EMISSOR")
            .m_08_dataExpedicao = rs.Fields("DATA_EXPEDICAO")
            .m_09_sexo = rs.Fields("SEXO")
            .m_10_dataNascimento = rs.Fields("DATA_NASCIMENTO")
            .m_11_estadoCivil = rs.Fields("ESTADO_CIVIL")
            .m_12_naturalidade = rs.Fields("NATURALIDADE")
            .m_13_nacionalidade = rs.Fields("NACIONALIDADE")
            .m_14_nomeMae = rs.Fields("NOME_MAE")
            .m_15_idEndereco = rs.Fields("ID_ENDERECO")
            .m_16_tipoImovel = rs.Fields("TIPO_IMOVEL")
            .m_17_resideDesde = rs.Fields("RESIDE_DESDE")
            .m_18_telefone1 = rs.Fields("TELEFONE1")
            .m_19_telefone2 = rs.Fields("TELEFONE2")
            .m_20_telefone3 = rs.Fields("TELEFONE3")
            .m_21_contatoReferencia1 = rs.Fields("CONTATO_REFERENCIA1")
            .m_22_telefoneReferencia1 = rs.Fields("TELEFONE_REFERENCIA1")
            .m_23_contatoReferencia2 = rs.Fields("CONTATO_REFERENCIA2")
            .m_24_telefoneReferencia2 = rs.Fields("TELEFONE_REFERENCIA1")
            .m_25_indicadoPor = rs.Fields("INDICADO_POR")
            .m_26_dataInclusao = rs.Fields("DATA_INCLUSAO")
            .m_27_usuarioInclusao = rs.Fields("USUARIO_INCLUSAO")
            .m_28_dataAlteracao = rs.Fields("DATA_ALTERACAO")
            .m_29_usuarioAlteracao = rs.Fields("USUARIO_ALTERACAO")
            .m_30_ctLock = rs.Fields("CT_LOCK")
            rs.MoveNext
        Loop
    End With

End Sub
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
    strSql = " SELECT * FROM empresaCliente WHERE ID_EMPRESACLIENTE = " & lngId
    
    rs.Open strSql, adoConn, adOpenKeyset, adLockOptimistic
    rs.Fields("BLOB_LOGO_EMPRESA") = adoStream.Read
    rs.Update
    rs.Close
    adoStream.Close
    'adoConn.Close
    
End Sub
Public Function carregarImagem(ByVal rs As ADODB.Recordset) As String

Dim lngImageSize As Long, lngOffset As Long, conChunkSize As Long
Dim bytChunk() As Byte
Dim strPathImagem As String

    carregarImagem = ""
    strPathImagem = App.Path & "\empresaCliente.jpg"
    
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

Public Function recuperarEmpresasDevedorasPorRota(ByVal idRota As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = "" & vbCrLf
    ssql = ssql & " select distinct " & vbCrLf
    ssql = ssql & " case when TIPO = 'J' then empresaCliente.NOME_FANTASIA" & vbCrLf
    ssql = ssql & "      when TIPO = 'F' then empresaCliente.NOME_PESSOA_FISICA" & vbCrLf
    ssql = ssql & " end as NOME," & vbCrLf
    ssql = ssql & " empresaCliente.*" & vbCrLf
    ssql = ssql & " From financiamento, financiamentoParcela, rota, empresaCliente"
    ssql = ssql & " Where financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " and financiamento.ID_ROTA = rota.ID_ROTA" & vbCrLf
    ssql = ssql & " and empresaCliente.ID_EMPRESACLIENTE = financiamento.ID_EMPRESACLIENTE" & vbCrLf
    ssql = ssql & " and financiamento.ID_ROTA = " & idRota & vbCrLf
    ssql = ssql & " and financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperarEmpresasDevedorasPorRota = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

