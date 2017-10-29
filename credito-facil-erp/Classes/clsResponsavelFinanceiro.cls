VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResponsavel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Depende de clsEmpresaCliente
Option Explicit

'Atributos de Classe
Public m_01_idResponsavel As Long
Public m_02_idEmpresaCliente As Long
Public m_03_situacao As String
Public m_04_nome As String
Public m_05_cpf As String
Public m_06_rg As String
Public m_07_orgaoEmissor As String
Public m_08_dataExpedicao As Date
Public m_09_sexo As String
Public m_10_dataNascimento As Date
Public m_11_estadoCivil As String
Public m_12_naturalidade As String
Public m_13_nacionalidade As String
Public m_14_nomeMae As String
Public m_15_idEndereco As Long
Public m_16_tipoImovel As String
Public m_17_resideDesde As Date
Public m_18_telefone1 As String
Public m_19_telefone2 As String
Public m_20_telefone3 As String
Public m_21_contatoReferencia1 As String
Public m_22_telefoneReferencia1 As String
Public m_23_contatoReferencia2 As String
Public m_24_telefoneReferencia2 As String
Public m_25_indicadoPor As String
Public m_26_dataInclusao As String
Public m_27_usuarioInclusao As String
Public m_28_dataAlteracao As String
Public m_29_usuarioAlteracao As String
Public m_30_ctLock As String

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String

Public rsResFin As ADODB.Recordset

Private oEndereco As New clsENDERECO
'Recebe uma conexão como parâmetro (com a transação aberta)
Public Function crudInsert(ByVal rsEndResFin As ADODB.Recordset, _
                           ByRef cnConexao As ADODB.Connection) As Long

Dim mRs         As ADODB.Recordset
Dim strSql      As String

    On Error GoTo trataerro
    
    Call moveEnderecoParaObjeto(rsEndResFin, "responsavelFinanceiro")

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                " & vbCrLf
    strSql = strSql & "INSERT INTO RESPONSAVELFINANCEIRO                                                 " & vbCrLf
    strSql = strSql & "            (ID_EMPRESACLIENTE,                                                   " & vbCrLf
    strSql = strSql & "             SITUACAO,                                                            " & vbCrLf
    strSql = strSql & "             NOME,                                                                " & vbCrLf
    strSql = strSql & "             CPF,                                                                 " & vbCrLf
    strSql = strSql & "             RG,                                                                  " & vbCrLf
    strSql = strSql & "             ORGAO_EMISSOR,                                                       " & vbCrLf
    strSql = strSql & "             DATA_EXPEDICAO,                                                      " & vbCrLf
    strSql = strSql & "             SEXO,                                                                " & vbCrLf
    strSql = strSql & "             DATA_NASCIMENTO,                                                     " & vbCrLf
    strSql = strSql & "             ESTADO_CIVIL,                                                        " & vbCrLf
    strSql = strSql & "             NATURALIDADE,                                                        " & vbCrLf
    strSql = strSql & "             NACIONALIDADE,                                                       " & vbCrLf
    strSql = strSql & "             NOME_MAE,                                                            " & vbCrLf
    strSql = strSql & "             ID_ENDERECO,                                                         " & vbCrLf
    strSql = strSql & "             TIPO_IMOVEL,                                                         " & vbCrLf
    strSql = strSql & "             RESIDE_DESDE,                                                        " & vbCrLf
    strSql = strSql & "             TELEFONE1,                                                           " & vbCrLf
    strSql = strSql & "             TELEFONE2,                                                           " & vbCrLf
    strSql = strSql & "             TELEFONE3,                                                           " & vbCrLf
    strSql = strSql & "             CONTATO_REFERENCIA1,                                                 " & vbCrLf
    strSql = strSql & "             TELEFONE_REFERENCIA1,                                                " & vbCrLf
    strSql = strSql & "             CONTATO_REFERENCIA2,                                                 " & vbCrLf
    strSql = strSql & "             TELEFONE_REFERENCIA2,                                                " & vbCrLf
    strSql = strSql & "             INDICADO_POR,                                                        " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                                                       " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                                                    " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                                                      " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                                                   " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                                             " & vbCrLf
    strSql = strSql & " VALUES                                                                           " & vbCrLf
    strSql = strSql & "        ('" & m_02_idEmpresaCliente & "'                                          " & vbCrLf
    strSql = strSql & "        ,'" & m_03_situacao & "'                                                  " & vbCrLf
    strSql = strSql & "        ,'" & m_04_nome & "'                                                      " & vbCrLf
    strSql = strSql & "        ,'" & m_05_cpf & "'                                                       " & vbCrLf
    strSql = strSql & "        ,'" & m_06_rg & "'                                                        " & vbCrLf
    strSql = strSql & "        ,'" & m_07_orgaoEmissor & "'                                              " & vbCrLf
    strSql = strSql & "        ,'" & m_08_dataExpedicao & "'                                             " & vbCrLf
    strSql = strSql & "        ,'" & m_09_sexo & "'                                                      " & vbCrLf
    strSql = strSql & "        ,'" & m_10_dataNascimento & "'                                            " & vbCrLf
    strSql = strSql & "        ,'" & m_11_estadoCivil & "'                                               " & vbCrLf
    strSql = strSql & "        ,'" & m_12_naturalidade & "'                                              " & vbCrLf
    strSql = strSql & "        ,'" & m_13_nacionalidade & "'                                             " & vbCrLf
    strSql = strSql & "        ,'" & m_14_nomeMae & "'                                                   " & vbCrLf
    If m_15_idEndereco = 0 Then
        m_15_idEndereco = oEndereco.crudInsert(cnConexao)
    End If
    strSql = strSql & "        , " & m_15_idEndereco & "                                                 " & vbCrLf
    strSql = strSql & "        ,'" & m_16_tipoImovel & "'                                                " & vbCrLf
    strSql = strSql & "        ,'" & m_17_resideDesde & "'                                               " & vbCrLf
    If Len(m_18_telefone1) > 0 Then
    strSql = strSql & "        ,'" & m_18_telefone1 & "'                                                 " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_19_telefone2) > 0 Then
    strSql = strSql & "        ,'" & m_19_telefone2 & "'                                                 " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_20_telefone3) > 0 Then
    strSql = strSql & "        ,'" & m_20_telefone3 & "'                                                 " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_21_contatoReferencia1) > 0 Then
    strSql = strSql & "        ,'" & m_21_contatoReferencia1 & "'                                        " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_22_telefoneReferencia1) > 0 Then
    strSql = strSql & "        ,'" & m_22_telefoneReferencia1 & "'                                       " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_23_contatoReferencia2) > 0 Then
    strSql = strSql & "        ,'" & m_23_contatoReferencia2 & "'                                        " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_24_telefoneReferencia2) > 0 Then
    strSql = strSql & "        ,'" & m_24_telefoneReferencia2 & "'                                       " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    If Len(m_25_indicadoPor) > 0 Then
    strSql = strSql & "        ,'" & m_25_indicadoPor & "'                                               " & vbCrLf
    Else
    strSql = strSql & "        , null                                                                    " & vbCrLf
    End If
    strSql = strSql & "        , getDate()                                                               " & vbCrLf
    strSql = strSql & "        ,'" & m_27_usuarioInclusao & "'                                           " & vbCrLf
    strSql = strSql & "        , getDate()                                                               " & vbCrLf
    strSql = strSql & "        ,'" & m_29_usuarioAlteracao & "'                                          " & vbCrLf
    strSql = strSql & "        , " & m_30_ctLock & ")                                                    " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    strSql = "SELECT ID_RESPONSAVEL FROM responsavelFinanceiro WITH(NOLOCK) WHERE ID_EMPRESACLIENTE = " & m_02_idEmpresaCliente
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    m_01_idResponsavel = mRs(0)
    crudInsert = m_01_idResponsavel
    Set mRs.ActiveConnection = Nothing
    
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Private Sub moveEnderecoParaObjeto(ByVal rs As ADODB.Recordset, tabelaPai As String)

    With oEndereco
        
        .m_stringConexao = m_stringConexao
        .m_timeOut = m_timeOut
        
        Do While Not rs.EOF
            'Instanciando Endereco
            .m_01_idSysObject = .consultaIdObjectEntidade(tabelaPai)
            If rs.Fields("ID_ENTIDADE") = 0 Then
                .m_02_IdEntidade = getNovoIdResponsavel()
            Else
                .m_02_IdEntidade = rs.Fields("ID_ENTIDADE")
            End If
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
'Recebe uma conexão como parâmetro (com a transação aberta)
Public Function crudUpdate(ByVal rsEndResFin As ADODB.Recordset, ByRef cnConexao As ADODB.Connection) As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String

    On Error GoTo trataerro
    
    Call moveEnderecoParaObjeto(rsEndResFin, "responsavelFinanceiro")
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                                                   " & vbCrLf
    strSql = strSql & "UPDATE responsavelFinanceiro                                                         " & vbCrLf
    strSql = strSql & " SET                                                                                 " & vbCrLf
    strSql = strSql & "   SITUACAO = '" & m_03_situacao & "'                                                " & vbCrLf
    strSql = strSql & " , NOME = '" & m_04_nome & "'                                                        " & vbCrLf
    strSql = strSql & " , CPF = '" & m_05_cpf & "'                                                          " & vbCrLf
    strSql = strSql & " , RG = '" & m_06_rg & "'                                                            " & vbCrLf
    strSql = strSql & " , ORGAO_EMISSOR = '" & m_07_orgaoEmissor & "'                                       " & vbCrLf
    strSql = strSql & " , DATA_EXPEDICAO = '" & m_08_dataExpedicao & "'                                     " & vbCrLf
    strSql = strSql & " , SEXO = '" & m_09_sexo & "'                                                        " & vbCrLf
    strSql = strSql & " , DATA_NASCIMENTO = '" & m_10_dataNascimento & "'                                   " & vbCrLf
    strSql = strSql & " , ESTADO_CIVIL = '" & m_11_estadoCivil & "'                                         " & vbCrLf
    strSql = strSql & " , NATURALIDADE = '" & m_12_naturalidade & "'                                        " & vbCrLf
    strSql = strSql & " , NACIONALIDADE = '" & m_13_nacionalidade & "'                                      " & vbCrLf
    strSql = strSql & " , NOME_MAE = '" & m_14_nomeMae & "'                                                 " & vbCrLf
    strSql = strSql & " , ID_ENDERECO = '" & oEndereco.crudUpdate(cnConexao) & "'                           " & vbCrLf
    strSql = strSql & " , TIPO_IMOVEL = '" & m_16_tipoImovel & "'                                           " & vbCrLf
    strSql = strSql & " , RESIDE_DESDE = '" & m_17_resideDesde & "'                                         " & vbCrLf
    'Caso este campo venha a ser null
    If Len(m_18_telefone1) > 0 Then
    strSql = strSql & " , TELEFONE1 = '" & m_18_telefone1 & "'                                              " & vbCrLf
    Else
    strSql = strSql & " , TELEFONE1 = NULL                                                                  " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_19_telefone2) > 0 Then
    strSql = strSql & " , TELEFONE2 = '" & m_19_telefone2 & "'                                              " & vbCrLf
    Else
    strSql = strSql & " , TELEFONE2 = NULL                                                                  " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_20_telefone3) > 0 Then
    strSql = strSql & " , TELEFONE3 = '" & m_20_telefone3 & "'                                              " & vbCrLf
    Else
    strSql = strSql & " , TELEFONE3 = NULL                                                                  " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_21_contatoReferencia1) > 0 Then
    strSql = strSql & " , CONTATO_REFERENCIA1 = '" & m_21_contatoReferencia1 & "'                           " & vbCrLf
    Else
    strSql = strSql & " , CONTATO_REFERENCIA1 = NULL                                                        " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_22_telefoneReferencia1) > 0 Then
    strSql = strSql & " , TELEFONE_REFERENCIA1 = '" & m_22_telefoneReferencia1 & "'                         " & vbCrLf
    Else
    strSql = strSql & " , TELEFONE_REFERENCIA1 = NULL                                                       " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_23_contatoReferencia2) > 0 Then
    strSql = strSql & " , CONTATO_REFERENCIA2 = '" & m_23_contatoReferencia2 & "'                           " & vbCrLf
    Else
    strSql = strSql & " , CONTATO_REFERENCIA2 = NULL                                                        " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_24_telefoneReferencia2) > 0 Then
    strSql = strSql & " , TELEFONE_REFERENCIA2 = '" & m_24_telefoneReferencia2 & "'                         " & vbCrLf
    Else
    strSql = strSql & " , TELEFONE_REFERENCIA2 = NULL                                                       " & vbCrLf
    End If
    'Caso este campo venha a ser null
    If Len(m_25_indicadoPor) > 0 Then
    strSql = strSql & " , INDICADO_POR = '" & m_25_indicadoPor & "'                                         " & vbCrLf
    Else
    strSql = strSql & " , INDICADO_POR = NULL                                                               " & vbCrLf
    End If
    strSql = strSql & " , DATA_ALTERACAO = GetDate()                                                        " & vbCrLf
    strSql = strSql & " , USUARIO_ALTERACAO = '" & m_29_usuarioAlteracao & "'                               " & vbCrLf
    strSql = strSql & " , CT_LOCK = " & m_30_ctLock & "                                                     " & vbCrLf
    strSql = strSql & " WHERE ID_RESPONSAVEL    = " & m_01_idResponsavel & "                                " & vbCrLf
    strSql = strSql & "   AND ID_EMPRESACLIENTE = " & m_02_idEmpresaCliente & "                             " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudUpdate = m_01_idResponsavel
    
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
'Recebe uma conexão como parâmetro (com a transação aberta)
Public Function crudDelete(ByVal rsEndResFin As ADODB.Recordset, ByRef cnConexao As ADODB.Connection) As Long

Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Call moveEnderecoParaObjeto(rsEndResFin, "responsavelFinanceiro")
  
    'Deleta o endereco do responsavel financeiro
    With oEndereco
        .m_03_idEndereco = .crudDelete(cnConexao)
    End With
    
    strSql = ""
    strSql = strSql & "DELETE                                                   " & vbCrLf
    strSql = strSql & "  FROM responsavelFinanceiro                             " & vbCrLf
    strSql = strSql & " WHERE ID_RESPONSAVEL    = " & m_01_idResponsavel & "    " & vbCrLf
    strSql = strSql & "   AND ID_EMPRESACLIENTE = " & m_02_idEmpresaCliente & " " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = m_01_idResponsavel

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consulta(ByVal idEmpresaCliente As Long, _
                         ByVal idResponsavelFinanceiro As Long, _
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
    ssql = ssql & " SELECT * FROM responsavelFinanceiro WITH(NOLOCK)" & vbCrLf
    ssql = ssql & " WHERE ID_EMPRESACLIENTE = " & idEmpresaCliente & vbCrLf
    ssql = ssql & "   AND ID_RESPONSAVEL    = " & idResponsavelFinanceiro & vbCrLf
    
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function getResponsavelFinanceiro(ByVal lngIdEmpresaCliente As Long, _
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
    ssql = ssql & " SELECT * FROM responsavelFinanceiro WITH(NOLOCK)" & vbCrLf
    ssql = ssql & " WHERE ID_EMPRESACLIENTE = " & lngIdEmpresaCliente & vbCrLf
    
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set getResponsavelFinanceiro = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Sub inicializaResFin()
  Set rsResFin = New ADODB.Recordset
  rsResFin.Fields.Append "ID_RESPONSAVEL", adBigInt
  rsResFin.Fields.Append "ID_EMPRESACLIENTE", adBigInt
  rsResFin.Fields.Append "SITUACAO", adVarChar, 1
  rsResFin.Fields.Append "NOME", adVarChar, 50
  rsResFin.Fields.Append "CPF", adVarChar, 14
  rsResFin.Fields.Append "RG", adVarChar, 20
  rsResFin.Fields.Append "ORGAO_EMISSOR", adVarChar, 10
  rsResFin.Fields.Append "DATA_EXPEDICAO", adVarChar, 10
  rsResFin.Fields.Append "SEXO", adVarChar, 1
  rsResFin.Fields.Append "DATA_NASCIMENTO", adVarChar, 10
  rsResFin.Fields.Append "ESTADO_CIVIL", adVarChar, 1
  rsResFin.Fields.Append "NATURALIDADE", adVarChar, 20
  rsResFin.Fields.Append "NACIONALIDADE", adVarChar, 20
  rsResFin.Fields.Append "NOME_MAE", adVarChar, 50
  rsResFin.Fields.Append "ID_ENDERECO", adBigInt
  rsResFin.Fields.Append "TIPO_IMOVEL", adVarChar, 1
  rsResFin.Fields.Append "RESIDE_DESDE", adVarChar, 10
  rsResFin.Fields.Append "TELEFONE1", adVarChar, 13
  rsResFin.Fields.Append "TELEFONE2", adVarChar, 13
  rsResFin.Fields.Append "TELEFONE3", adVarChar, 13
  rsResFin.Fields.Append "CONTATO_REFERENCIA1", adVarChar, 50
  rsResFin.Fields.Append "TELEFONE_REFERENCIA1", adVarChar, 13
  rsResFin.Fields.Append "CONTATO_REFERENCIA2", adVarChar, 50
  rsResFin.Fields.Append "TELEFONE_REFERENCIA2", adVarChar, 13
  rsResFin.Fields.Append "INDICADO_POR", adVarChar, 50
  rsResFin.Fields.Append "DATA_INCLUSAO", adVarChar, 10
  rsResFin.Fields.Append "USUARIO_INCLUSAO", adVarChar, 20
  rsResFin.Fields.Append "DATA_ALTERACAO", adVarChar, 10
  rsResFin.Fields.Append "USUARIO_ALTERACAO", adVarChar, 20
  rsResFin.Fields.Append "CT_LOCK", adInteger
End Sub

Private Sub Class_Initialize()
    'inicializaResFin
End Sub
Public Function getNovoIdResponsavel() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT ISNULL(IDENT_CURRENT('responsavelFinanceiro') + IDENT_INCR('responsavelFinanceiro'),0)" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    getNovoIdResponsavel = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  getNovoIdResponsavel = 0
End Function
Public Function recuperarResponsavelFicanceiro(ByVal idEmpresaCliente) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM responsavelFinanceiro WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_EMPRESACLIENTE = " & idEmpresaCliente & vbCrLf
    
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperarResponsavelFicanceiro = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function


