VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinanciamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_FINANCIAMENTO As Long
Public m_02_ID_EMPRESACLIENTE As Long
Public m_03_ID_LINHACREDITO As Long
Public m_04_ID_ROTA As Long
Public m_05_VALOR_SACADO As Double
Public m_06_TAXA As Double
Public m_07_QTD_PARCELAS As Integer
Public m_08_VALOR_PARCELA As Double
Public m_09_DATA_PRIMEIRA_PARCELA As Date
Public m_10_DATA_ULTIMA_PARCELA As Date
Public m_11_SALDO_DEVEDOR As Double
Public m_12_DATA_INCLUSAO As String
Public m_13_USUARIO_INCLUSAO As String
Public m_14_DATA_ALTERACAO As String
Public m_15_USUARIO_ALTERACAO As String
Public m_16_CT_LOCK As Long
Public m_17_ID_CAIXA As Long
Public m_18_TAXA_JUROS_MORA As Double

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String

Private oParcelas As New clsFinanciamentoParcela
Private oLancamento As New clsCaixaLancamento
Private ocaixa As New clsCaixa
Private oEmpresaCliente As New clsEMPRESACLIENTE

Public Enum TipoLancamentosFuturos
   todos = 1
   dentro_prazo = 2
   fora_prazo = 3
End Enum
Public Enum ClassificacaoLancamentosFuturos
   clientes = 1
   vencimento = 2
End Enum

Public Function crudInsert(ByVal rsParcelas As ADODB.Recordset) As Long
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
    strSql = strSql & "SET DATEFORMAT DMY" & vbCrLf
    strSql = strSql & "INSERT INTO [credito_facil].[dbo].[financiamento]" & vbCrLf
    strSql = strSql & "           ([ID_EMPRESACLIENTE]" & vbCrLf
    strSql = strSql & "           ,[ID_LINHACREDITO]" & vbCrLf
    strSql = strSql & "           ,[ID_ROTA]" & vbCrLf
    strSql = strSql & "           ,[VALOR_SACADO]" & vbCrLf
    strSql = strSql & "           ,[TAXA]" & vbCrLf
    strSql = strSql & "           ,[QTD_PARCELAS]" & vbCrLf
    strSql = strSql & "           ,[VALOR_PARCELA]" & vbCrLf
    strSql = strSql & "           ,[DATA_PRIMEIRA_PARCELA]" & vbCrLf
    strSql = strSql & "           ,[DATA_ULTIMA_PARCELA]" & vbCrLf
    strSql = strSql & "           ,[SALDO_DEVEDOR]" & vbCrLf
    strSql = strSql & "           ,[DATA_INCLUSAO]" & vbCrLf
    strSql = strSql & "           ,[USUARIO_INCLUSAO]" & vbCrLf
    strSql = strSql & "           ,[DATA_ALTERACAO]" & vbCrLf
    strSql = strSql & "           ,[USUARIO_ALTERACAO]" & vbCrLf
    strSql = strSql & "           ,[CT_LOCK]" & vbCrLf
    strSql = strSql & "           ,[ID_CAIXA]" & vbCrLf
    strSql = strSql & "           ,[TAXA_JUROS_MORA])" & vbCrLf
    strSql = strSql & "     Values" & vbCrLf
    strSql = strSql & "           (" & m_02_ID_EMPRESACLIENTE & vbCrLf
    strSql = strSql & "           ," & m_03_ID_LINHACREDITO & vbCrLf
    strSql = strSql & "           ," & m_04_ID_ROTA & vbCrLf
    strSql = strSql & "           ," & Replace(m_05_VALOR_SACADO, ",", ".") & vbCrLf
    strSql = strSql & "           ," & Replace(m_06_TAXA, ",", ".") & vbCrLf
    strSql = strSql & "           ," & m_07_QTD_PARCELAS & vbCrLf
    strSql = strSql & "           ," & Replace(m_08_VALOR_PARCELA, ",", ".") & vbCrLf
    strSql = strSql & "           ,'" & m_09_DATA_PRIMEIRA_PARCELA & "'" & vbCrLf
    strSql = strSql & "           ,'" & m_10_DATA_ULTIMA_PARCELA & "'" & vbCrLf
    strSql = strSql & "           ," & Replace(m_11_SALDO_DEVEDOR, ",", ".") & vbCrLf
    strSql = strSql & "           ,getdate()" & vbCrLf
    strSql = strSql & "           ,'" & m_13_USUARIO_INCLUSAO & "'" & vbCrLf
    strSql = strSql & "           ,getdate()" & vbCrLf
    strSql = strSql & "           ,'" & m_15_USUARIO_ALTERACAO & "'" & vbCrLf
    strSql = strSql & "           ," & m_16_CT_LOCK & vbCrLf
    strSql = strSql & "           ," & m_17_ID_CAIXA & vbCrLf
    strSql = strSql & "           ," & Replace(m_18_TAXA_JUROS_MORA, ",", ".") & ")" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    strSql = "SELECT @@IDENTITY"
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        
    m_01_ID_FINANCIAMENTO = mRs(0)
    crudInsert = m_01_ID_FINANCIAMENTO
    'Insere as parcelas do financimanento
    crudInsert = oParcelas.crudInsert(rsParcelas, cnConexao, m_01_ID_FINANCIAMENTO)
    'Registra o lançamento no caixa
    oLancamento.mTIMEOUT = m_timeOut: oLancamento.mSTRING_CONEXAO = m_stringConexao
    oEmpresaCliente.m_timeOut = m_timeOut: oEmpresaCliente.m_stringConexao = m_stringConexao
    ocaixa.mTIMEOUT = m_timeOut
    ocaixa.mSTRING_CONEXAO = m_stringConexao
    oLancamento.m_02_ID_CAIXA = m_17_ID_CAIXA
    oLancamento.m_03_VALOR = m_05_VALOR_SACADO * -1
    oLancamento.m_04_TIPO = "E"
    oLancamento.m_05_SALDO_ANTERIOR = ocaixa.getSaldo(m_17_ID_CAIXA)
    oLancamento.m_06_SALDO_ATUAL = oLancamento.m_05_SALDO_ANTERIOR + oLancamento.m_03_VALOR
    oLancamento.m_08_OBS = "Saque - Liberação de Financiamento / Empresa (" & oEmpresaCliente.getNomeEmpresaCliente(m_02_ID_EMPRESACLIENTE) & ")"
    oLancamento.m_09_USUARIO_INCLUSAO = m_13_USUARIO_INCLUSAO
    crudInsert = oLancamento.adicionar(cnConexao)
    
    cnConexao.CommitTrans
    Set cnConexao = Nothing

Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function GetNovoIDFinanciamento() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT ISNULL(IDENT_CURRENT('financiamento') + IDENT_INCR('financiamento'),0)" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDFinanciamento = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDFinanciamento = 0
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
    strSql = strSql & "SET DATEFORMAT DMY" & vbCrLf
    strSql = strSql & "Update [credito_facil].[dbo].[financiamento]" & vbCrLf
    strSql = strSql & "   SET [ID_ROTA] = " & m_04_ID_ROTA & vbCrLf
    'strSql = strSql & "      ,[SALDO_DEVEDOR] = " & m_11_SALDO_DEVEDOR & vbCrLf
    strSql = strSql & "      ,[DATA_ALTERACAO] = getdate()" & vbCrLf
    strSql = strSql & "      ,[USUARIO_ALTERACAO] = '" & m_15_USUARIO_ALTERACAO & "'" & vbCrLf
    strSql = strSql & "      ,[CT_LOCK] = CT_LOCK + 1" & vbCrLf
    strSql = strSql & "      ,[TAXA_JUROS_MORA] = " & Replace(m_18_TAXA_JUROS_MORA, ",", ".") & vbCrLf
    strSql = strSql & " WHERE ID_FINANCIAMENTO = " & m_01_ID_FINANCIAMENTO
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudUpdate = m_01_ID_FINANCIAMENTO
    cnConexao.CommitTrans
    Set cnConexao = Nothing
    
Exit Function
trataerro:
  cnConexao.RollbackTrans
  Set cnConexao = Nothing
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudUpdate = 0
End Function
Public Function consultaPorIdFinanciamento(ByVal idFinanciamento As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM financiamento WITH (NOLOCK)" & vbCrLf
    ssql = ssql & " WHERE id_financiamento = " & idFinanciamento & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consultaPorIdFinanciamento = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function consultaDetalhada(ByVal idEmpresaCliente As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "set dateformat dmy" & vbCrLf
    ssql = ssql & " SELECT convert(varchar(10),DATA_PRIMEIRA_PARCELA,103) as DATA_PRIMEIRA_PARCELA , convert(varchar(10),DATA_ULTIMA_PARCELA,103) as DATA_ULTIMA_PARCELA,convert(varchar(10),DATA_VENCIMENTO,103) AS DATA_VENC, *" & vbCrLf
    'ssql = ssql & "  ID_FINANCIAMENTO" & vbCrLf
    'ssql = ssql & " ,ID_EMPRESACLIENTE" & vbCrLf
    'ssql = ssql & " ,ID_LINHACREDITO" & vbCrLf
    'ssql = ssql & " ,ID_ROTA" & vbCrLf
    'ssql = ssql & " ,VALOR_SACADO" & vbCrLf
    'ssql = ssql & " ,TAXA" & vbCrLf
    'ssql = ssql & " ,QTD_PARCELAS" & vbCrLf
    'ssql = ssql & " ,VALOR_PARCELA" & vbCrLf
    'ssql = ssql & " ,convert(varchar(10),DATA_PRIMEIRA_PARCELA,103) as DATA_PRIMEIRA_PARCELA" & vbCrLf
    'ssql = ssql & " ,convert(varchar(10),DATA_ULTIMA_PARCELA,103) as DATA_ULTIMA_PARCELA" & vbCrLf
    'ssql = ssql & " ,SALDO_DEVEDOR" & vbCrLf
    'ssql = ssql & " ,DATA_INCLUSAO" & vbCrLf
    'ssql = ssql & " ,USUARIO_INCLUSAO" & vbCrLf
    'ssql = ssql & " ,DATA_ALTERACAO" & vbCrLf
    'ssql = ssql & " ,USUARIO_ALTERACAO" & vbCrLf
    'ssql = ssql & " ,CT_LOCK" & vbCrLf
    'ssql = ssql & " ,ID_CAIXA" & vbCrLf
    'ssql = ssql & " ,TAXA_JUROS_MORA" & vbCrLf
    ssql = ssql & " FROM financiamento WITH (NOLOCK), financiamentoParcela with(nolock)" & vbCrLf
    ssql = ssql & "  Where financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO And id_EmpresaCliente = " & idEmpresaCliente & vbCrLf
    ssql = ssql & "  AND financiamento.saldo_devedor > 0" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consultaDetalhada = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function ResumoAuditoriaFinanciamentosAtuais() As ADODB.Recordset
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = strSql & " SELECT" & vbCrLf
    strSql = strSql & "  funcionario.NOME AS ROTA" & vbCrLf
    strSql = strSql & " ,CASE empresaCliente.TIPO" & vbCrLf
    strSql = strSql & "     WHEN 'F' THEN empresaCliente.NOME_PESSOA_FISICA" & vbCrLf
    strSql = strSql & "     WHEN 'J' THEN empresaCliente.RAZAO_SOCIAL" & vbCrLf
    strSql = strSql & "  END As CLIENTE" & vbCrLf
    strSql = strSql & " ,financiamento.SALDO_DEVEDOR" & vbCrLf
    strSql = strSql & " FROM rota INNER JOIN funcionario ON rota.ID_FUNCIONARIO = funcionario.ID_FUNCIONARIO" & vbCrLf
    strSql = strSql & " INNER JOIN financiamento ON financiamento.ID_ROTA = rota.ID_ROTA" & vbCrLf
    strSql = strSql & " INNER JOIN empresaCliente ON financiamento.ID_EMPRESACLIENTE = empresaCliente.ID_EMPRESACLIENTE" & vbCrLf
    strSql = strSql & " Where financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    strSql = strSql & " order by ROTA, CLIENTE " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set ResumoAuditoriaFinanciamentosAtuais = mRs
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description

End Function
Public Function crudDelete() As Long

Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String
Dim idEmpresaCliente As Long

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    cnConexao.BeginTrans
    
    'Excluir as parcelas
    crudDelete = oParcelas.crudDelete(m_01_ID_FINANCIAMENTO, cnConexao)
        
    'Recupera o IdEmpresaCliente desde financiamento
    strSql = ""
    strSql = strSql & "SELECT id_empresaCliente FROM financiamento" & vbCrLf
    strSql = strSql & " WHERE id_financiamento = " & m_01_ID_FINANCIAMENTO & vbCrLf
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    idEmpresaCliente = mRs(0)
    
    strSql = ""
    strSql = strSql & "DELETE                                       " & vbCrLf
    strSql = strSql & "  FROM financiamento                         " & vbCrLf
    strSql = strSql & " WHERE id_financiamento = " & m_01_ID_FINANCIAMENTO & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = m_01_ID_FINANCIAMENTO
    
    'Registra o lançamento no caixa
    oLancamento.mTIMEOUT = m_timeOut:    oLancamento.mSTRING_CONEXAO = m_stringConexao
    ocaixa.mTIMEOUT = m_timeOut:    ocaixa.mSTRING_CONEXAO = m_stringConexao
    oEmpresaCliente.m_timeOut = m_timeOut: oEmpresaCliente.m_stringConexao = m_stringConexao
    oLancamento.m_02_ID_CAIXA = m_17_ID_CAIXA
    oLancamento.m_03_VALOR = m_05_VALOR_SACADO
    oLancamento.m_04_TIPO = "E"
    oLancamento.m_05_SALDO_ANTERIOR = ocaixa.getSaldo(m_17_ID_CAIXA)
    oLancamento.m_06_SALDO_ATUAL = oLancamento.m_05_SALDO_ANTERIOR + oLancamento.m_03_VALOR
    oLancamento.m_08_OBS = "Estorno de Financiamento / Empresa (" & oEmpresaCliente.getNomeEmpresaCliente(idEmpresaCliente) & ")"
    oLancamento.m_09_USUARIO_INCLUSAO = m_13_USUARIO_INCLUSAO
    crudDelete = oLancamento.adicionar(cnConexao)
    
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
    ssql = ssql & " SELECT * FROM financiamento WITH (NOLOCK)" & vbCrLf
    ssql = ssql & " WHERE id_EmpresaCliente = " & idEmpresaCliente & vbCrLf
    ssql = ssql & " AND saldo_devedor > 0"

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarFinanciamentos() As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = "SELECT * FROM financiamento WITH (NOLOCK) "

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarFinanciamentos = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function EmpresaPossuiFinanciamentoPendente(ByVal idEmpresaCliente As Long) As Boolean
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT * FROM financiamento WHERE SALDO_DEVEDOR > 0 AND ID_EMPRESACLIENTE = " & idEmpresaCliente
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    EmpresaPossuiFinanciamentoPendente = IIf(mRs.EOF, False, True)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RelatorioLancamentosFuturos(ByVal tipoLancamentos As TipoLancamentosFuturos, _
                                            ByVal classificacao As ClassificacaoLancamentosFuturos) As ADODB.Recordset
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = strSql & " " & vbCrLf
    strSql = strSql & "  set dateformat dmy" & vbCrLf
    strSql = strSql & "  select" & vbCrLf
    strSql = strSql & "  case when empresaCliente.NOME_FANTASIA  is null or empresaCliente.NOME_FANTASIA = '' then" & vbCrLf
    strSql = strSql & "    empresaCliente.NOME_PESSOA_FISICA else" & vbCrLf
    strSql = strSql & "    empresaCliente.RAZAO_SOCIAL" & vbCrLf
    strSql = strSql & "  end as CLIENTES," & vbCrLf
    strSql = strSql & "  financiamento.VALOR_SACADO," & vbCrLf
    strSql = strSql & "  financiamento.TAXA," & vbCrLf
    strSql = strSql & "  financiamento.VALOR_SACADO * (financiamento.TAXA/100) as DESAGIO," & vbCrLf
    strSql = strSql & "  financiamento.DATA_ULTIMA_PARCELA as VENCIMENTO," & vbCrLf
    strSql = strSql & "  financiamento.VALOR_PARCELA as DIARIA," & vbCrLf
    strSql = strSql & "  financiamento.QTD_PARCELAS * financiamento.VALOR_PARCELA as TOTAL" & vbCrLf
    strSql = strSql & "  From financiamento, empresaCliente" & vbCrLf
    strSql = strSql & "  Where financiamento.ID_EMPRESACLIENTE = empresaCliente.ID_EMPRESACLIENTE" & vbCrLf
    strSql = strSql & "  and financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    Select Case tipoLancamentos
      Case todos

      Case dentro_prazo
         strSql = strSql & "  and financiamento.DATA_ULTIMA_PARCELA >= getdate()" & vbCrLf
      Case fora_prazo
         strSql = strSql & "  and financiamento.DATA_ULTIMA_PARCELA < getdate()" & vbCrLf
    End Select
    Select Case classificacao
      Case clientes
         strSql = strSql & "  order by CLIENTES" & vbCrLf
      Case vencimento
         strSql = strSql & "  order by VENCIMENTO" & vbCrLf
    End Select
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set RelatorioLancamentosFuturos = mRs
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description

End Function

