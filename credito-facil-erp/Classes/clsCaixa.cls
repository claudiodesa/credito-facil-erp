VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_CAIXA As Long
Public m_02_DATA_ABERTURA As Date
Public m_03_DATA_FECHAMENTO As Date
Public m_04_SALDO_ABERTURA As Double
Public m_05_SALDO_FECHAMENTO As Double
Public m_06_USUARIO_ABERTURA As String
Public m_07_USUARIO_FECHAMENTO As String
'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public mTIMEOUT As String
Public mSTRING_CONEXAO As String

Public Function abrirCaixa() As Long
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
    strSql = strSql & "INSERT INTO CAIXA                                        " & vbCrLf
    strSql = strSql & "            (DATA_ABERTURA,                              " & vbCrLf
    strSql = strSql & "             SALDO_ABERTURA,                             " & vbCrLf
    strSql = strSql & "             USUARIO_ABERTURA)                           " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        (Getdate()                                       " & vbCrLf
    strSql = strSql & "        ," & Replace(m_04_SALDO_ABERTURA, ",", ".") & "                     " & vbCrLf
    strSql = strSql & "        ,'" & m_06_USUARIO_ABERTURA & "')                " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    strSql = "SELECT @@IDENTITY"
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    abrirCaixa = mRs(0)
    Set cnConexao = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  abrirCaixa = 0
End Function
Public Function GetNovoIDCaixa() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('caixa') + IDENT_INCR('caixa')" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDCaixa = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDCaixa = 0
End Function
Public Function fecharCaixa() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open
    
    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY" & vbCrLf
    strSql = strSql & "Update [credito_facil].[dbo].[caixa]" & vbCrLf
    strSql = strSql & "   SET [DATA_FECHAMENTO] = Getdate()" & vbCrLf
    strSql = strSql & "      ,[SALDO_FECHAMENTO] = " & Replace(m_05_SALDO_FECHAMENTO, ",", ".") & vbCrLf
    strSql = strSql & "      ,[USUARIO_FECHAMENTO] ='" & m_07_USUARIO_FECHAMENTO & "'" & vbCrLf
    strSql = strSql & " WHERE ID_CAIXA = " & m_01_ID_CAIXA
    
   Set mRs = CtxCreateRecordSet(cnConexao, strSql)

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    fecharCaixa = m_01_ID_CAIXA
    
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  fecharCaixa = 0
End Function

Public Function consulta(ByVal idCaixa As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM CAIXA WITH (NOLOCK) " & vbCrLf
    ssql = ssql & " WHERE ID_CAIXA = " & idCaixa
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function listarCaixasFechados() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " select" & vbCrLf
    ssql = ssql & " caixa.ID_CAIXA," & vbCrLf
    ssql = ssql & " convert(varchar(10),DATA_ABERTURA,103) + ' - ' + 'Abriu com: ' + convert(varchar(10),SALDO_ABERTURA) + ' | Fechou com: ' + convert(varchar(10),SALDO_FECHAMENTO)" & vbCrLf
    ssql = ssql & " From caixa" & vbCrLf
    ssql = ssql & " Where caixa.DATA_FECHAMENTO Is Not Null" & vbCrLf
    ssql = ssql & " order by ID_CAIXA desc" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set listarCaixasFechados = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function consultaResumoCaixaAtual() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " set dateformat dmy" & vbCrLf
    ssql = ssql & " SELECT" & vbCrLf
    ssql = ssql & " caixa.ID_CAIXA," & vbCrLf
    ssql = ssql & " caixa.DATA_ABERTURA," & vbCrLf
    ssql = ssql & " caixa.SALDO_ABERTURA," & vbCrLf
    ssql = ssql & " caixa.USUARIO_ABERTURA," & vbCrLf
    ssql = ssql & " caixa.DATA_FECHAMENTO," & vbCrLf
    ssql = ssql & " caixa.SALDO_FECHAMENTO," & vbCrLf
    ssql = ssql & " caixa.USUARIO_FECHAMENTO," & vbCrLf
    ssql = ssql & " caixa.SALDO_FECHAMENTO - caixa.SALDO_ABERTURA AS APURADO," & vbCrLf
    ssql = ssql & " caixaLancamento.TIPO," & vbCrLf
    ssql = ssql & " caixaLancamento.VALOR," & vbCrLf
    ssql = ssql & " caixaLancamento.OBSERVACAO," & vbCrLf
    ssql = ssql & " caixaLancamento.USUARIO_INCLUSAO" & vbCrLf
    ssql = ssql & " From" & vbCrLf
    ssql = ssql & " caixa WITH(NOLOCK)," & vbCrLf
    ssql = ssql & " caixaLancamento WITH(NOLOCK)" & vbCrLf
    ssql = ssql & " Where" & vbCrLf
    ssql = ssql & "     caixa.ID_CAIXA = caixaLancamento.ID_CAIXA" & vbCrLf
    ssql = ssql & " and caixa.DATA_FECHAMENTO is null" & vbCrLf
    ssql = ssql & " Order By" & vbCrLf
    ssql = ssql & "     caixa.ID_CAIXA," & vbCrLf
    ssql = ssql & "     caixaLancamento.ID_LANCAMENTO" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consultaResumoCaixaAtual = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consultaResumoCaixaFechado(ByVal idCaixa As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = ssql & " set dateformat dmy" & vbCrLf
    ssql = ssql & " SELECT" & vbCrLf
    ssql = ssql & " caixa.ID_CAIXA," & vbCrLf
    ssql = ssql & " caixa.DATA_ABERTURA," & vbCrLf
    ssql = ssql & " caixa.SALDO_ABERTURA," & vbCrLf
    ssql = ssql & " caixa.USUARIO_ABERTURA," & vbCrLf
    ssql = ssql & " caixa.DATA_FECHAMENTO," & vbCrLf
    ssql = ssql & " caixa.SALDO_FECHAMENTO," & vbCrLf
    ssql = ssql & " caixa.USUARIO_FECHAMENTO," & vbCrLf
    ssql = ssql & " caixa.SALDO_FECHAMENTO - caixa.SALDO_ABERTURA AS APURADO," & vbCrLf
    ssql = ssql & " caixaLancamento.TIPO," & vbCrLf
    ssql = ssql & " caixaLancamento.VALOR," & vbCrLf
    ssql = ssql & " caixaLancamento.OBSERVACAO," & vbCrLf
    ssql = ssql & " caixaLancamento.USUARIO_INCLUSAO" & vbCrLf
    ssql = ssql & " From" & vbCrLf
    ssql = ssql & " caixa WITH(NOLOCK)," & vbCrLf
    ssql = ssql & " caixaLancamento WITH(NOLOCK)" & vbCrLf
    ssql = ssql & " Where" & vbCrLf
    ssql = ssql & "     caixa.ID_CAIXA = caixaLancamento.ID_CAIXA" & vbCrLf
    ssql = ssql & " and caixa.DATA_FECHAMENTO is not null" & vbCrLf
    ssql = ssql & " and caixa.ID_CAIXA = " & idCaixa & vbCrLf
    ssql = ssql & " Order By" & vbCrLf
    ssql = ssql & "     caixa.ID_CAIXA," & vbCrLf
    ssql = ssql & "     caixaLancamento.ID_LANCAMENTO" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consultaResumoCaixaFechado = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function IdUltimoCaixaAberto() As Long
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = "SELECT ID_CAIXA FROM caixa WITH(NOLOCK) WHERE DATA_FECHAMENTO IS NULL"
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    If mRs.EOF Then Exit Function
    IdUltimoCaixaAberto = mRs(0)
    Set mRs.ActiveConnection = Nothing
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function CaixaAberto() As Boolean
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = ""
    ssql = "SELECT count(ID_CAIXA) FROM caixa WITH(NOLOCK) WHERE DATA_FECHAMENTO IS NULL"
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    If mRs(0) > 0 Then
        CaixaAberto = True
    End If
    Set mRs.ActiveConnection = Nothing
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function IdUltimoCaixa() As Long
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset
Dim QtCaixas As Long

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
    ssql = "SELECT COUNT(*) FROM CAIXA WITH(NOLOCK)"
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    QtCaixas = mRs(0)
    
    If QtCaixas > 0 Then
        ssql = "SELECT ISNULL(MAX(ID_CAIXA),0) FROM caixa WITH(NOLOCK)"
        Set mRs = CtxCreateRecordSet(mCn, ssql)
        Set mRs.ActiveConnection = Nothing
        IdUltimoCaixa = mRs(0)
        Set mCn = Nothing
    End If

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function getSaldo(ByVal idCaixa As Long) As Double

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = mTIMEOUT
    mCn.ConnectionString = mSTRING_CONEXAO
    mCn.Open
    
'    ssql = ""
'    ssql = ssql & " select caixa.SALDO_ABERTURA + SUM(caixaLancamento.VALOR) as SALDO_ATUAL" & vbCrLf
'    ssql = ssql & " from caixa,caixaLancamento with(nolock)" & vbCrLf
'    ssql = ssql & " Where caixa.ID_CAIXA = caixaLancamento.ID_CAIXA" & vbCrLf
'    ssql = ssql & " and   caixa.ID_CAIXA = " & idCaixa & vbCrLf
'    ssql = ssql & " group by caixa.SALDO_ABERTURA" & vbCrLf
    ssql = ""
    ssql = ssql & " declare @soma as numeric(18,2)"
    ssql = ssql & " set @soma = (select isnull(SUM(coalesce(caixaLancamento.VALOR,0)),0) from caixaLancamento where ID_CAIXA = " & idCaixa & ")"
    ssql = ssql & " select distinct caixa.SALDO_ABERTURA + @soma as SALDO_ATUAL"
    ssql = ssql & "  from caixa,caixaLancamento with(nolock)"
    ssql = ssql & "  Where caixa.ID_CAIXA = " & idCaixa

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    If mRs.EOF Then Exit Function
    getSaldo = mRs(0)
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

