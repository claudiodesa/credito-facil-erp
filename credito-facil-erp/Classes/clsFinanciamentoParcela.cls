VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinanciamentoParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_PARCELAMENTO As Long
Public m_02_ID_FINANCIAMENTO As Long
Public m_03_NUM_PARCELA As Integer
Public m_04_DATA_VENCIMENTO As Date
Public m_05_VALOR_COBRADO As Double
Public m_06_DATA_PAGAMENTO As Date
Public m_07_VALOR_RECEBIDO As Double
Public m_08_DIAS_ATRASO As Integer
Public m_09_SALDO_DEVEDOR As Double
Public m_10_DATA_INCLUSAO As String
Public m_11_USUARIO_INCLUSAO As String
Public m_12_DATA_ALTERACAO As String
Public m_13_USUARIO_ALTERACAO As String
Public m_14_CT_LOCK As Long

'Array de parcelas
Public rsFinanciamentoParcela As ADODB.Recordset
Private ocaixa      As New clsCaixa
Private oLancamento As New clsCaixaLancamento
Private oFinancimento As New clsFinanciamento
Private oFeriado As New clsFERIADO
Private oEmpresaCliente As New clsEMPRESACLIENTE

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public m_timeOut As String
Public m_stringConexao As String
Public Sub inicializaParcela()
  Set rsFinanciamentoParcela = New ADODB.Recordset
  rsFinanciamentoParcela.Fields.Append "ID_PARCELAMENTO", adBigInt
  rsFinanciamentoParcela.Fields.Append "ID_FINANCIAMENTO", adBigInt
  rsFinanciamentoParcela.Fields.Append "NUM_PARCELA", adInteger
  rsFinanciamentoParcela.Fields.Append "DATA_VENCIMENTO", adVarChar, 10
  rsFinanciamentoParcela.Fields.Append "VALOR_COBRADO", adDouble
  rsFinanciamentoParcela.Fields.Append "DATA_PAGAMENTO", adVarChar, 10
  rsFinanciamentoParcela.Fields.Append "VALOR_RECEBIDO", adDouble
  rsFinanciamentoParcela.Fields.Append "DIAS_ATRASO", adInteger
  rsFinanciamentoParcela.Fields.Append "SALDO_DEVEDOR", adDouble
  rsFinanciamentoParcela.Fields.Append "DATA_INCLUSAO", adVarChar, 10
  rsFinanciamentoParcela.Fields.Append "USUARIO_INCLUSAO", adVarChar, 20
  rsFinanciamentoParcela.Fields.Append "DATA_ALTERACAO", adVarChar, 10
  rsFinanciamentoParcela.Fields.Append "USUARIO_ALTERACAO", adVarChar, 20
  rsFinanciamentoParcela.Fields.Append "CT_LOCK", adInteger
End Sub
Public Function pagarParcela(ByVal idCaixa As Long, ByVal idEmpresa As Long) As Long

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset
Dim rsFinanciamento As ADODB.Recordset
Dim SaldoDevedor As Double
Dim ValorComJurosMora As Double
Dim JurosMora As Long
Dim VencimentoNovaParcela As Date

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    mCn.BeginTrans
    
    'Verifica se o pagamento realizado e da ultim.parcela e o valor pago eh menos q o valor da parcela,
    'neste caso devera ser recebido o pagamento e processada uma nova parcela, com acrescimo parametrizado No
    'financiamento, sobre o saldo devedor.
    oFinancimento.m_stringConexao = m_stringConexao
    oFinancimento.m_timeOut = m_timeOut
    Set rsFinanciamento = oFinancimento.consultaPorIdFinanciamento(m_02_ID_FINANCIAMENTO)
    
    'Preparar para efetuar o pagamento de parcela
    ssql = ""
    ssql = ssql & " SET DATEFORMAT DMY" & vbCrLf
    ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
    ssql = ssql & " SET DATA_PAGAMENTO = '" & m_06_DATA_PAGAMENTO & "'" & vbCrLf
    ssql = ssql & "    ,VALOR_COBRADO =" & Replace(m_05_VALOR_COBRADO, ",", ".") & vbCrLf
    ssql = ssql & "    ,VALOR_RECEBIDO =" & Replace(m_07_VALOR_RECEBIDO, ",", ".") & vbCrLf
    ssql = ssql & "    ,DIAS_ATRASO = CASE " & vbCrLf
    ssql = ssql & "                     when DATEDIFF(dd,DATA_VENCIMENTO,'" & m_06_DATA_PAGAMENTO & "') < 0 then " & vbCrLf
    ssql = ssql & "                         0" & vbCrLf
    ssql = ssql & "                     else DATEDIFF(dd,DATA_VENCIMENTO,'" & m_06_DATA_PAGAMENTO & "')" & vbCrLf
    ssql = ssql & "                   END" & vbCrLf
    ssql = ssql & "    ,DATA_ALTERACAO = getdate()" & vbCrLf
    ssql = ssql & "    ,USUARIO_ALTERACAO = '" & m_13_USUARIO_ALTERACAO & "'" & vbCrLf
    ssql = ssql & " WHERE" & vbCrLf
    ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
    ssql = ssql & " AND NUM_PARCELA = " & m_03_NUM_PARCELA & vbCrLf
    
    'faz o pagamento
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    
    'Verificar se está pagando na ultima parcela um valor inferior à parcela
    If m_03_NUM_PARCELA >= rsFinanciamento("QTD_PARCELAS") And m_07_VALOR_RECEBIDO < m_05_VALOR_COBRADO Then
        'Cria uma nova parcela, com vencimento no proximo dia util, no valor do saldo devedor, acrescido do juros de mora
        'acordada no financimento (padrao 0.75%)
        inicializaParcela
        VencimentoNovaParcela = Format(Now() + 1, "dd/mm/yyyy")
        oFeriado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
        oFeriado.mTIMEOUT = gstrTimeOutGeral
        rsFinanciamentoParcela.Open
        With rsFinanciamentoParcela
            .AddNew
            .Fields("ID_FINANCIAMENTO") = m_02_ID_FINANCIAMENTO
            .Fields("NUM_PARCELA") = m_03_NUM_PARCELA + 1
            Do While (FinalDeSemana(VencimentoNovaParcela) Or oFeriado.DataFeriado(VencimentoNovaParcela))
               VencimentoNovaParcela = VencimentoNovaParcela + 1
            Loop
            .Fields("DATA_VENCIMENTO") = Format(VencimentoNovaParcela, "dd/mm/yyyy")
            ValorComJurosMora = ((m_05_VALOR_COBRADO - m_07_VALOR_RECEBIDO) * rsFinanciamento("TAXA_JUROS_MORA") / 100) + (m_05_VALOR_COBRADO - m_07_VALOR_RECEBIDO)
            ValorComJurosMora = Format(ValorComJurosMora, "0.00")
            ValorComJurosMora = CCur(ValorComJurosMora)
            .Fields("VALOR_COBRADO") = ValorComJurosMora
            .Fields("SALDO_DEVEDOR") = ValorComJurosMora
            .Fields("USUARIO_INCLUSAO") = LogInUserID
            .Fields("USUARIO_ALTERACAO") = LogInUserID
            .Update
        End With
        
        JurosMora = crudInsert(rsFinanciamentoParcela, mCn, m_02_ID_FINANCIAMENTO)
      
    ElseIf m_07_VALOR_RECEBIDO = rsFinanciamento("SALDO_DEVEDOR") Then 'Se quitou antecipadamente
        ssql = ""
        ssql = ssql & " SET DATEFORMAT DMY" & vbCrLf
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        ssql = ssql & " SET DATA_PAGAMENTO = '" & m_06_DATA_PAGAMENTO & "'" & vbCrLf
        ssql = ssql & "    ,VALOR_RECEBIDO =" & 0 & vbCrLf
        ssql = ssql & "    ,DIAS_ATRASO = 0" & vbCrLf
        ssql = ssql & "    ,DATA_ALTERACAO = getdate()" & vbCrLf
        ssql = ssql & "    ,USUARIO_ALTERACAO = '" & m_13_USUARIO_ALTERACAO & "'" & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        ssql = ssql & " AND NUM_PARCELA > " & m_03_NUM_PARCELA & vbCrLf
        'faz a quitação das parcelas que faltam
        Set mRs = CtxCreateRecordSet(mCn, ssql)

    ElseIf m_07_VALOR_RECEBIDO < m_05_VALOR_COBRADO Then 'Senao, verifica se pagou na parcela um valor inferior ao devido
        'Se pagou com valor menor que a parcela, lançar a diferença na próxima parcela
        ssql = "" & vbCrLf
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        ssql = ssql & " SET VALOR_COBRADO = VALOR_COBRADO + " & Replace(m_05_VALOR_COBRADO - m_07_VALOR_RECEBIDO, ",", ".") & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        ssql = ssql & " AND NUM_PARCELA = " & m_03_NUM_PARCELA + 1
        
        'atualiza a parcela seguinte caso o pagamento atual não tenha sido suficiente
        Set mRs = CtxCreateRecordSet(mCn, ssql)
     ElseIf m_07_VALOR_RECEBIDO > m_05_VALOR_COBRADO And m_03_NUM_PARCELA < rsFinanciamento("QTD_PARCELAS") Then  'Senao, verifica se pagou na parcela um valor superior ao devido Then
        'Se pagou com valor superior à parcela, abater a diferença na próxima parcela
        ssql = "" & vbCrLf
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        ssql = ssql & " SET VALOR_COBRADO = VALOR_COBRADO - " & Replace(m_07_VALOR_RECEBIDO - m_05_VALOR_COBRADO, ",", ".") & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        ssql = ssql & " AND NUM_PARCELA = " & m_03_NUM_PARCELA + 1
        
        'atualiza a parcela seguinte caso o pagamento atual não tenha sido suficiente
        Set mRs = CtxCreateRecordSet(mCn, ssql)
     
     ElseIf m_07_VALOR_RECEBIDO < 0 Then 'Se o valor recebido for negativo, significa saldo de sobra do cliente, que deverá ser abatido da proxima parcela
        'Se o cliente possui saldo de sobra por pagamento além do valor da parcela, efetuado anteriormente, abater a diferença da próxima parcela
        ssql = "" & vbCrLf
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        ssql = ssql & " SET VALOR_COBRADO = VALOR_COBRADO + " & Replace(m_07_VALOR_RECEBIDO, ",", ".") & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        ssql = ssql & " AND NUM_PARCELA = " & m_03_NUM_PARCELA + 1
        
        'atualiza a parcela seguinte caso o pagamento atual não tenha sido suficiente
        Set mRs = CtxCreateRecordSet(mCn, ssql)
     
     End If
    
    If m_07_VALOR_RECEBIDO > 0 Then
        'Prepara para atualizar o saldo devedor das parcelas
        ssql = ""
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        If ValorComJurosMora = 0 Then
            ssql = ssql & " SET SALDO_DEVEDOR = SALDO_DEVEDOR - " & Replace(m_07_VALOR_RECEBIDO, ",", ".") & vbCrLf
        Else
            ssql = ssql & " SET SALDO_DEVEDOR = " & Replace(ValorComJurosMora, ",", ".")
        End If
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        ssql = ssql & " AND NUM_PARCELA >= " & m_03_NUM_PARCELA
        'atualiza o saldo devedor nas parcelas
        Set mRs = CtxCreateRecordSet(mCn, ssql)
        
        'Prepara para atualizar o saldo devedor do financiamento
        ssql = ""
        ssql = ssql & " UPDATE financiamento" & vbCrLf
        If ValorComJurosMora = 0 Then
            ssql = ssql & " SET SALDO_DEVEDOR = ( SALDO_DEVEDOR - " & Replace(m_07_VALOR_RECEBIDO, ",", ".") & ")" & vbCrLf
        Else
            ssql = ssql & " SET SALDO_DEVEDOR = " & Replace(ValorComJurosMora, ",", ".")
        End If
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        
        'atualiza o saldo devedor no financiamento
        Set mRs = CtxCreateRecordSet(mCn, ssql)
    End If
    
    If m_07_VALOR_RECEBIDO > 0 Then
        'faz o lançamento no caixa, do recebimento da parcela
        ocaixa.mTIMEOUT = m_timeOut: ocaixa.mSTRING_CONEXAO = m_stringConexao
        oEmpresaCliente.m_timeOut = m_timeOut: oEmpresaCliente.m_stringConexao = m_stringConexao
        oLancamento.m_02_ID_CAIXA = idCaixa
        oLancamento.m_03_VALOR = m_07_VALOR_RECEBIDO
        oLancamento.m_04_TIPO = "R"
        oLancamento.m_05_SALDO_ANTERIOR = ocaixa.getSaldo(idCaixa)
        oLancamento.m_07_DATA = m_06_DATA_PAGAMENTO
        oLancamento.m_08_OBS = "Recebimento Parcela.: (" & Format(m_03_NUM_PARCELA, "00") & ") Empresa: (" & oEmpresaCliente.getNomeEmpresaCliente(idEmpresa) & ")"
        oLancamento.m_09_USUARIO_INCLUSAO = LogInUserID
        
        If oLancamento.adicionar(mCn) > 0 Then
        End If
    End If
     
     If m_07_VALOR_RECEBIDO = 0 Then 'Se valor recebido foi zero, acrescentar 1 dia de atraso
        ssql = "" & vbCrLf
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        ssql = ssql & " SET DIAS_ATRASO = DIAS_ATRASO + 1" & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & m_02_ID_FINANCIAMENTO & vbCrLf
        ssql = ssql & " AND NUM_PARCELA = " & m_03_NUM_PARCELA
        
        'atualiza 1 dia de atraso por pagamento zerado
        Set mRs = CtxCreateRecordSet(mCn, ssql)
     End If
    
    mCn.CommitTrans
    pagarParcela = m_02_ID_FINANCIAMENTO
        
     
Exit Function
trataerro:
  mCn.RollbackTrans
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  pagarParcela = 0
End Function
Public Function estornarParcela(ByVal idCaixa As Long, ByVal idFinanciamento As Long, ByVal idParcela As Long, ByVal Valor As Double, idEmpresaCliente As Long) As Long

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset
Dim rsFinanciamento As ADODB.Recordset
Dim SaldoDevedor As Double

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    mCn.BeginTrans
    
    'Preparar para efetuar o estorno de parcela
    ssql = ""
    ssql = ssql & " SET DATEFORMAT DMY" & vbCrLf
    ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
    ssql = ssql & " SET DATA_PAGAMENTO    = null" & vbCrLf
    ssql = ssql & "    ,VALOR_RECEBIDO    = null" & vbCrLf
    ssql = ssql & "    ,DIAS_ATRASO       = 0" & vbCrLf
    ssql = ssql & " WHERE" & vbCrLf
    ssql = ssql & "     ID_FINANCIAMENTO = " & idFinanciamento & vbCrLf
    ssql = ssql & " AND NUM_PARCELA = " & idParcela & vbCrLf
    
    'faz o estorno
    Set mRs = CtxCreateRecordSet(mCn, ssql)
    
    If Valor > 0 Then
        'Prepara para atualizar o saldo devedor das parcelas
        ssql = ""
        ssql = ssql & " UPDATE financiamentoParcela" & vbCrLf
        ssql = ssql & " SET SALDO_DEVEDOR = SALDO_DEVEDOR + " & Replace(Valor, ",", ".") & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & idFinanciamento & vbCrLf
        ssql = ssql & " AND NUM_PARCELA >= " & idParcela
        'atualiza o saldo devedor nas parcelas
        Set mRs = CtxCreateRecordSet(mCn, ssql)
        
        'Prepara para atualizar o saldo devedor do financiamento
        ssql = ""
        ssql = ssql & " UPDATE financiamento" & vbCrLf
        ssql = ssql & " SET SALDO_DEVEDOR = ( SALDO_DEVEDOR + " & Replace(Valor, ",", ".") & ")" & vbCrLf
        ssql = ssql & " WHERE" & vbCrLf
        ssql = ssql & "     ID_FINANCIAMENTO = " & idFinanciamento & vbCrLf
        
        'atualiza o saldo devedor no financiamento
        Set mRs = CtxCreateRecordSet(mCn, ssql)
    End If
        
    
    
    'faz o lançamento no caixa, do recebimento da parcela
    ocaixa.mTIMEOUT = m_timeOut: ocaixa.mSTRING_CONEXAO = m_stringConexao
    oEmpresaCliente.m_timeOut = m_timeOut: oEmpresaCliente.m_stringConexao = m_stringConexao
    oLancamento.m_02_ID_CAIXA = idCaixa
    oLancamento.m_03_VALOR = Valor * -1
    oLancamento.m_04_TIPO = "A"
    oLancamento.m_05_SALDO_ANTERIOR = ocaixa.getSaldo(idCaixa)
    oLancamento.m_07_DATA = Now()
    oLancamento.m_08_OBS = "Estorno Pag.  Parcela.: (" & Format(idParcela, "00") & ") Empresa: (" & oEmpresaCliente.getNomeEmpresaCliente(idEmpresaCliente) & ")"
    oLancamento.m_09_USUARIO_INCLUSAO = LogInUserID
        
    If oLancamento.adicionar(mCn) > 0 Then
    End If
    
    mCn.CommitTrans
    estornarParcela = idFinanciamento
        
     
Exit Function
trataerro:
  mCn.RollbackTrans
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  estornarParcela = 0
End Function

Public Function crudInsert(ByVal rsParcelas As ADODB.Recordset, ByRef cnConexao As ADODB.Connection, ByVal idFinanciamento As Long) As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String

    On Error GoTo trataerro
    
    rsParcelas.MoveFirst
    Do While Not rsParcelas.EOF
        strSql = ""
        strSql = strSql & "SET DATEFORMAT DMY" & vbCrLf
        strSql = strSql & "INSERT INTO [credito_facil].[dbo].[financiamentoParcela]" & vbCrLf
        strSql = strSql & "           ([ID_FINANCIAMENTO]" & vbCrLf
        strSql = strSql & "           ,[NUM_PARCELA]" & vbCrLf
        strSql = strSql & "           ,[DATA_VENCIMENTO]" & vbCrLf
        strSql = strSql & "           ,[VALOR_COBRADO]" & vbCrLf
        strSql = strSql & "           ,[DATA_PAGAMENTO]" & vbCrLf
        strSql = strSql & "           ,[VALOR_RECEBIDO]" & vbCrLf
        strSql = strSql & "           ,[DIAS_ATRASO]" & vbCrLf
        strSql = strSql & "           ,[SALDO_DEVEDOR]" & vbCrLf
        strSql = strSql & "           ,[DATA_INCLUSAO]" & vbCrLf
        strSql = strSql & "           ,[USUARIO_INCLUSAO]" & vbCrLf
        strSql = strSql & "           ,[DATA_ALTERACAO]" & vbCrLf
        strSql = strSql & "           ,[USUARIO_ALTERACAO]" & vbCrLf
        strSql = strSql & "           ,[CT_LOCK])" & vbCrLf
        strSql = strSql & "     Values" & vbCrLf
        strSql = strSql & "           (" & idFinanciamento & vbCrLf
        strSql = strSql & "           ," & rsParcelas("NUM_PARCELA") & vbCrLf
        strSql = strSql & "           ,'" & rsParcelas("DATA_VENCIMENTO") & "'" & vbCrLf
        strSql = strSql & "           ," & Replace(rsParcelas("VALOR_COBRADO"), ",", ".") & vbCrLf
        strSql = strSql & "           ,null" & vbCrLf
        strSql = strSql & "           ,null" & vbCrLf
        strSql = strSql & "           ,0" & vbCrLf
        strSql = strSql & "           ," & Replace(rsParcelas("SALDO_DEVEDOR"), ",", ".") & vbCrLf
        strSql = strSql & "           ,getdate()" & vbCrLf
        strSql = strSql & "           ,'" & rsParcelas("USUARIO_INCLUSAO") & "'" & vbCrLf
        strSql = strSql & "           ,getdate()" & vbCrLf
        strSql = strSql & "           ,'" & rsParcelas("USUARIO_ALTERACAO") & "'" & vbCrLf
        strSql = strSql & "           ,0)" & vbCrLf
        
        Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        Set mRs.ActiveConnection = Nothing
        strSql = "select @@identity"
        Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        crudInsert = mRs(0)
        rsParcelas.MoveNext
    Loop

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function crudDelete(ByVal idFinanciamento As Long, ByRef cnConexao As ADODB.Connection) As Long

Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    strSql = ""
    strSql = strSql & "DELETE                                       " & vbCrLf
    strSql = strSql & "  FROM financiamentoParcela                  " & vbCrLf
    strSql = strSql & " WHERE id_financiamento = " & idFinanciamento & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    crudDelete = idFinanciamento

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consulta(ByVal idFinanciamento As Long) As ADODB.Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM financiamentoParcela WITH (NOLOCK)" & vbCrLf
    ssql = ssql & " WHERE id_Financiamento = " & idFinanciamento

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function consultaCobranca(ByVal dtCobranca As String, ByVal idRota As Long) As Recordset

Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
   'ssql = ssql & " DECLARE @DATA AS VARCHAR(10)"
   'ssql = ssql & " SET @DATA = '16/08/2011'"
    ssql = ssql & " set dateformat dmy" & vbCrLf
    ssql = ssql & " select" & vbCrLf
    ssql = ssql & " right('0000' + convert(varchar(4),empresaCliente.ID_EMPRESACLIENTE),4) + ' - ' +" & vbCrLf
    ssql = ssql & " Case empresaCliente.TIPO" & vbCrLf
    ssql = ssql & "     when 'J' then" & vbCrLf
    ssql = ssql & "         empresaCliente.NOME_FANTASIA" & vbCrLf
    ssql = ssql & "     when 'F' then" & vbCrLf
    ssql = ssql & "         empresaCliente.NOME_PESSOA_FISICA" & vbCrLf
    ssql = ssql & " end +" & vbCrLf
    ssql = ssql & "  case when convert(varchar(2),datepart(DD,responsavelFinanceiro.DATA_NASCIMENTO)) + '/' +  convert(varchar(4),datepart(MM,responsavelFinanceiro.DATA_NASCIMENTO)) = convert(varchar(2),datepart(DD,'" & dtCobranca & "')) + '/' +  convert(varchar(4),datepart(MM,'" & dtCobranca & "')) then" & vbCrLf
    ssql = ssql & "     ' *P*'" & vbCrLf
    ssql = ssql & "  Else '' end" & vbCrLf
    ssql = ssql & " as CLIENTE," & vbCrLf
    ssql = ssql & "  responsavelFinanceiro.TELEFONE1," & vbCrLf
    ssql = ssql & " financiamento.SALDO_DEVEDOR," & vbCrLf
    'ssql = ssql & " isnull(convert(varchar(2),(" & vbCrLf
    'ssql = ssql & "          select count(financiamentoParcela.NUM_PARCELA)" & vbCrLf
    'ssql = ssql & "          From" & vbCrLf
    'ssql = ssql & "          empresaCliente        with(nolock)," & vbCrLf
    'ssql = ssql & "          financiamento         with(nolock)," & vbCrLf
    'ssql = ssql & "          financiamentoParcela  with(nolock)" & vbCrLf
    'ssql = ssql & "          Where" & vbCrLf
    'ssql = ssql & "                 empresaCliente.ID_EMPRESACLIENTE = financiamento.ID_EMPRESACLIENTE" & vbCrLf
    'ssql = ssql & "          and    financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    'ssql = ssql & "          and    financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    'ssql = ssql & "          and    financiamentoParcela.DATA_PAGAMENTO is not null),2),0)" & vbCrLf
    ssql = ssql & " convert(varchar(2),financiamentoParcela.NUM_PARCELA)" & vbCrLf
    ssql = ssql & "  + '/' +" & vbCrLf
    ssql = ssql & " convert(varchar(2),financiamento.QTD_PARCELAS) as PARCELAS_PAGAS," & vbCrLf
    ssql = ssql & " SUM(financiamentoParcela.VALOR_COBRADO) AS VALOR_DEVIDO," & vbCrLf
    ssql = ssql & " '' as RECEBIDO," & vbCrLf
    ssql = ssql & " '' as DATA," & vbCrLf
    ssql = ssql & " CONVERT(VARCHAR(10),financiamento.DATA_PRIMEIRA_PARCELA,103)," & vbCrLf
    ssql = ssql & " CONVERT(VARCHAR(10),financiamento.DATA_ULTIMA_PARCELA,103)" & vbCrLf
    ssql = ssql & " From" & vbCrLf
    ssql = ssql & " empresaCliente        with(nolock)," & vbCrLf
    ssql = ssql & " financiamento         with(nolock)," & vbCrLf
    ssql = ssql & " financiamentoParcela  with(nolock)," & vbCrLf
    ssql = ssql & " responsavelFinanceiro with(nolock)" & vbCrLf
    ssql = ssql & " Where" & vbCrLf
    ssql = ssql & "     empresaCliente.ID_EMPRESACLIENTE = financiamento.ID_EMPRESACLIENTE" & vbCrLf
    ssql = ssql & " and financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " and empresaCliente.ID_EMPRESACLIENTE = responsavelFinanceiro.ID_EMPRESACLIENTE" & vbCrLf
    ssql = ssql & " and financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    ssql = ssql & " and financiamentoParcela.DATA_PAGAMENTO is null " & vbCrLf
    'ssql = ssql & "      (financiamentoParcela.DATA_PAGAMENTO is not null and financiamentoParcela.VALOR_RECEBIDO = 0.00))" & vbCrLf
    ssql = ssql & " and financiamentoParcela.DATA_VENCIMENTO <= '" & dtCobranca & "'" & vbCrLf
    ssql = ssql & " and financiamento.ID_ROTA = " & idRota & vbCrLf
    ssql = ssql & " Group By" & vbCrLf
    ssql = ssql & "     empresaCliente.ID_EMPRESACLIENTE," & vbCrLf
    ssql = ssql & "     empresaCliente.TIPO," & vbCrLf
    ssql = ssql & "     empresaCliente.NOME_FANTASIA," & vbCrLf
    ssql = ssql & "     empresaCliente.NOME_PESSOA_FISICA," & vbCrLf
    ssql = ssql & "     financiamento.ID_FINANCIAMENTO," & vbCrLf
    ssql = ssql & "     financiamento.SALDO_DEVEDOR," & vbCrLf
    ssql = ssql & "     financiamento.QTD_PARCELAS," & vbCrLf
    ssql = ssql & "     financiamento.VALOR_PARCELA," & vbCrLf
    ssql = ssql & "     financiamento.DATA_PRIMEIRA_PARCELA," & vbCrLf
    ssql = ssql & "     financiamento.DATA_ULTIMA_PARCELA," & vbCrLf
    ssql = ssql & "     responsavelFinanceiro.DATA_NASCIMENTO," & vbCrLf
    ssql = ssql & "     responsavelFinanceiro.TELEFONE1," & vbCrLf
    ssql = ssql & "     financiamentoParcela.ID_FINANCIAMENTO," & vbCrLf
    ssql = ssql & "     financiamentoParcela.NUM_PARCELA" & vbCrLf
    ssql = ssql & " order By empresaCliente.NOME_PESSOA_FISICA " & vbCrLf
    

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consultaCobranca = mRs
  
    Set mCn = Nothing
    Exit Function
    
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function recuperaParcelasClienteEmpresa(ByVal idEmpresaCliente As Long) As ADODB.Recordset

Dim ssql As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open

    ssql = ""
    ssql = ssql & " select financiamentoParcela.*" & vbCrLf
    ssql = ssql & " From financiamento, financiamentoParcela" & vbCrLf
    ssql = ssql & " Where financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " and financiamento.ID_EMPRESACLIENTE = " & idEmpresaCliente & vbCrLf
    ssql = ssql & " and financiamento.SALDO_DEVEDOR > 0" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperaParcelasClienteEmpresa = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function recuperaParcelasDoDiaPorRota(ByVal strData As String, ByVal lngIdRota As Long) As ADODB.Recordset

Dim ssql As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open

    ssql = ""
    ssql = ssql & " set dateformat dmy" & vbCrLf
    ssql = ssql & " select" & vbCrLf
    ssql = ssql & " empresaCliente.ID_EMPRESACLIENTE," & vbCrLf
    ssql = ssql & " case when empresaCliente.NOME_FANTASIA  is null or empresaCliente.NOME_FANTASIA = '' then" & vbCrLf
    ssql = ssql & "   empresaCliente.NOME_PESSOA_FISICA else" & vbCrLf
    ssql = ssql & "   empresaCliente.RAZAO_SOCIAL" & vbCrLf
    ssql = ssql & " end as EMPRESA," & vbCrLf
    ssql = ssql & " financiamentoParcela.*" & vbCrLf
    ssql = ssql & " From financiamento, financiamentoParcela, empresaCliente" & vbCrLf
    ssql = ssql & " Where financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " and   financiamento.ID_EMPRESACLIENTE = empresaCliente.ID_EMPRESACLIENTE" & vbCrLf
    ssql = ssql & " and financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    ssql = ssql & " and financiamentoParcela.DATA_VENCIMENTO = '" & strData & "'" & vbCrLf
    ssql = ssql & " and financiamentoParcela.VALOR_RECEBIDO is null" & vbCrLf
    ssql = ssql & " and financiamento.ID_ROTA = " & lngIdRota & vbCrLf
    ssql = ssql & " order by empresa" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperaParcelasDoDiaPorRota = mRs
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function recuperaParcelasFinanciamento(ByVal idFinanciamento As Long) As ADODB.Recordset

Dim ssql As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open

    ssql = ""
    ssql = ssql & " select financiamentoParcela.*" & vbCrLf
    ssql = ssql & " From financiamento, financiamentoParcela" & vbCrLf
    ssql = ssql & " Where financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " and financiamento.ID_FINANCIAMENTO = " & idFinanciamento & vbCrLf
    ssql = ssql & " and financiamento.SALDO_DEVEDOR > 0" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set recuperaParcelasFinanciamento = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function gerarNovoVencimentoParcela(ByVal idFinanciamento As Long) As ADODB.Recordset

Dim ssql As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open

    ssql = ""
    ssql = ssql & " select Dateadd(day,1,MAX(financiamentoParcela.DATA_VENCIMENTO))" & vbCrLf
    ssql = ssql & " From financiamento with(nolock), financiamentoParcela with(nolock)" & vbCrLf
    ssql = ssql & " Where financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " and financiamento.ID_FINANCIAMENTO = " & idFinanciamento & vbCrLf
    ssql = ssql & " and financiamento.SALDO_DEVEDOR > 0" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set gerarNovoVencimentoParcela = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function existeParcelaPendenteAntesDesta(ByVal idEmpresaCli As Long, ByVal intParcela As Integer) As Boolean

Dim ssql As String
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open

    ssql = ""
    'verificar se existem parcelas não pagas antes da parcela que se está tentando pagar
    ssql = ssql & " select financiamentoParcela.* from financiamento join financiamentoParcela on financiamento.ID_FINANCIAMENTO = financiamentoParcela.ID_FINANCIAMENTO" & vbCrLf
    ssql = ssql & " where financiamento.ID_EMPRESACLIENTE = " & idEmpresaCli & vbCrLf
    ssql = ssql & "  and financiamento.SALDO_DEVEDOR > 0" & vbCrLf
    ssql = ssql & "  and financiamentoParcela.NUM_PARCELA < " & intParcela & vbCrLf
    ssql = ssql & "  and financiamentoParcela.VALOR_RECEBIDO is null" & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    
    If Not mRs.EOF Then
      existeParcelaPendenteAntesDesta = True
    End If
    
    Set mRs.ActiveConnection = Nothing
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
