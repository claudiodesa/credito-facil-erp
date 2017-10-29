VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe
Public m_01_ID_LANCAMENTO As Long
Public m_02_ID_CAIXA As Long
Public m_03_VALOR As Double
Public m_04_TIPO As String
Public m_05_SALDO_ANTERIOR As Double
Public m_06_SALDO_ATUAL As Double
Public m_07_DATA As Date
Public m_08_OBS As String
Public m_09_USUARIO_INCLUSAO As String

'Variável para instruções SQL
Public ssql As String
'Variáveis de Conexão ao Banco de Dados
Public mTIMEOUT As String
Public mSTRING_CONEXAO As String

Public Function adicionar(Optional ByRef cnCon As ADODB.Connection) As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    If cnCon Is Nothing Then
        Set cnConexao = CreateObject("ADODB.Connection")
        cnConexao.CommandTimeout = mTIMEOUT
        cnConexao.ConnectionString = mSTRING_CONEXAO
        cnConexao.Open
    End If

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY" & vbCrLf
    strSql = strSql & "INSERT INTO [credito_facil].[dbo].[caixaLancamento]" & vbCrLf
    strSql = strSql & "           ([ID_CAIXA]" & vbCrLf
    strSql = strSql & "           ,[VALOR]" & vbCrLf
    strSql = strSql & "           ,[TIPO]" & vbCrLf
    strSql = strSql & "           ,[SALDO_ANTERIOR]" & vbCrLf
    strSql = strSql & "           ,[SALDO_ATUAL]" & vbCrLf
    strSql = strSql & "           ,[DATA]" & vbCrLf
    strSql = strSql & "           ,[OBSERVACAO]" & vbCrLf
    strSql = strSql & "           ,[USUARIO_INCLUSAO])" & vbCrLf
    strSql = strSql & "     Values" & vbCrLf
    strSql = strSql & "           ( " & m_02_ID_CAIXA & vbCrLf
    strSql = strSql & "           , " & Replace(m_03_VALOR, ",", ".") & vbCrLf
    strSql = strSql & "           ,'" & m_04_TIPO & "'" & vbCrLf
    strSql = strSql & "           , " & Replace(m_05_SALDO_ANTERIOR, ",", ".") & vbCrLf
    strSql = strSql & "           , " & Replace(m_05_SALDO_ANTERIOR + m_03_VALOR, ",", ".") & vbCrLf
    strSql = strSql & "           , getdate()" & vbCrLf
    strSql = strSql & "           ,'" & m_08_OBS & "'" & vbCrLf
    strSql = strSql & "           ,'" & m_09_USUARIO_INCLUSAO & "')" & vbCrLf
    
    If cnCon Is Nothing Then
        Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        Set mRs.ActiveConnection = Nothing
        strSql = "SELECT @@IDENTITY"
        Set mRs = CtxCreateRecordSet(cnConexao, strSql)
        adicionar = mRs(0)
        Set cnConexao = Nothing
    Else
        Set mRs = CtxCreateRecordSet(cnCon, strSql)
        Set mRs.ActiveConnection = Nothing
        strSql = "SELECT @@IDENTITY"
        Set mRs = CtxCreateRecordSet(cnCon, strSql)
        adicionar = mRs(0)
    End If

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  adicionar = 0
End Function
Public Function GetNovoIdLancamento() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = mTIMEOUT
    cnConexao.ConnectionString = mSTRING_CONEXAO
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('caixaLancamento') + IDENT_INCR('caixaLancamento')" & vbCrLf
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    GetNovoIdLancamento = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIdLancamento = 0
End Function
