VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsROTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Atributos de Classe (PAI)
Public m_01_idRota As Long
Public m_02_idFuncionario As Long
Public m_03_usuarioInclusao As String
Public m_04_dataInclusao As String
Public m_05_usuarioAlteracao As String
Public m_06_dataAlteracao As String
Public m_07_ctLock As Long
'Atributos de Classe (FILHOS)
Public mrsDETALHE As ADODB.Recordset
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

    strSql = ""
    strSql = strSql & "SET DATEFORMAT DMY                                       " & vbCrLf
    strSql = strSql & "INSERT INTO ROTA                                         " & vbCrLf
    strSql = strSql & "            (ID_FUNCIONARIO,                             " & vbCrLf
    strSql = strSql & "             USUARIO_INCLUSAO,                           " & vbCrLf
    strSql = strSql & "             DATA_INCLUSAO,                              " & vbCrLf
    strSql = strSql & "             USUARIO_ALTERACAO,                          " & vbCrLf
    strSql = strSql & "             DATA_ALTERACAO,                             " & vbCrLf
    strSql = strSql & "             CT_LOCK)                                    " & vbCrLf
    strSql = strSql & " VALUES                                                  " & vbCrLf
    strSql = strSql & "        ('" & m_02_idFuncionario & "'                    " & vbCrLf
    strSql = strSql & "        ,'" & m_03_usuarioInclusao & "'                  " & vbCrLf
    strSql = strSql & "        ,GetDate()" & "                                  " & vbCrLf
    strSql = strSql & "        ,'" & m_05_usuarioAlteracao & "'                 " & vbCrLf
    strSql = strSql & "        ,GetDate()" & "                                  " & vbCrLf
    strSql = strSql & "        ," & m_07_ctLock & ")                           " & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    strSql = "SELECT ID_ROTA FROM ROTA WITH(NOLOCK) WHERE ID_FUNCIONARIO = " & m_02_idFuncionario
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    crudInsert = mRs(0)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
        
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsert = 0
End Function
Public Function crudInsertDet() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    With mrsDETALHE
    .MoveFirst
    Do While Not .EOF
    
      strSql = ""
      strSql = strSql & "SET DATEFORMAT DMY                                         " & vbCrLf
      strSql = strSql & "INSERT INTO ROTA_DET                                       " & vbCrLf
      strSql = strSql & "            (ID_ROTA,                                      " & vbCrLf
      strSql = strSql & "             ID_BAIRRO,                                    " & vbCrLf
      strSql = strSql & "             VL_COMISSAO)                                  " & vbCrLf
      strSql = strSql & " VALUES                                                    " & vbCrLf
      strSql = strSql & "        (" & m_01_idRota & "                               " & vbCrLf
      strSql = strSql & "        ," & .Fields("ID_BAIRRO") & "                      " & vbCrLf
      strSql = strSql & "        ," & Replace(.Fields("VL_COMISSAO"), ",", ".") & ")" & vbCrLf
      
      Set mRs = CtxCreateRecordSet(cnConexao, strSql)
      .MoveNext
     
    Loop
    End With
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing

    crudInsertDet = m_01_idRota
        
Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  crudInsertDet = 0
End Function

Public Function GetNovoIDRota() As Long
Dim mRs         As ADODB.Recordset
Dim strSql      As String
Dim cnConexao   As ADODB.Connection

    On Error GoTo trataerro
    
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open

    strSql = ""
    strSql = "SELECT IDENT_CURRENT('ROTA') + IDENT_INCR('ROTA')" & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    GetNovoIDRota = mRs(0)

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
  GetNovoIDRota = 0
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
    
    With mrsDETALHE
    mrsDETALHE.MoveFirst
    
    'Atualiza o valor de comissão em cada registro
    Do While Not .EOF
     strSql = ""
     strSql = strSql & "UPDATE ROTA_DET SET VL_COMISSAO = " & Replace(.Fields("VL_COMISSAO"), ",", ".") & vbCrLf
     strSql = strSql & "WHERE ID_ROTA = " & m_01_idRota & vbCrLf
     strSql = strSql & "  AND ID_BAIRRO = " & .Fields("ID_BAIRRO") & vbCrLf
     Set mRs = CtxCreateRecordSet(cnConexao, strSql)
     .MoveNext
    Loop
    
    End With

    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudUpdate = m_01_idRota
    
Exit Function
trataerro:
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
    
    strSql = ""
    strSql = strSql & "DELETE                              " & vbCrLf
    strSql = strSql & "  FROM ROTA                         " & vbCrLf
    strSql = strSql & " WHERE ID_ROTA = " & m_01_idRota & vbCrLf
    
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudDelete = m_01_idRota

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function crudDeleteDet() As Long
Dim cnConexao       As ADODB.Connection
Dim mRs             As ADODB.Recordset
Dim strSql          As String

    On Error GoTo trataerro
  
    Set cnConexao = CreateObject("ADODB.Connection")
    cnConexao.CommandTimeout = m_timeOut
    cnConexao.ConnectionString = m_stringConexao
    cnConexao.Open
    
    strSql = ""
    strSql = strSql & "DELETE                                         " & vbCrLf
    strSql = strSql & "  FROM ROTA_DET                                " & vbCrLf
    strSql = strSql & " WHERE ID_ROTA = " & m_01_idRota & "          " & vbCrLf
    Set mRs = CtxCreateRecordSet(cnConexao, strSql)
    
    Set mRs.ActiveConnection = Nothing
    Set cnConexao = Nothing
    
    crudDeleteDet = m_01_idRota

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function

Public Function consulta(ByVal lngCodRota As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM rota, rota_det, funcionario WITH(NOLOCK)                                     " & vbCrLf
    ssql = ssql & " WHERE rota.ID_FUNCIONARIO = funcionario.ID_FUNCIONARIO and rota.ID_ROTA = rota_det.ID_ROTA " & vbCrLf
    ssql = ssql & "   AND rota.ID_ROTA = " & lngCodRota & "                                                    " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set consulta = mRs
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function ConsultaPorFuncionario(ByVal idFuncionario As Long) As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM rota WITH(NOLOCK)                                     " & vbCrLf
    ssql = ssql & " WHERE rota.ID_FUNCIONARIO = " & idFuncionario & "                   " & vbCrLf

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set ConsultaPorFuncionario = mRs
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Public Function RecuperarRotas() As ADODB.Recordset
Dim mCn       As ADODB.Connection
Dim mRs       As ADODB.Recordset

    On Error GoTo trataerro
  
    Set mCn = CreateObject("ADODB.Connection")
    mCn.CommandTimeout = m_timeOut
    mCn.ConnectionString = m_stringConexao
    mCn.Open
    
    ssql = ""
    ssql = ssql & " SELECT * FROM rota, funcionario WITH(NOLOCK)"
    ssql = ssql & " WHERE rota.ID_FUNCIONARIO= funcionario.ID_FUNCIONARIO"

    Set mRs = CtxCreateRecordSet(mCn, ssql)
    Set mRs.ActiveConnection = Nothing
    Set RecuperarRotas = mRs
  
    Set mCn = Nothing

Exit Function
trataerro:
  Err.Raise vbObjectError + 1, Err.Source, Err.Description
End Function
Private Sub Class_Initialize()
  InicializaDetalhe
End Sub
Public Sub InicializaDetalhe()
  Set mrsDETALHE = New ADODB.Recordset
  mrsDETALHE.Fields.Append "ID_BAIRRO", adBigInt
  mrsDETALHE.Fields.Append "VL_COMISSAO", adDouble
End Sub

