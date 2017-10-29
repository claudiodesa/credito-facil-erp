VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControladorCreditoFacil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Classes a que o controller tem acesso
Public oUsuario As New clsUSUARIO
Public oFuncao As New clsFUNCAO
Public oBairro As New clsBAIRRO
Public oFeriado As New clsFERIADO
Public oEmpresa As New clsEMPRESA
Public oEstado As New clsESTADO
Public oMunicipio As New clsMUNICIPIO
Public oEndereco As New clsENDERECO
Public oRota As New clsROTA
Public oFuncionario As New clsFUNCIONARIO
Public oRamo As New clsRAMOATIVIDADE
Public oEmpresaCliente As New clsEMPRESACLIENTE
Public oResponsavel As New clsResponsavel
Public oLinhaCred As New clsLinhaCredito
Public oParcelas As New clsFinanciamentoParcela

Public Function RecuperarUsuario(strLogin, gstrConexaoCreditoFacil, gstrTimeOutGeral) As ADODB.Recordset
    oUsuario.m_timeOut = gstrTimeOutGeral
    oUsuario.m_stringConexao = gstrConexaoCreditoFacil
    Set RecuperarUsuario = oUsuario.ConsultaLogin(strLogin)
End Function

