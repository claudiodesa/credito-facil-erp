Attribute VB_Name = "mldGlobal"
Option Explicit

Public Function CtxCreateRecordSet(strConexao As Variant, Sql As String) As ADODB.Recordset

    Set CtxCreateRecordSet = CreateObject("ADODB.RecordSet")
    CtxCreateRecordSet.CursorLocation = adUseClient
    CtxCreateRecordSet.CursorType = adOpenStatic
    CtxCreateRecordSet.LockType = adLockBatchOptimistic
    If Sql <> "" Then
       CtxCreateRecordSet.ActiveConnection = strConexao
       CtxCreateRecordSet.Open Sql
    End If

End Function

