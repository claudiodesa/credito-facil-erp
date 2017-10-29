Attribute VB_Name = "Global"
Option Explicit

Public gstrConexaoCreditoFacil As String
Public gstrTimeOutGeral    As String
Public blnLoginOK As Boolean
Public blnLoginAdmin As Boolean
Public gstrVersao As String
Public DBConn As ADODB.Connection
Public LogInUserID As String, LogInUserName As String
Public gstrLogoRel As String
Public gstrEmpresaMestre As Long

Public Function SelectNewID(Cn As ADODB.Connection, ByVal TableName As String, Optional ByVal IDFieldName As String = "ID") As Long
    Dim Request As String, rs As ADODB.Recordset
    Dim NewID As Long
    
    Request = "SELECT MAX(" & IDFieldName & ") FROM " & TableName
    Set rs = Cn.Execute(Request)
    
    If rs Is Nothing Then
        NewID = 1
    Else
        If rs.RecordCount = 0 Then
            NewID = 1
        Else
            rs.MoveFirst
            
            If IsNull(rs.Fields(0).value) Then
                NewID = 1
            Else
                NewID = CLng(rs.Fields(0).value) + 1
            End If
        End If
    End If
    
    SelectNewID = NewID
End Function

Sub Main()

Call Inicializar
'LoadDatabase
'frmSplash.Top = MDICreditoFacil.Height / 2                                'centraliza vertical e
'frmSplash.Left = MDICreditoFacil.Width / 2                                'horizontalmente
'frmSplash.Show vbModal
frmLogin.Show

End Sub


Sub Inicializar()
Dim K As RegObj.RegKey
Dim mDatabase As String, rootkey, strOS As String

On Error GoTo trataerro
  
    gstrTimeOutGeral = 3600
    'Lê do registro do windows a string de conecão com o banco de dados
    Set K = RegObj.RegKeyFromHKey(HKEY_LOCAL_MACHINE)
    rootkey = "SOFTWARE"
    Set K = K.SubKeys(rootkey)
    'MsgBox K.SubKeys(rootkey)
    gstrConexaoCreditoFacil = K.SubKeys("CreditoFacil").Values("ConexaoCreditoFacil").value
    Set K = Nothing
    'MsgBox gstrConexaoCreditoFacil
    'gstrConexaoCreditoFacil = gstrConexaoCreditoFacil '"Provider=SQLOLEDB.1;Password=288744cla;Persist Security Info=True;User ID=sa;Initial Catalog=credito_facil;Data Source=AQUAT888"
    
    Exit Sub
    
trataerro:
Err.Raise Err.Number, , Err.Description
gstrConexaoCreditoFacil = gstrConexaoCreditoFacil '"Provider=SQLOLEDB.1;Password=288744cla;Persist Security Info=True;User ID=sa;Initial Catalog=credito_facil;Data Source=AQUAT888"

End Sub
Public Function DeCriptSenha(Psenha As String) As Variant

Dim v_sqlerrm As String
Dim SenhaCript As String

Dim var1 As String

Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Const chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim to_text As String

to_text = ""
offset = NumericPassword(chave)
Rnd -1
Randomize offset
str_len = Len(Psenha)
For i = 1 To str_len
    ch = Asc(Mid$(Psenha, i, 1))
    If ch >= MIN_ASC And ch <= MAX_ASC Then
        ch = ch - MIN_ASC
        offset = Int((NUM_ASC + 1) * Rnd)
        ch = ((ch - offset) Mod NUM_ASC)
        If ch < 0 Then ch = ch + NUM_ASC
        ch = ch + MIN_ASC
        to_text = to_text & Chr$(ch)
    End If
Next i

DeCriptSenha = to_text
    
End Function
'Funções para criptografar as senhas
Public Function CriptSenha(Psenha As String) As Variant
    Dim v_sqlerrm As String
    Dim SenhaCript As String
    Dim var1 As String
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    Const chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim to_text As String
        
    to_text = ""
    offset = NumericPassword(chave)
    Rnd -1
    Randomize offset
    str_len = Len(Psenha)
    For i = 1 To str_len
        ch = Asc(Mid$(Psenha, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
    
    CriptSenha = to_text
End Function
Private Function NumericPassword(ByVal password As String) As Long
    Dim value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ' Adiciona a próxima letra
        ch = Asc(Mid$(password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = value
End Function

  Public Function extenso(ByVal Valor As Double, ByVal MoedaPlural As String, ByVal MoedaSingular As String) As String
  Dim StrValor As String, Negativo As Boolean
  Dim Buf As String, Parcial As Integer
  Dim Posicao As Integer, Unidades
  Dim Dezenas, Centenas, PotenciasSingular
  Dim PotenciasPlural

  Negativo = (Valor < 0)
  Valor = Abs(CDec(Valor))
  If Valor Then
    Unidades = Array(vbNullString, "Um", "Dois", _
               "Três", "Quatro", "Cinco", _
               "Seis", "Sete", "Oito", "Nove", _
               "Dez", "Onze", "Doze", "Treze", _
               "Quatorze", "Quinze", "Dezesseis", _
               "Dezessete", "Dezoito", "Dezenove")
    Dezenas = Array(vbNullString, vbNullString, _
              "Vinte", "Trinta", "Quarenta", _
              "Cinqüenta", "Sessenta", "Setenta", _
              "Oitenta", "Noventa")
    Centenas = Array(vbNullString, "Cento", _
               "Duzentos", "Trezentos", _
               "Quatrocentos", "Quinhentos", _
               "Seiscentos", "Setecentos", _
               "Oitocentos", "Novecentos")
    PotenciasSingular = Array(vbNullString, " Mil", _
                        " Milhão", " Bilhão", _
                        " Trilhão", " Quatrilhão")
    PotenciasPlural = Array(vbNullString, " Mil", _
                      " Milhões", " Bilhões", _
                      " Trilhões", " Quatrilhões")

    StrValor = Left(Format(Valor, String(18, "0") & _
               ".000"), 18)
    For Posicao = 1 To 18 Step 3
      Parcial = Val(Mid(StrValor, Posicao, 3))
      If Parcial Then
        If Parcial = 1 Then
          Buf = "Um" & PotenciasSingular((18 - _
                Posicao) \ 3)
        ElseIf Parcial = 100 Then
          Buf = "Cem" & PotenciasSingular((18 - _
                Posicao) \ 3)
        Else
          Buf = Centenas(Parcial \ 100)
          Parcial = Parcial Mod 100
          If Parcial <> 0 And Buf <> vbNullString Then
            Buf = Buf & " e "
          End If
          If Parcial < 20 Then
            Buf = Buf & Unidades(Parcial)
          Else
            Buf = Buf & Dezenas(Parcial \ 10)
            Parcial = Parcial Mod 10
            If Parcial <> 0 And Buf <> vbNullString Then
              Buf = Buf & " e "
            End If
            Buf = Buf & Unidades(Parcial)
          End If
          Buf = Buf & PotenciasPlural((18 - Posicao) \ 3)
        End If
        If Buf <> vbNullString Then
          If extenso <> vbNullString Then
            Parcial = Val(Mid(StrValor, Posicao, 3))
            If Posicao = 16 And (Parcial < 100 Or _
                (Parcial Mod 100) = 0) Then
              extenso = extenso & " e "
            Else
              extenso = extenso & ", "
            End If
          End If
          extenso = extenso & Buf
        End If
      End If
    Next
    If extenso <> vbNullString Then
      If Negativo Then
        extenso = "Menos " & extenso
      End If
      If Int(Valor) = 1 Then
        extenso = extenso & " " & MoedaSingular
      Else
        extenso = extenso & " " & MoedaPlural
      End If
    End If
    Parcial = Int((Valor - Int(Valor)) * _
              100 + 0.1)
    If Parcial Then
      Buf = extenso(Parcial, "Centavos", _
            "Centavo")
      If extenso <> vbNullString Then
        extenso = extenso & " e "
      End If
      extenso = extenso & Buf
    End If
  End If
End Function
Public Sub Centraliza(Parent As Form, Child As Form)
Dim iTop As Integer
Dim iLeft As Integer
    'If Parent.WindowState <> 0 Then Exit Sub
      iTop = ((Parent.Height - Child.Height) \ 2)
      iLeft = ((Parent.Width - Child.Width) \ 2)
      Child.Move iLeft, iTop
End Sub
Public Function FinalDeSemana(Data As Date) As Boolean

    If Weekday(Data) = vbSunday Or Weekday(Data) = vbSaturday Then
        FinalDeSemana = True
    End If

End Function
