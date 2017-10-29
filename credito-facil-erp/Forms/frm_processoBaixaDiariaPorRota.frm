VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_processoBaixaDiariaPorRota 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Processo de baixa em parcelas de financiamentos - Por Agente"
   ClientHeight    =   8985
   ClientLeft      =   2340
   ClientTop       =   825
   ClientWidth     =   14490
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_processoBaixaDiariaPorRota.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Lançar Pagamentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12300
      TabIndex        =   11
      Top             =   8610
      Width           =   1935
   End
   Begin VB.TextBox txtSaldo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   5910
      TabIndex        =   7
      Top             =   8730
      Visible         =   0   'False
      Width           =   1905
   End
   Begin FPSpread.vaSpread vasCobranca 
      Height          =   7635
      Left            =   300
      TabIndex        =   4
      Top             =   930
      Width           =   13935
      _Version        =   196608
      _ExtentX        =   24580
      _ExtentY        =   13467
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   12648447
      MaxCols         =   13
      MaxRows         =   25
      ScrollBars      =   2
      SpreadDesigner  =   "frm_processoBaixaDiariaPorRota.frx":058A
      UserResize      =   1
   End
   Begin VB.Frame FraDevedores 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dada de Vencimento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   5250
      TabIndex        =   2
      Top             =   90
      Width           =   3915
      Begin VB.CommandButton cmdExibirRecebiveis 
         Caption         =   "Exibir Recebíveis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1980
         TabIndex        =   9
         Top             =   360
         Width           =   1785
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   2850
         TabIndex        =   5
         Text            =   "ID"
         Top             =   150
         Width           =   915
      End
      Begin VB.ComboBox cboEmpresasDevedores_ 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker DTPickerVencimento 
         Height          =   345
         Left            =   270
         TabIndex        =   8
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         Format          =   17563649
         CurrentDate     =   40758
      End
   End
   Begin VB.Frame FraRotaAgente 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rota / Agente Cobrador"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   4935
      Begin VB.ComboBox cboRotas 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4545
      End
   End
   Begin VB.Label lblRows 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   10
      Top             =   8610
      Width           =   4965
   End
   Begin VB.Label lblSaldo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Saldo do Caixa:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4020
      TabIndex        =   6
      Top             =   8610
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frm_processoBaixaDiariaPorRota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oControle As New ControladorCreditoFacil
Private oParcelas As New clsFinanciamentoParcela
Private ocaixa    As New clsCaixa
Const col_IdParcelamento As Integer = 1
Const col_IdFinanciamento As Integer = 2
Const col_IdEmpresa As Integer = 3
Const col_Empresa As Integer = 4
Const col_NumParcela As Integer = 5
Const col_ValorParcela As Integer = 6
Const col_VencParcela As Integer = 7
Const col_ValorRecebido As Integer = 8
Const col_DataRecebido As Integer = 9
Const col_DiasAtraso As Integer = 10
Const col_SaldoDevedor As Integer = 11
Const col_Situacao As Integer = 12
Const col_botao As Integer = 13


Private Sub PopulaVencimento()

Dim i As Integer
   
   For i = 1 To vasCobranca.MaxRows
      vasCobranca.Col = col_VencParcela
      
   Next
End Sub

Private Function QtdParcelasCliente(id As Long) As Integer

Dim rs As New ADODB.Recordset

   Set rs = oParcelas.recuperaParcelasFinanciamento(id)
   QtdParcelasCliente = rs.RecordCount

End Function

Private Sub cmdExibirRecebiveis_Click()

Dim QtPagas As Integer

   PopulaParcelas
   'QtPagas = ContParcelasPagas
   'If QtPagas >= 1 Then
   '    HabilitaBotaoEstornar (QtPagas)
   'End If
   Command1.Enabled = True

End Sub


Private Sub Command1_Click()

Dim i As Integer
Dim confirmouPagamento As Boolean

Dim ValorRecebido As String
Dim SaldoDevedor As Double
Dim strNomeCliente As String
Dim idEmpresa As Long
Dim blnPagar As Boolean

If MsgBox("Confirma os lançamentos de pagamento de parcelas informados?", vbYesNo) = vbYes Then

   Command1.Enabled = False

   'Percorra todo o grid,
   For i = 1 To vasCobranca.MaxRows
   
      vasCobranca.Col = col_ValorRecebido
      vasCobranca.Row = i
      
      'considerando linhas onde o recebido foi informado.
      If vasCobranca.Text <> "" Then
         'Registra a data do pagamento
         vasCobranca.Col = col_DataRecebido
         'considera linhas onde data recebido esteja vazia
         If vasCobranca.Text = "" Then
            vasCobranca.Text = Format(Now(), "dd/mm/yyyy")
            
            'Não permitir pagamento quando o valor recebido ultrapassar o saldo devedor
            vasCobranca.Col = col_SaldoDevedor
            SaldoDevedor = vasCobranca.Text
            vasCobranca.Col = col_ValorRecebido
            ValorRecebido = vasCobranca.Text
            vasCobranca.Col = col_Empresa
            strNomeCliente = vasCobranca.Text
            vasCobranca.Col = col_IdEmpresa
            idEmpresa = CLng(vasCobranca.Text)
            If ValorRecebido > SaldoDevedor Then
                MsgBox "Valor recebido não pode ultrapassar o saldo devedor do cliente " & strNomeCliente
                Exit Sub
            End If

            With oParcelas
               vasCobranca.Col = col_IdFinanciamento
               .m_02_ID_FINANCIAMENTO = vasCobranca.Text 'ID_FINANCIAMENTO
               vasCobranca.Col = col_NumParcela
               .m_03_NUM_PARCELA = vasCobranca.Text 'Num.Parcela
               vasCobranca.Col = col_ValorParcela
               .m_05_VALOR_COBRADO = vasCobranca.Text 'Valor Cobrado
               vasCobranca.Col = col_ValorRecebido
               .m_07_VALOR_RECEBIDO = vasCobranca.Text 'Valor Recebido
               vasCobranca.Col = col_DataRecebido
               .m_06_DATA_PAGAMENTO = vasCobranca.Text
               .m_13_USUARIO_ALTERACAO = LogInUserID
               vasCobranca.Col = col_VencParcela
               .m_04_DATA_VENCIMENTO = vasCobranca.Text
               vasCobranca.Col = col_IdFinanciamento
               
               'Só registrar o pagamento caso nenhuma parcela anterior estiver pendente
               If oParcelas.existeParcelaPendenteAntesDesta(idEmpresa, .m_03_NUM_PARCELA) Then
                  blnPagar = False
                  MsgBox "Existe parcela anterior a este vencimento sem a respectiva baixa para a empresa " & strNomeCliente & ". Efetue as baixas pendentes."
               Else
                  blnPagar = True
                  'Crítica
                  If (.m_07_VALOR_RECEBIDO < .m_05_VALOR_COBRADO) And .m_03_NUM_PARCELA = QtdParcelasCliente(.m_02_ID_FINANCIAMENTO) Then
                      MsgBox "Na última parcela era esperada a aquitação, como o valor recebido foi inferior ao valor da parcela," & vbCrLf & "será gerada uma nova parcela com juros de mora, sobre o saldo devedor restante!", , "Aviso pagamento cliente: " & strNomeCliente
                      'Exit Sub
                  End If
               End If
              
            End With
               
            If blnPagar Then
               If oParcelas.pagarParcela(txtID, idEmpresa) <> vasCobranca.Text Then
                   MsgBox "Erro ao processar o pagamento da parcela do cliente " & strNomeCliente
                   Exit Sub
               End If
               PintarLinhaPagamento (i)
               txtSaldo = "R$ " & Format(ocaixa.getSaldo(txtID), "0.00")
            End If

         End If
      End If
   Next
End If

End Sub

Private Sub Form_Activate()

If Not ocaixa.CaixaAberto Then
    MsgBox "Para acessar o Processo de Baixa das parcelas, favor abra o caixa."
    Unload Me
Else
    txtID = ocaixa.IdUltimoCaixaAberto
End If

ocaixa.mTIMEOUT = gstrTimeOutGeral
ocaixa.mSTRING_CONEXAO = gstrConexaoCreditoFacil
txtSaldo = "R$ " & CStr(Format(ocaixa.getSaldo(txtID), "0.00"))
DTPickerVencimento.value = Mid(Now, 1, 10)

End Sub

Private Sub Form_Load()

ocaixa.mTIMEOUT = gstrTimeOutGeral
ocaixa.mSTRING_CONEXAO = gstrConexaoCreditoFacil
oParcelas.m_stringConexao = gstrConexaoCreditoFacil
oParcelas.m_timeOut = gstrTimeOutGeral

PopulaRotas

End Sub
Private Function ContParcelasPagas() As Integer
    Dim i As Integer
    
    'Seta a coluna que será lido o valor
    vasCobranca.Col = col_Situacao
    For i = 1 To vasCobranca.MaxRows
        vasCobranca.Row = i
        If vasCobranca.Text = "PAGO" Then
            ContParcelasPagas = ContParcelasPagas + 1
        End If
    Next
    
End Function
Private Sub PopulaRotas()

Dim rs As ADODB.Recordset

oControle.oRota.m_timeOut = gstrTimeOutGeral
oControle.oRota.m_stringConexao = gstrConexaoCreditoFacil
Set rs = oControle.oRota.RecuperarRotas
cboRotas.Clear
Do While Not rs.EOF
  cboRotas.AddItem rs("NOME")
  cboRotas.ItemData(cboRotas.NewIndex) = rs("ID_ROTA")
  rs.MoveNext
Loop

cboRotas.ListIndex = -1

End Sub
Private Sub PopulaParcelas()

    Dim rs As ADODB.Recordset
    Dim i As Integer
    vasCobranca.MaxRows = 0
    
    oParcelas.m_timeOut = gstrTimeOutGeral
    oParcelas.m_stringConexao = gstrConexaoCreditoFacil
    Set rs = oParcelas.recuperaParcelasDoDiaPorRota(DTPickerVencimento.value, _
                                                    cboRotas.ItemData(cboRotas.ListIndex))
     
    If rs.EOF Then lblRows.Caption = "Não há recebíveis ou pendências para esta data": Exit Sub
    vasCobranca.MaxRows = rs.RecordCount
    lblRows.Caption = "(" & rs.RecordCount & ") recebíveis pendentes nesta data"
    rs.MoveFirst
     
    For i = 1 To rs.RecordCount
        
        'If i = 19 Then Stop
        vasCobranca.Row = i
        vasCobranca.RowHeight(i) = 16
        vasCobranca.Col = col_IdParcelamento
        vasCobranca.Text = CLng(rs("ID_PARCELAMENTO"))
                
        vasCobranca.Col = col_IdFinanciamento
        vasCobranca.Text = CLng(rs("ID_FINANCIAMENTO"))
        
        vasCobranca.Col = col_IdEmpresa
        vasCobranca.Text = CLng(rs("ID_EMPRESACLIENTE"))
        
        vasCobranca.Col = col_Empresa
        vasCobranca.Text = rs("EMPRESA")
        
        vasCobranca.Col = col_NumParcela
        vasCobranca.Text = rs("NUM_PARCELA")
                
        vasCobranca.Col = col_ValorParcela
        vasCobranca.Text = Format(rs("VALOR_COBRADO"), "0.00")
        If IsNull(rs("DATA_PAGAMENTO")) Then  'Se já foi pago fica azul, senão fica vermelha
            vasCobranca.ForeColor = vbRed
        Else
            vasCobranca.ForeColor = vbBlue
        End If
        
        vasCobranca.Col = col_VencParcela
        vasCobranca.Text = Format(rs("DATA_VENCIMENTO"), "dd/mm/yyyy")
        
        vasCobranca.Col = col_ValorRecebido
        vasCobranca.Text = IIf(IsNull(rs("VALOR_RECEBIDO")), "", Format(rs("VALOR_RECEBIDO"), "0.00"))
        If IsNull(rs("DATA_PAGAMENTO")) Then  'Se já foi pago fica azul, senão fica vermelha
            'vasCobranca.ForeColor = vbRed
        Else
            vasCobranca.ForeColor = vbBlue
            vasCobranca.CellType = CellTypeStaticText
            vasCobranca.TypeHAlign = TypeHAlignRight
        End If
        
        vasCobranca.Col = col_DataRecebido
        
        If Not IsNull(rs("DATA_PAGAMENTO")) Then
            vasCobranca.CellType = CellTypeStaticText
            vasCobranca.TypeHAlign = TypeHAlignCenter
            vasCobranca.ForeColor = vbBlue
        End If
        vasCobranca.Text = IIf(IsNull(rs("DATA_PAGAMENTO")), "", Format(rs("DATA_PAGAMENTO"), "dd/mm/yyyy"))
        
        'Dias de atraso
        vasCobranca.Col = col_DiasAtraso
        'Se já foi pago, considerar a data gravada, senão recalcular
        If Not IsNull(rs("DATA_PAGAMENTO")) Then
            vasCobranca.Text = rs("DIAS_ATRASO")
        Else
            vasCobranca.Text = CStr(calculaAtrasoPagamento(i))
        End If
        
        vasCobranca.Col = col_SaldoDevedor
        vasCobranca.Text = Format(rs("SALDO_DEVEDOR"), "0.00")
                
        'Se já foi pago, desfaz tipo botão e mostra apenas label q ja foi pago
        If Not IsNull(rs("DATA_PAGAMENTO")) Then
            vasCobranca.Col = col_botao
            vasCobranca.TypeButtonText = ""
            'vasCobranca.CellType = CellTypeStaticText
            'vasCobranca.ForeColor = vbBlack
            'vasCobranca.Text = ""
            'vasCobranca.TypeHAlign = TypeHAlignCenter
            'vasCobranca.TypeVAlign = TypeVAlignCenter
            'Muda a cor da linha inteira
            vasCobranca.BackColor = vbGray
            
            'informa no campo status que ja foi PAGO
            vasCobranca.Col = col_Situacao
            vasCobranca.Text = "PAGO"
            vasCobranca.TypeHAlign = TypeHAlignCenter
            vasCobranca.TypeVAlign = TypeVAlignCenter
            'Muda a cor da linha inteira
            vasCobranca.BackColor = vbCyan
        End If
        rs.MoveNext
    Next
End Sub
Private Sub HabilitaBotaoEstornar(ByVal Row)
    
    'Posiciona na coluna que será modificada
    vasCobranca.Col = col_botao
    vasCobranca.Row = Row
    vasCobranca.CellType = CellTypeButton
    vasCobranca.ForeColor = 0
    vasCobranca.TypeButtonText = "Estornar"
    vasCobranca.TypeButtonColor = vbYellow
    vasCobranca.Refresh

End Sub

Private Function calculaAtrasoPagamento(ByVal Row As Integer) As Integer

     Dim DiaVencimento As Date
     Dim DiaPagamento As Date
 
     vasCobranca.Row = Row
     vasCobranca.Col = col_VencParcela
     DiaVencimento = vasCobranca.Text
     vasCobranca.Col = col_DataRecebido
     DiaPagamento = IIf(vasCobranca.Text = "", Format(Now(), "dd/mm/yyyy"), Format(vasCobranca.Text, "dd/mm/yyyy"))
     calculaAtrasoPagamento = DiaPagamento - DiaVencimento
     If calculaAtrasoPagamento < 0 Then calculaAtrasoPagamento = 0
     vasCobranca.Col = col_DiasAtraso
     If calculaAtrasoPagamento > 0 Then vasCobranca.ForeColor = vbRed Else vasCobranca.ForeColor = vbBlack

End Function

Private Sub vasCobranca_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim OriginalCol As Integer
    Dim ValorReceber As String
    Dim DataRecebimento As String
    Dim ValorRecebido As String
    Dim SaldoDevedor As Double
    
    OriginalCol = Col
    vasCobranca.Row = Row
    
    vasCobranca.Col = col_ValorParcela
    ValorReceber = vasCobranca.Text
    vasCobranca.Col = col_ValorRecebido
    ValorRecebido = vasCobranca.Text
    vasCobranca.Col = col_DataRecebido
    DataRecebimento = vasCobranca.Text
    vasCobranca.Col = col_SaldoDevedor
    SaldoDevedor = vasCobranca.Text
    
    'vasCobranca.Col = col_Situacao
    'If vasCobranca.Text = "PAGO" Then Exit Sub
    
    'If Row > 1 Then
    '    If ExistePendenciaPagamentoAntesDestaParcela(Row) Then
    '        MsgBox "Existe parcela pendente antes desta que você está tentando pagar, dê a baixa na sequência de parcelas!"
    '        vasCobranca.Col = col_botao
    '        vasCobranca.Row = Row
    '        vasCobranca.TypeButtonText = "Pagar"
    '        vasCobranca.Col = col_ValorRecebido
    '        vasCobranca.Text = ""
    '        vasCobranca.Col = col_DataRecebido
    '        vasCobranca.Text = ""
    '        Exit Sub
    '    End If
    'End If
    
    OriginalCol = Col
    vasCobranca.Row = Row
    vasCobranca.Col = Col
    
    
    If OriginalCol = col_botao Then
        
        If vasCobranca.TypeButtonText = "Confirmar" Then
        
            'Se já informou a valor recebido e a data de recebimento, confirma a operação de pagamento
            vasCobranca.Col = col_botao
            If IsDate(DataRecebimento) And IsNumeric(ValorRecebido) And vasCobranca.TypeButtonText = "Confirmar" Then
            
                'Não permitir pagamento quando o valor recebido ultrapassar o saldo devedor
                If ValorRecebido > SaldoDevedor Then
                    MsgBox "Valor recebido não pode ultrapassar o saldo devedor."
                    Exit Sub
                End If
                If MsgBox("Confirma o pagamento da parcela?", vbYesNo) = vbYes Then
                    
                  With oParcelas
                    'MsgBox "Pagando..."
                    vasCobranca.Col = col_IdFinanciamento
                    .m_02_ID_FINANCIAMENTO = vasCobranca.Text 'ID_FINANCIAMENTO
                    vasCobranca.Col = col_NumParcela
                    .m_03_NUM_PARCELA = vasCobranca.Text 'Num.Parcela
                    vasCobranca.Col = col_ValorParcela
                    .m_05_VALOR_COBRADO = vasCobranca.Text 'Valor Cobrado
                    vasCobranca.Col = col_ValorRecebido
                    .m_07_VALOR_RECEBIDO = vasCobranca.Text 'Valor Recebido
                    vasCobranca.Col = col_DataRecebido
                    .m_06_DATA_PAGAMENTO = vasCobranca.Text
                    .m_13_USUARIO_ALTERACAO = LogInUserID
                    vasCobranca.Col = col_IdFinanciamento
                    
                    'Crítica
                    If (.m_07_VALOR_RECEBIDO < .m_05_VALOR_COBRADO) And .m_03_NUM_PARCELA = vasCobranca.MaxRows Then
                        MsgBox "Na última parcela era esperada a aquitação, como o valor recebido foi inferior ao valor da parcela, será gerada uma nova parcela com juros de mora, sobre o saldo devedor restante!"
                        'Exit Sub
                    End If
                  End With
                    
                    If oParcelas.pagarParcela(txtID, cboEmpresasDevedores.ItemData(cboEmpresasDevedores.ListIndex)) <> vasCobranca.Text Then
                        MsgBox "Erro ao processar o pagamento da parcela"
                        Exit Sub
                    End If
                    
                Else
                    Exit Sub
                End If
            End If
        ElseIf vasCobranca.TypeButtonText = "Estornar" Then
            If MsgBox("Deseja estornar este pagamento?", vbYesNo) = vbYes Then
                Dim idFinanc As Long
                Dim numParc As Integer
                Dim Valor As Double
                vasCobranca.Col = col_IdFinanciamento
                idFinanc = vasCobranca.Text
                vasCobranca.Col = col_NumParcela
                numParc = vasCobranca.Text
                vasCobranca.Col = col_ValorRecebido
                Valor = vasCobranca.Text
                Call oParcelas.estornarParcela(txtID, idFinanc, numParc, Valor, cboEmpresasDevedores.ItemData(cboEmpresasDevedores.ListIndex))
                'MsgBox "Estorno realizado!"
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        
        txtSaldo = "R$ " & Format(ocaixa.getSaldo(txtID), "0.00")
        Command1.Enabled = True
    End If
    
End Sub
Private Function ExistePendenciaPagamentoAntesDestaParcela(ByVal Row As Long) As Boolean

Dim i As Integer

For i = 1 To Row - 1
    vasCobranca.Row = i
    vasCobranca.Col = col_Situacao
    If InStr(1, vasCobranca.Text, "PAGO") = 0 Then
        ExistePendenciaPagamentoAntesDestaParcela = True
        Exit Function
    End If
Next

End Function

Private Sub vasCobranca_Change(ByVal Col As Long, ByVal Row As Long)
    If Col = col_DataRecebido Then
        vasCobranca.Col = col_DiasAtraso
        vasCobranca.Text = calculaAtrasoPagamento(Row)
        Col = col_DataRecebido
    End If
End Sub

Private Sub vasCobranca_Click(ByVal Col As Long, ByVal Row As Long)

Dim OriginalCol As Integer
Dim ValorReceber As String
Dim ValorRecebido As String
Dim DataRecebimento As String

OriginalCol = Col

vasCobranca.Row = Row

vasCobranca.Col = col_ValorParcela
ValorReceber = vasCobranca.Text
vasCobranca.Col = col_ValorRecebido
ValorRecebido = vasCobranca.Text
vasCobranca.Col = col_DataRecebido
DataRecebimento = vasCobranca.Text

vasCobranca.Col = col_Situacao
If vasCobranca.Text = "Pago" Then Exit Sub

'Se clicou no botão 'Pagar'
If OriginalCol = col_botao Then
    'Se o valor recebido não foi preenchido, auto-preenche com o valor da parcela
    'Se a data de recebimento não foi preenchida, auto-preenche com a data atual
    If ValorRecebido = "" Or DataRecebimento = "" Then
        
        If ValorRecebido = "" Then
            vasCobranca.Col = col_ValorRecebido
            vasCobranca.Text = ValorReceber
            'cmdCobaia.SetFocus
        End If
        
        If DataRecebimento = "" Then
            vasCobranca.Col = col_DataRecebido
            vasCobranca.Text = Format(Now(), "dd/mm/yyyy")
        End If
        
        vasCobranca.Col = col_NumParcela
        vasCobranca.Action = ActionActiveCell
        vasCobranca.Refresh
        
        vasCobranca.Col = col_botao
        vasCobranca.TypeButtonText = "Confirmar"
        vasCobranca.Action = ActionDeselectBlock
        
        Exit Sub
        
    End If

End If

End Sub

Private Sub PintarLinhaPagamento(Row As Integer)

Dim color As Variant
color = vbGreen

   vasCobranca.Row = Row
   vasCobranca.Col = col_Empresa: vasCobranca.BackColor = color
   vasCobranca.Col = col_NumParcela: vasCobranca.BackColor = color
   vasCobranca.Col = col_ValorParcela: vasCobranca.BackColor = color
   vasCobranca.Col = col_VencParcela: vasCobranca.BackColor = color
   vasCobranca.Col = col_ValorRecebido: vasCobranca.BackColor = color
   vasCobranca.Col = col_DataRecebido: vasCobranca.BackColor = color
   vasCobranca.Col = col_DiasAtraso: vasCobranca.BackColor = color
   vasCobranca.Col = col_SaldoDevedor: vasCobranca.BackColor = color
   vasCobranca.Col = col_Situacao: vasCobranca.BackColor = color
   
End Sub
