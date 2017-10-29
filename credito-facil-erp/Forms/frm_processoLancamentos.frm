VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6CF9344E-3A55-11D5-B99A-0060083D6B0C}#1.0#0"; "UCNumero.ocx"
Begin VB.Form frm_processoLancamentos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Processo Lançamentos de Caixa"
   ClientHeight    =   5265
   ClientLeft      =   1920
   ClientTop       =   3345
   ClientWidth     =   7755
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_processoLancamentos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6540
      TabIndex        =   17
      Text            =   "ID"
      Top             =   3390
      Width           =   915
   End
   Begin VB.Frame fraDetalhesCaixa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalhes do caixa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   180
      TabIndex        =   8
      Top             =   3240
      Width           =   7335
      Begin VB.TextBox txtDataAbertura 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Left            =   330
         MaxLength       =   10
         TabIndex        =   10
         Top             =   630
         Width           =   1725
      End
      Begin UCNumero.ctlNumero ctlNumValorAbertura 
         Height          =   375
         Left            =   5940
         TabIndex        =   9
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         BackColor       =   -2147483633
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Calibri"
         FontSize        =   12
         Mascara         =   "999999,99"
         TipoDeDado      =   1
         MaxValue        =   999999.99
         MinValue        =   -999999.99
      End
      Begin VB.Label lblLabel1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   330
         TabIndex        =   13
         Top             =   1050
         Width           =   6765
      End
      Begin VB.Label lblDataAbertura 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Abertura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   330
         TabIndex        =   12
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label lblSaldoInicial 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldo atual (R$)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5850
         TabIndex        =   11
         Top             =   330
         Width           =   1305
      End
      Begin VB.Image img_saco 
         Height          =   240
         Left            =   5640
         Picture         =   "frm_processoLancamentos.frx":058A
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame FraOperacao 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dados da operação"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Width           =   7335
      Begin VB.TextBox txtObs 
         BackColor       =   &H80000016&
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
         Left            =   210
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1530
         Width           =   6885
      End
      Begin VB.ComboBox cboTipo 
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
         ItemData        =   "frm_processoLancamentos.frx":0B14
         Left            =   210
         List            =   "frm_processoLancamentos.frx":0B1E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker DTPickerData 
         Height          =   405
         Left            =   1830
         TabIndex        =   2
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   94896129
         CurrentDate     =   40762
      End
      Begin UCNumero.ctlNumero ctlNumValor 
         Height          =   375
         Left            =   5910
         TabIndex        =   3
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Calibri"
         FontSize        =   12
         Mascara         =   "999999,99"
         TipoDeDado      =   1
         MaxValue        =   999999.99
         MinValue        =   -999999.99
      End
      Begin VB.Label lblObs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Anotação"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   16
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label lblTipo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   450
         Width           =   855
      End
      Begin VB.Label lblData 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1830
         TabIndex        =   6
         Top             =   450
         Width           =   855
      End
      Begin VB.Label lblValorR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Valor (R$)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6300
         TabIndex        =   5
         Top             =   420
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5010
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3030
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":0B35
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":10CF
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":1669
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":1C03
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":219D
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":2737
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":2CD1
            Key             =   "abre"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":326B
            Key             =   "fecha"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":3805
            Key             =   "Pesquisar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":3D9F
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLancamentos.frx":4339
            Key             =   "Pen"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroFuncao 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Key             =   "Novo"
            Description     =   "Novo"
            Object.ToolTipText     =   "Abertura de Caixa"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lançar"
            Key             =   "Lancar"
            Description     =   "Lançar"
            Object.ToolTipText     =   "Fechar o Caixa"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Description     =   "Fecha a janela atual"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_processoLancamentos.frx":48D3
   End
End
Attribute VB_Name = "frm_processoLancamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oLancamento As New clsCaixaLancamento
Private ocaixa As New clsCaixa
Private mstrTipoOperacao As String


Private Sub Form_Load()

    ocaixa.mTIMEOUT = gstrTimeOutGeral
    ocaixa.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oLancamento.mTIMEOUT = gstrTimeOutGeral
    oLancamento.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    
    If Not ocaixa.CaixaAberto Then
        stbmsg.SimpleText = "O caixa não está aberto, não é possível fazer lançamentos."
        ToolbarCadastroFuncao.Buttons("Novo").Enabled = False
        ToolbarCadastroFuncao.Buttons("Lancar").Enabled = False
        Exit Sub
    End If
    
    CarregarCaixa
    stbmsg.SimpleText = "O caixa está aberto."
    ToolbarCadastroFuncao.Buttons("Lancar").Enabled = False

End Sub
Private Sub CarregarCaixa()

Dim rsCaixa As ADODB.Recordset

    Set rsCaixa = ocaixa.consulta(ocaixa.IdUltimoCaixaAberto)
    MoveObjTela rsCaixa

End Sub
Private Sub MoveObjTela(ByVal rsCaixa As ADODB.Recordset)

    txtDataAbertura = Format(rsCaixa("DATA_ABERTURA"), "dd/mm/yyyy")
    txtID = rsCaixa("ID_CAIXA")
    lblLabel1 = " Aberto por: " & rsCaixa("USUARIO_ABERTURA")
    Call AtualizarSaldo

End Sub
Private Sub ToolbarCadastroFuncao_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "Novo"
        Click_Novo
    Case "Lancar"
        Click_Lancar
    Case "Sair"
        Unload Me

End Select

End Sub
Private Sub Click_Lancar()

    If Not Validar Then Exit Sub
    
    If MsgBox("Confirma Lançamento de " & cboTipo.Text & " no valor de " & Format(IIf(cboTipo.ListIndex = 0, Format(IIf(ctlNumValor.Texto > 0, ctlNumValor.Texto * -1, ctlNumValor.Texto), "Currency"), ctlNumValor.Texto), "Currency") & " ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    MoveTelaObj
    MsgBox "Lançamento efetuado com sucesso!"
    mstrTipoOperacao = ""
    Limpacampos
    ToolbarCadastroFuncao.Buttons("Lancar").Enabled = False
    AtualizarSaldo
        
End Sub
Private Sub MoveTelaObj()

    oLancamento.m_02_ID_CAIXA = txtID
    'Se for depesa e não estiver negativo, converter para negativo
    If cboTipo.Text = "Despesas" Then
        ctlNumValor.Texto = IIf(ctlNumValor.Texto > 0, ctlNumValor.Texto * -1, ctlNumValor.Texto)
    End If
    oLancamento.m_03_VALOR = ctlNumValor.Texto
    oLancamento.m_04_TIPO = IIf(cboTipo.ListIndex = 0, "D", "A") 'Despesa ou Ajuste
    oLancamento.m_05_SALDO_ANTERIOR = ocaixa.getSaldo(txtID)
    oLancamento.m_07_DATA = Format(DTPickerData, "dd/mm/yyyy")
    oLancamento.m_08_OBS = txtObs
    oLancamento.m_09_USUARIO_INCLUSAO = LogInUserID
    oLancamento.adicionar
    
End Sub
Private Sub Limpacampos()

    cboTipo.ListIndex = -1
    DTPickerData = Now
    txtObs = ""
    ctlNumValor.Texto = ""
    stbmsg.SimpleText = ""
    cboTipo.SetFocus

End Sub

Private Sub Click_Novo()

    mstrTipoOperacao = "Novo"
    Limpacampos
    stbmsg.SimpleText = "Inserindo lançamento"
    ToolbarCadastroFuncao.Buttons("Lancar").Enabled = True
        
End Sub
Private Function Validar() As Boolean

    If cboTipo.ListIndex = -1 Then
        MsgBox "Especifique o tipo de lançamento."
        cboTipo.SetFocus
        Exit Function
    End If
    
    If ctlNumValor.Texto = "" Then
        MsgBox "Valor do lançamento é obrigatório"
        ctlNumValor.SetFocus
        Exit Function
    End If
    
    If ctlNumValor.Texto = 0 Then
        MsgBox "Valor do lançamento não pode ser igual a zero"
        ctlNumValor.SetFocus
        Exit Function
    End If
    
    If (cboTipo.ListIndex = 0 And ctlNumValor.Texto > ocaixa.getSaldo(txtID)) Or (ctlNumValor.Texto < 0) And ocaixa.getSaldo(txtID) + ctlNumValor.Texto < 0 Then 'Despesas (lançar valor negativo precisa de saldo)
        MsgBox "Valor do lançamento de " & cboTipo.Text & " é maior que o saldo atual em caixa, que é de: " & Format(ocaixa.getSaldo(txtID), "Currency"), vbInformation, "SALDO INSUFICIENTE"
        ctlNumValor.SetFocus
        Exit Function
    End If
    
    Validar = True
    
End Function
Private Sub AtualizarSaldo()

    ctlNumValorAbertura.Texto = Format(ocaixa.getSaldo(txtID), "0.00")
    ctlNumValorAbertura.ForeColor = IIf(ctlNumValorAbertura.Texto >= 0, vbBlue, vbRed)

End Sub
