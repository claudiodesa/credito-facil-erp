VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm_crudRotas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Rotas de Cobrança"
   ClientHeight    =   6030
   ClientLeft      =   6345
   ClientTop       =   4365
   ClientWidth     =   6510
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_crudRotas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   6510
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraCampos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rota de atuação do agente"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   1950
      Width           =   6285
      Begin VB.CommandButton cmdAdicionarItem 
         Caption         =   "[+]"
         Height          =   375
         Left            =   5820
         TabIndex        =   21
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox cboAgente 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   5745
      End
      Begin FPSpread.vaSpread vaSpreadRota 
         Height          =   1725
         Left            =   60
         TabIndex        =   18
         Top             =   1920
         Width           =   5805
         _Version        =   196608
         _ExtentX        =   10239
         _ExtentY        =   3043
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   12648447
         MaxCols         =   4
         MaxRows         =   0
         ScrollBars      =   2
         ShadowColor     =   4210752
         ShadowText      =   16777215
         SpreadDesigner  =   "frm_crudRotas.frx":058A
      End
      Begin VB.TextBox txtValorComissao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   4860
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1440
         Width           =   945
      End
      Begin VB.ComboBox cboBairro 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   2445
      End
      Begin VB.ComboBox cboEstado 
         BackColor       =   &H00E0E0E0&
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cboCidade 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frm_crudRotas.frx":0953
         Left            =   780
         List            =   "frm_crudRotas.frx":0955
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label lblAgente 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label lblValorComissao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comissão"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4860
         TabIndex        =   17
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label lblBairro 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2430
         TabIndex        =   16
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   1170
         Width           =   705
      End
      Begin VB.Label lblCidade 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   780
         TabIndex        =   14
         Top             =   1170
         Width           =   765
      End
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rota"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      Top             =   750
      Width           =   6285
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   375
         Left            =   5580
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   570
         Width           =   465
      End
      Begin VB.TextBox txtCodigoEntidade 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   1
         Top             =   570
         Width           =   915
      End
      Begin VB.TextBox txtDescricaoEntidade 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   10
         Top             =   570
         Width           =   4365
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   5310
         TabIndex        =   9
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         TabIndex        =   12
         Top             =   330
         Width           =   885
      End
      Begin VB.Label lblDescricao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rota do Agente"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   330
         Width           =   1725
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   3750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudRotas.frx":0957
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudRotas.frx":0EF1
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudRotas.frx":148B
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudRotas.frx":1A25
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudRotas.frx":1FBF
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudRotas.frx":2559
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroRota 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   1164
      ButtonWidth     =   1349
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Novo"
            Key             =   "Novo"
            Description     =   "Novo"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Description     =   "Grava o registro atual"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Description     =   "Exclui o registro atual"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Description     =   "Fecha a janela atual"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_crudRotas.frx":2AF3
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5775
      Width           =   6510
      _ExtentX        =   11483
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
End
Attribute VB_Name = "frm_crudRotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCreditoFacil       As New ControladorCreditoFacil
Private mstrTipoOperacao    As String
Private mCtLock             As Long
Private mCtLockEndereco     As Long

Private Sub AdicionarItem(ByVal lngID_BAIRRO As Long, strBairro As String, curValorComissao As Currency)

vaSpreadRota.MaxRows = vaSpreadRota.MaxRows + 1
vaSpreadRota.Row = vaSpreadRota.MaxRows
vaSpreadRota.Col = 1
vaSpreadRota.Text = lngID_BAIRRO
vaSpreadRota.Col = 2
vaSpreadRota.Text = strBairro
vaSpreadRota.Col = 3
vaSpreadRota.Text = curValorComissao

cboBairro.ListIndex = -1
txtValorComissao = "0,00"
cboBairro.SetFocus

End Sub

Private Sub cboBairro_KeyPress(KeyAscii As Integer)

If cboBairro.ListIndex = -1 Then Exit Sub
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cmdAdicionarItem_Click()
  
  If cboEstado.ListIndex = -1 Or cboCidade.ListIndex = -1 Or cboBairro.ListIndex = -1 Then Exit Sub
  If Not IsNumeric(txtValorComissao) Then
    MsgBox "O valor de comissão está incorreto ou não foi preenchido, favor escreva no seguinte formato #.##", vbInformation, "INFORMACÃO"
    txtValorComissao.SetFocus
    Exit Sub
  End If
  If VerificaConstaBairroGrid(cboBairro.ItemData(cboBairro.ListIndex)) Then Exit Sub
  AdicionarItem cboBairro.ItemData(cboBairro.ListIndex), cboBairro.Text, txtValorComissao

End Sub

Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    If Screen.ActiveControl.Name = "cmdSelecaoEntidade" Then
    
        Set frmPesquisa.rsResultset = oCreditoFacil.oRota.RecuperarRotas()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_ROTA"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "NOME"
        frmPesquisa.Caption = frmPesquisa.Caption & " Rotas Cadastradas"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtCodigoEntidade_LostFocus
    
    End If


End Sub

Private Sub PopulaComboEstado()

Dim rs As ADODB.Recordset

Set rs = oCreditoFacil.oEstado.RecuperaEstados()
cboEstado.Clear
Do While Not rs.EOF
  cboEstado.AddItem rs("SIGLA")
  cboEstado.ItemData(cboEstado.NewIndex) = rs("ID_ESTADO")
  rs.MoveNext
Loop

cboEstado.ListIndex = -1

End Sub

Private Sub PopulaComboCidade()

Dim rs As ADODB.Recordset

If cboEstado.ListIndex = -1 Then Exit Sub

Set rs = oCreditoFacil.oMunicipio.RecuperarMunicipios(gstrConexaoCreditoFacil, gstrTimeOutGeral, cboEstado.ItemData((cboEstado.ListIndex)))
cboCidade.Clear
Do While Not rs.EOF
  cboCidade.AddItem rs("DESCRICAO")
  cboCidade.ItemData(cboCidade.NewIndex) = rs("ID_MUNICIPIO")
  rs.MoveNext
Loop

cboCidade.ListIndex = -1

End Sub

Private Sub PopulaComboBairro()

Dim rs As ADODB.Recordset

If cboEstado.ListIndex = -1 Or cboCidade.ListIndex = -1 Then Exit Sub

Set rs = oCreditoFacil.oBairro.RecuperarBairros(gstrConexaoCreditoFacil, gstrTimeOutGeral, cboEstado.ItemData((cboEstado.ListIndex)), cboCidade.ItemData((cboCidade.ListIndex)))
cboBairro.Clear
Do While Not rs.EOF
  cboBairro.AddItem rs("DESCRICAO_BAIRRO")
  cboBairro.ItemData(cboBairro.NewIndex) = rs("ID_BAIRRO")
  rs.MoveNext
Loop

cboBairro.ListIndex = -1

End Sub
Private Sub PopulaComboAgente()

Dim rs As ADODB.Recordset

Set rs = oCreditoFacil.oFuncionario.RecuperarFuncionarios()
cboAgente.Clear
Do While Not rs.EOF
  cboAgente.AddItem rs("NOME")
  cboAgente.ItemData(cboAgente.NewIndex) = rs("ID_FUNCIONARIO")
  rs.MoveNext
Loop

cboAgente.ListIndex = -1

End Sub

Private Sub cboCidade_Click()
  PopulaComboBairro
End Sub

Private Sub cboEstado_Click()
  PopulaComboCidade
End Sub

Private Sub Form_Load()
  
  FraCampos.Visible = False
  oCreditoFacil.oRota.m_timeOut = gstrTimeOutGeral
  oCreditoFacil.oRota.m_stringConexao = gstrConexaoCreditoFacil
  
  oCreditoFacil.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
  oCreditoFacil.oFuncionario.m_timeOut = gstrTimeOutGeral
  
  oCreditoFacil.oEstado.mTIMEOUT = gstrTimeOutGeral
  oCreditoFacil.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
  
  oCreditoFacil.oBairro.mTIMEOUT = gstrTimeOutGeral
  oCreditoFacil.oBairro.mSTRING_CONEXAO = gstrConexaoCreditoFacil
  
  mstrTipoOperacao = ""
  ToolbarCadastroRota.Buttons("Salvar").Enabled = False
  ToolbarCadastroRota.Buttons("Excluir").Enabled = False
  PopulaComboAgente
  PopulaComboEstado
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoFacil = Nothing
End Sub

Private Sub ToolbarCadastroRota_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
  Case 1 'Novo
    Novo_Click
  Case 2 'Salvar
    Salvar_Click
  Case 3 'Excluir
    Excluir_Click
  Case 4 'Fechar
    Fechar_Click
End Select

End Sub
Private Sub Salvar_Click()

'Validações
If txtCodigoEntidade.Text = "" Or txtCodigoEntidade.Text = "0" Then Exit Sub

If Not ValidarInsert Then
    Exit Sub
End If

If mstrTipoOperacao = "A" Then
    If MsgBox("Deseja realmente alterar o registro?", vbYesNo, "Pergunta") = vbNo Then Exit Sub
End If

'Move dados para tela para o objeto
MoveTelaParaObjetoCab mstrTipoOperacao

If txtID.Text = 0 Then Exit Sub

If mstrTipoOperacao = "I" Then
    MsgBox "Registro incluido com sucesso!", vbInformation, "SUCESSO"
ElseIf mstrTipoOperacao = "A" Then
    MsgBox "Registro alterado com sucesso!", vbInformation, "SUCESSO"
End If

Call Limpacampos
txtCodigoEntidade.Text = ""
ToolbarCadastroRota.Buttons("Salvar").Enabled = False

End Sub
Private Sub Novo_Click()

txtCodigoEntidade.Text = oCreditoFacil.oRota.GetNovoIDRota
DoEvents
cboAgente.Enabled = True
cboAgente.SetFocus
ToolbarCadastroRota.Buttons("Salvar").Enabled = True
mstrTipoOperacao = "I"
stbmsg.SimpleText = "Incluindo"


End Sub
Private Sub Excluir_Click()

If txtCodigoEntidade = "" Or txtCodigoEntidade = "0" Then Exit Sub

If MsgBox("Deseja realmente excluir o registro?", vbYesNo, "Pergunta") = vbNo Then Exit Sub

mstrTipoOperacao = "E"

MoveTelaParaObjetoCab mstrTipoOperacao

MsgBox "Registro excluido com sucesso!", vbInformation, "SUCESSO"

Call Limpacampos
txtCodigoEntidade = ""


End Sub
Private Sub Fechar_Click()
  Unload Me
End Sub
'Validar Campos do Formulario
Private Function ValidarInsert() As Boolean

Dim rsFuncao As ADODB.Recordset
    
  If cboAgente.ListIndex = -1 Then
     MsgBox "Você deve selecionar o funcionário (agente) da rota.", vbInformation, "Mensagem"
     cboAgente.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If oCreditoFacil.oFuncionario.FuncionarioPossuiRota(cboAgente.ItemData(cboAgente.ListIndex)) And mstrTipoOperacao = "I" Then
     MsgBox "O funcionário selecionado já possui uma rota criada, você pode consultar as rotas cadastradas e alterar a rota deste.", vbInformation, "Mensagem"
     cboAgente.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If GridVazio Then
     MsgBox "A rota deve constar de pelo menos um bairro associado.", vbInformation, "Mensagem"
     cboEstado.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  ValidarInsert = True
  
End Function
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
    Dim lngI As Long
    
    'Atributos
    oCreditoFacil.oRota.m_01_idRota = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
    oCreditoFacil.oRota.m_02_idFuncionario = cboAgente.ItemData(cboAgente.ListIndex)
    If txtID.Text = "" Then
      oCreditoFacil.oRota.m_03_usuarioInclusao = LogInUserID
      oCreditoFacil.oRota.m_04_dataInclusao = Now
    End If
    oCreditoFacil.oRota.m_05_usuarioAlteracao = LogInUserID
    oCreditoFacil.oRota.m_06_dataAlteracao = Now
    oCreditoFacil.oRota.m_07_ctLock = mCtLock
    
    'Popula detalhes
    oCreditoFacil.oRota.InicializaDetalhe
    oCreditoFacil.oRota.mrsDETALHE.Open
    
    With oCreditoFacil.oRota.mrsDETALHE
        For lngI = 1 To vaSpreadRota.MaxRows
            vaSpreadRota.Col = 1
            vaSpreadRota.Row = lngI
            If Not vaSpreadRota.RowHidden Then
                .AddNew
                .Fields("ID_BAIRRO") = CLng(vaSpreadRota.Text)
                vaSpreadRota.Col = 3
                .Fields("VL_COMISSAO") = Replace(vaSpreadRota.Text, ",", ".")
                oCreditoFacil.oRota.mrsDETALHE.Update
            End If
        Next
    End With
        
   ' oCreditoFacil.oRota.mrsDETALHE.Close
        
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oRota.crudInsert
        txtID.Text = oCreditoFacil.oRota.crudInsertDet
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oRota.crudDeleteDet
        txtID.Text = oCreditoFacil.oRota.crudInsertDet
    Else
        txtID.Text = oCreditoFacil.oRota.crudDeleteDet
        txtID.Text = oCreditoFacil.oRota.crudDelete
    End If
    
End Sub
Private Sub Limpacampos()

FraCampos.Visible = True
txtDescricaoEntidade.Text = ""
cboAgente.ListIndex = -1
cboEstado.ListIndex = -1
cboCidade.ListIndex = -1
cboBairro.ListIndex = -1
txtValorComissao.Text = ""
vaSpreadRota.MaxRows = 0

End Sub

Private Sub txtCodigoEntidade_Change()

Limpacampos
ToolbarCadastroRota.Buttons("Excluir").Enabled = False
ToolbarCadastroRota.Buttons("Salvar").Enabled = False
txtID.Text = ""
stbmsg.SimpleText = ""

End Sub
Private Sub txtCodigoEntidade_LostFocus()

Dim rs As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
        
    oCreditoFacil.oRota.m_timeOut = gstrTimeOutGeral
    oCreditoFacil.oRota.m_stringConexao = gstrConexaoCreditoFacil
    Set rs = oCreditoFacil.oRota.consulta(txtCodigoEntidade)
    
    If rs.EOF Then
      If oCreditoFacil.oRota.GetNovoIDRota <> txtCodigoEntidade Then
        ToolbarCadastroRota.Buttons("Salvar").Enabled = False
      End If
      'txtCodigoEntidade = ""
      Exit Sub
    End If
    
    FraCampos.Visible = True
    ToolbarCadastroRota.Buttons("Excluir").Enabled = True
    ToolbarCadastroRota.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rs
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents

End Sub

Private Sub txtCodigoEntidade_KeyPress(KeyAscii As Integer)
    
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
      KeyAscii = 0
  End If

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rs As ADODB.Recordset)
    
    Dim lngI As Long
    Dim rsBairro As ADODB.Recordset
    
    txtID.Text = rs("ID_ROTA")
    mCtLock = rs("CT_LOCK")
    
    txtDescricaoEntidade = rs("NOME")
    
    For lngI = 1 To cboAgente.ListCount
      If rs("ID_FUNCIONARIO") = cboAgente.ItemData(lngI - 1) Then
        cboAgente.ListIndex = lngI - 1
      End If
    Next
    
    vaSpreadRota.MaxRows = rs.RecordCount
    rs.MoveFirst
    lngI = 0
    Do While Not rs.EOF
      lngI = lngI + 1
      vaSpreadRota.Row = lngI
      vaSpreadRota.Col = 1
      vaSpreadRota.Text = CLng(rs("ID_BAIRRO"))
      vaSpreadRota.Col = 2
      Set rsBairro = oCreditoFacil.oBairro.consulta(rs("ID_BAIRRO"))
      vaSpreadRota.Text = rsBairro("DESCRICAO_BAIRRO")
      vaSpreadRota.Col = 3
      vaSpreadRota.Text = FormatCurrency(rs("VL_COMISSAO"), 2)
      rs.MoveNext
    Loop
    
    cboAgente.Enabled = False
    
End Sub
Private Function GridVazio() As Boolean

  Dim lngCount As Long
  GridVazio = True
  
  If vaSpreadRota.MaxRows = 0 Then GridVazio = True: Exit Function
  
  For lngCount = 1 To vaSpreadRota.MaxRows
    vaSpreadRota.Row = lngCount
    If vaSpreadRota.RowHidden = False Then
      GridVazio = False
      Exit Function
    End If
  Next
  
End Function

Private Sub txtValorComissao_GotFocus()
    txtValorComissao.SelStart = 0
    txtValorComissao.SelLength = Len(txtValorComissao)
End Sub

Private Sub txtValorComissao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdicionarItem_Click
End If
End Sub

Private Sub vaSpreadRota_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If Col = 4 Then
  vaSpreadRota.Row = Row
  vaSpreadRota.RowHidden = True
End If

End Sub
Private Function VerificaConstaBairroGrid(ByVal lngID_BAIRRO As Long) As Boolean

  Dim lngI As Long
  VerificaConstaBairroGrid = False
  vaSpreadRota.Col = 1
  
  If vaSpreadRota.MaxRows = 0 Then Exit Function
  
  For lngI = 1 To vaSpreadRota.MaxRows
    vaSpreadRota.Row = lngI
    If vaSpreadRota.Text = lngID_BAIRRO And vaSpreadRota.RowHidden = False Then
      VerificaConstaBairroGrid = True
      Exit Function
    End If
  Next

End Function
