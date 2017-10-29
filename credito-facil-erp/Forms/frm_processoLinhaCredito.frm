VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_processoLinhaCredito 
   Caption         =   "Linhas de Crédito - Pré-aprovação"
   ClientHeight    =   7170
   ClientLeft      =   5130
   ClientTop       =   4095
   ClientWidth     =   9675
   Icon            =   "frm_processoLinhaCredito.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9675
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraLinhaCredito 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Linha de Crédito"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   30
      TabIndex        =   11
      Top             =   1890
      Width           =   9585
      Begin VB.TextBox txtDataInicioOperacao 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   7650
         MaxLength       =   14
         TabIndex        =   34
         Top             =   570
         Width           =   1755
      End
      Begin VB.Frame FraResponsavelFinanceiro 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Responsável Financeiro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2265
         Left            =   120
         TabIndex        =   22
         Top             =   990
         Width           =   9345
         Begin VB.Frame FraPropriedadesImoveis 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Propriedades / Imóveis"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1065
            Left            =   90
            TabIndex        =   29
            Top             =   1110
            Width           =   3615
            Begin VB.TextBox txtResideDesde 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
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
               Left            =   2160
               MaxLength       =   14
               TabIndex        =   33
               Top             =   540
               Width           =   1365
            End
            Begin VB.ComboBox cboTipoImovel 
               BackColor       =   &H8000000F&
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
               ItemData        =   "frm_processoLinhaCredito.frx":058A
               Left            =   90
               List            =   "frm_processoLinhaCredito.frx":058C
               Style           =   1  'Simple Combo
               TabIndex        =   30
               Top             =   540
               Width           =   1965
            End
            Begin VB.Label lblTipoImovel 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Imóvel"
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
               Left            =   90
               TabIndex        =   32
               Top             =   300
               Width           =   1185
            End
            Begin VB.Label lblResideDesde 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Reside Desde"
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
               Left            =   2160
               TabIndex        =   31
               Top             =   300
               Width           =   1335
            End
         End
         Begin VB.TextBox txtSituacao 
            BackColor       =   &H8000000F&
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
            Left            =   7890
            MaxLength       =   14
            TabIndex        =   27
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtCPF_Responsavel 
            BackColor       =   &H8000000F&
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
            Left            =   5280
            MaxLength       =   14
            TabIndex        =   24
            Top             =   720
            Width           =   2445
         End
         Begin VB.TextBox txtNomeResposavelFinanceiro 
            BackColor       =   &H8000000F&
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
            Left            =   90
            MaxLength       =   50
            TabIndex        =   23
            Top             =   720
            Width           =   4965
         End
         Begin VB.Label lblSituacao 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Situação"
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
            Left            =   7890
            TabIndex        =   28
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label lblCPF 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CPF"
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
            Left            =   5280
            TabIndex        =   26
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblNomeCompleto 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Completo"
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
            Left            =   90
            TabIndex        =   25
            Top             =   480
            Width           =   1485
         End
      End
      Begin VB.CheckBox chkAprovado 
         Caption         =   "Limite Aprovado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7560
         TabIndex        =   3
         Top             =   4590
         Width           =   1845
      End
      Begin VB.CheckBox chkPossuiFinanciamento 
         Caption         =   "Possui financiamento em aberto"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4590
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox txtLimiteCredito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   3870
         Width           =   1635
      End
      Begin VB.TextBox txtVendaDiaria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   510
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "0,00"
         Top             =   3330
         Width           =   1635
      End
      Begin VB.ComboBox cboRamo 
         BackColor       =   &H8000000F&
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
         Left            =   3270
         Style           =   1  'Simple Combo
         TabIndex        =   15
         Top             =   570
         Width           =   4305
      End
      Begin VB.Frame FraTipoEmpresa 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo Empresa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3105
         Begin VB.OptionButton OptPessoaFisica 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pessoa física"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   14
            Top             =   270
            Width           =   1395
         End
         Begin VB.OptionButton OptPessoaJuridica 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pessoa jurídica"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1500
            TabIndex        =   13
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Label lblLimiteCredito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Crédito (R$)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5070
         TabIndex        =   20
         Top             =   3960
         Width           =   2685
      End
      Begin VB.Label lblVendaDiaria 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Faturamento diário (R$)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   19
         Top             =   3420
         Width           =   2865
      End
      Begin VB.Label lblRamoAtuacao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ramo atuação"
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
         Left            =   3270
         TabIndex        =   17
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblEmAtividade 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Em atividade desde"
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
         Left            =   7650
         TabIndex        =   16
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Empresa Cliente"
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
      Left            =   30
      TabIndex        =   5
      Top             =   690
      Width           =   9615
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   405
         Left            =   8040
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   6795
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   8670
         TabIndex        =   6
         Text            =   "ID"
         Top             =   150
         Width           =   885
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
         TabIndex        =   10
         Top             =   330
         Width           =   885
      End
      Begin VB.Label lblNomeFantasia 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         TabIndex        =   9
         Top             =   330
         Width           =   1935
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2070
      Top             =   6480
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
            Picture         =   "frm_processoLinhaCredito.frx":058E
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLinhaCredito.frx":0B28
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLinhaCredito.frx":10C2
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLinhaCredito.frx":165C
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLinhaCredito.frx":1BF6
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoLinhaCredito.frx":2190
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroLinhaCredito 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
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
      MouseIcon       =   "frm_processoLinhaCredito.frx":272A
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6915
      Width           =   9675
      _ExtentX        =   17066
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
Attribute VB_Name = "frm_processoLinhaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oEmpresaCliente     As New clsEMPRESACLIENTE
Private oLinhaCredito       As New clsLinhaCredito
Private oResponsavel        As New clsResponsavel
Private oRamo               As New clsRAMOATIVIDADE
Private mstrTipoOperacao    As String
Private mCtLock             As Long
Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    
        oEmpresaCliente.m_timeOut = gstrTimeOutGeral
        oEmpresaCliente.m_stringConexao = gstrConexaoCreditoFacil
        Set frmPesquisa.rsResultset = oEmpresaCliente.recuperarEmpresasCliente()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_EMPRESACLIENTE"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "NOME"
        frmPesquisa.Caption = frmPesquisa.Caption & " Empresas Cliente Cadastradas"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtCodigoEntidade_LostFocus
    
    

End Sub
Private Sub Form_Load()
    PopulaRamos
    PopulaTipoImovel
    oEmpresaCliente.m_timeOut = gstrTimeOutGeral
    oEmpresaCliente.m_stringConexao = gstrConexaoCreditoFacil
    oRamo.mTIMEOUT = gstrTimeOutGeral
    oRamo.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    ToolbarCadastroLinhaCredito.Buttons("Excluir").Enabled = False
    ToolbarCadastroLinhaCredito.Buttons("Salvar").Enabled = False
    
End Sub

Private Sub PopulaRamos()

Dim rs As ADODB.Recordset

oRamo.mTIMEOUT = gstrTimeOutGeral
oRamo.mSTRING_CONEXAO = gstrConexaoCreditoFacil
Set rs = oRamo.RecuperarRamos()
cboRamo.Clear
Do While Not rs.EOF
  cboRamo.AddItem rs("DESCRICAO")
  cboRamo.ItemData(cboRamo.NewIndex) = rs("ID_RAMO")
  rs.MoveNext
Loop

cboRamo.ListIndex = -1

End Sub
Private Sub PopulaTipoImovel()
    cboTipoImovel.AddItem "Alugado", 0
    cboTipoImovel.AddItem "Financiado", 1
    cboTipoImovel.AddItem "Próprio", 2
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set oEmpresaCliente = Nothing
Set oLinhaCredito = Nothing
Set oResponsavel = Nothing

End Sub

Private Sub ToolbarCadastroLinhaCredito_ButtonClick(ByVal Button As MSComctlLib.Button)
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

If Not Validar Then
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
ToolbarCadastroLinhaCredito.Buttons("Salvar").Enabled = False

End Sub
'Validar Campos de ClienteEmpresa
Private Function Validar() As Boolean

  If Trim(Len(txtDescricaoEntidade)) = 0 Then
     MsgBox "Deve selecionar uma empresa para gerar ou consultar linha de crédito.", vbInformation, "Mensagem"
     txtCodigoEntidade.SetFocus
     Validar = False
     Exit Function
  End If
  
  If Not IsNumeric(txtLimiteCredito) Then
    MsgBox "O valor do lmite de crédito está incorreto ou não foi preenchido, favor escreva no seguinte formato #.##", vbInformation, "INFORMACÃO"
    txtLimiteCredito.SetFocus
    Validar = False
    Exit Function
  End If
  
  If CCur(txtLimiteCredito) <= 0 Then
    MsgBox "O limite de crédito não pode ser zero ou negativo, favor escreva no seguinte formato #.##", vbInformation, "INFORMACÃO"
    txtLimiteCredito.SetFocus
    Validar = False
    Exit Function
  End If
  
  Validar = True
  
End Function


Private Sub Novo_Click()

oLinhaCredito.mTIMEOUT = gstrTimeOutGeral
oLinhaCredito.mSTRING_CONEXAO = gstrConexaoCreditoFacil
txtID.Text = oLinhaCredito.GetNovoIDLinhaCredito
cmdSelecaoEntidade_Click
'ToolbarCadastroLinhaCredito.Buttons("Salvar").Enabled = True
'mstrTipoOperacao = "I"
'stbmsg.SimpleText = "Incluindo"
DoEvents

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
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
    'On Error GoTo trataerro
    Dim rsEndEmpCli As New ADODB.Recordset
    Dim rsEndResFin As New ADODB.Recordset
    Dim rsResFin As New ADODB.Recordset
    
    
    
    'Atributos da empresaCliente
    With oLinhaCredito
    
        .m_01_idLinhaCredito = txtID.Text
        .m_02_idEmpresaCliente = txtCodigoEntidade
        .m_03_limite = txtLimiteCredito
        .m_04_aprovado = IIf(chkAprovado.value = 1, "S", "N")
        .m_05_usuarioInclusao = LogInUserID
        .m_06_dataInclusao = Now()
        .m_07_usuarioAlteracao = LogInUserID
        .m_08_dataAlteracao = Now()
        .m_09_ctLock = mCtLock
            
    End With
        
    With oLinhaCredito
        .mTIMEOUT = gstrTimeOutGeral
        .mSTRING_CONEXAO = gstrConexaoCreditoFacil
    
        If strOperacao = "I" Then
            txtID.Text = .crudInsert()
        ElseIf strOperacao = "A" Then
            txtID.Text = .crudUpdate()
        Else
            txtID.Text = .crudDelete()
        End If
    
    End With
    
' trataerro
 '   Stop
End Sub
Private Sub Limpacampos()
        
    'Campos da empresaCliente
    txtDescricaoEntidade.Text = ""
    OptPessoaJuridica.value = False
    OptPessoaFisica.value = False
    cboRamo.ListIndex = -1
    txtDataInicioOperacao.Text = ""
    txtVendaDiaria = "0,00"
    txtNomeResposavelFinanceiro = ""
    txtCPF_Responsavel = ""
    txtSituacao = ""
    cboTipoImovel.ListIndex = -1
    txtResideDesde = ""
    txtLimiteCredito = "0,00"
    chkAprovado.value = 0
    chkPossuiFinanciamento.value = 0

End Sub

Private Sub txtCodigoEntidade_Change()

    Limpacampos
    ToolbarCadastroLinhaCredito.Buttons("Excluir").Enabled = False
    ToolbarCadastroLinhaCredito.Buttons("Salvar").Enabled = False
    txtID.Text = ""
    stbmsg.SimpleText = ""
    
End Sub

Private Sub txtCodigoEntidade_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCodigoEntidade_LostFocus()

Dim rsEmpCli As ADODB.Recordset
Dim rsResFin As ADODB.Recordset
Dim rsLinha  As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
        
    'Recupera a empresaCliente
    Set rsEmpCli = oEmpresaCliente.consulta(txtCodigoEntidade)
    
    'Caso não exista a empresaCliente, sair
    If rsEmpCli.EOF Then
        Exit Sub
    End If
    
    txtDescricaoEntidade = rsEmpCli("NOME")
    
    'Recupera o responsavelFinanceiro
    With oResponsavel
        If Not rsEmpCli.EOF Then
            .m_timeOut = gstrTimeOutGeral
            .m_stringConexao = gstrConexaoCreditoFacil
            Set rsResFin = .recuperarResponsavelFicanceiro(rsEmpCli("ID_EMPRESACLIENTE"))
        End If
    End With
    
    If rsResFin.EOF Then Exit Sub
    
    'Recupera linha de crédito
    oLinhaCredito.mTIMEOUT = gstrTimeOutGeral
    oLinhaCredito.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    Set rsLinha = oLinhaCredito.consulta(rsEmpCli("ID_EMPRESACLIENTE"))
        
    If rsLinha.EOF Then
        mstrTipoOperacao = "I"
        stbmsg.SimpleText = "Incluindo"
        txtID = oLinhaCredito.GetNovoIDLinhaCredito()
        ToolbarCadastroLinhaCredito.Buttons("Salvar").Enabled = True
    Else
        txtID = rsLinha("ID_LINHACREDITO")
        ToolbarCadastroLinhaCredito.Buttons("Excluir").Enabled = True
        ToolbarCadastroLinhaCredito.Buttons("Salvar").Enabled = True
        mstrTipoOperacao = "A"
        stbmsg.SimpleText = "Alterando"
    End If
    
    MoveObjetoParaTelaCab rsEmpCli, rsResFin, rsLinha
    txtLimiteCredito.SetFocus
    txtLimiteCredito.SelStart = 0
    txtLimiteCredito.SelLength = Len(txtLimiteCredito)
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByVal rsEmpCli As ADODB.Recordset, _
                                  ByVal rsResFin As ADODB.Recordset, _
                                  ByVal rsLinha As ADODB.Recordset)
    
    Dim i As Integer
    
    'Descarregando dados da empresa
    mCtLock = rsEmpCli("CT_LOCK")
    
    Select Case rsEmpCli("TIPO")
        Case "F" 'Pessoa Física
            OptPessoaFisica.value = True
        Case "J" 'Pessoa Jurídica
            OptPessoaJuridica.value = True
    End Select
        
    For i = 1 To cboRamo.ListCount
        If cboRamo.ItemData(i - 1) = rsEmpCli("ID_RAMO") Then
            cboRamo.ListIndex = i - 1
            Exit For
        End If
    Next
    
    txtDataInicioOperacao = Format(rsEmpCli("INICIOU_ATIVIDADE"), "dd/mm/yyyy")
    txtVendaDiaria = Format(rsEmpCli("VENDA_DIARIA"), "0.00")
    
    'Descarregando dados do responsavel
    If Not rsResFin.EOF Then
      
      txtNomeResposavelFinanceiro = rsResFin("NOME")
      txtSituacao = IIf(rsResFin("SITUACAO") = "A", "Ativo", "Desativado")
      txtCPF_Responsavel = rsResFin("CPF")
      Select Case rsResFin("TIPO_IMOVEL")
        Case "A" 'Alugado
            cboTipoImovel.ListIndex = 0
        Case "F" 'Financiado
            cboTipoImovel.ListIndex = 1
        Case "P" 'Próprio
            cboTipoImovel.ListIndex = 2
      End Select
      txtResideDesde.Text = Format(rsResFin("RESIDE_DESDE"), "dd/mm/yyyy")
      
      'Validando se este limite
      If rsResFin("SITUACAO") = "D" Then
        MsgBox "O representante financeiro desta empresa foi desativo, não será possível alterar o limite realizar empréstimos nestas condições", vbInformation, "RESPONSÁVEL FINANCEIRO INATIVO"
        ToolbarCadastroLinhaCredito.Buttons("SALVAR").Enabled = False
        ToolbarCadastroLinhaCredito.Buttons("EXCLUIR").Enabled = False
      End If
      
      'Descarregando a linha
      If Not rsLinha.EOF Then
        txtLimiteCredito = Format(rsLinha("LIMITE"), "0.00")
        chkAprovado = IIf(Trim(rsLinha("APROVADO")) = "S", 1, 0)
      End If
      
    End If
    
End Sub
Private Sub txtLimiteCredito_Change()
    txtLimiteCredito = Replace(txtLimiteCredito, ".", ",")
    txtLimiteCredito.SelStart = Len(txtLimiteCredito)
End Sub

