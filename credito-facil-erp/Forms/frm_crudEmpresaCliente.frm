VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{21F8D070-6EA8-40F7-8555-9E5FA3E03CB5}#1.0#0"; "Calendario.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_crudEmpresaCliente 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Empresas Clientes"
   ClientHeight    =   9810
   ClientLeft      =   3285
   ClientTop       =   2370
   ClientWidth     =   9855
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_crudEmpresaCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
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
      Left            =   60
      TabIndex        =   52
      Top             =   750
      Width           =   9555
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   8610
         TabIndex        =   54
         Top             =   120
         Width           =   885
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
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   570
         Width           =   6795
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
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   570
         Width           =   915
      End
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   315
         Left            =   8010
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   540
         Width           =   465
      End
      Begin VB.Label lblNomeFantasia 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia"
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
         TabIndex        =   56
         Top             =   330
         Width           =   1935
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
         TabIndex        =   55
         Top             =   330
         Width           =   885
      End
   End
   Begin TabDlg.SSTab SSTClienteEmpresa 
      Height          =   7485
      Left            =   60
      TabIndex        =   0
      Top             =   2010
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13203
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cliente Empresa"
      TabPicture(0)   =   "frm_crudEmpresaCliente.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblRamoAtuacao"
      Tab(0).Control(1)=   "lblEmAtividade"
      Tab(0).Control(2)=   "lblVendaDiaria"
      Tab(0).Control(3)=   "ImageDefault"
      Tab(0).Control(4)=   "CtlDataInicioOperacao"
      Tab(0).Control(5)=   "cboRamo"
      Tab(0).Control(6)=   "FraTipoEmpresa"
      Tab(0).Control(7)=   "FraIdentificacaoPF"
      Tab(0).Control(8)=   "txtVendaDiaria"
      Tab(0).Control(9)=   "FraIdentificacaoPJ"
      Tab(0).Control(10)=   "FraEndereco"
      Tab(0).Control(11)=   "FraLOGO"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Responsável Financeiro"
      TabPicture(1)   =   "frm_crudEmpresaCliente.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraEnderecoResponsavel"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraIdentificacaoResponsavel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FraPropriedade"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FraContatos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FraReferênciasPessoais"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtIDResponsavelFinanceiro"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtIDResponsavelFinanceiro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7290
         TabIndex        =   117
         Top             =   510
         Width           =   2175
      End
      Begin VB.Frame FraReferênciasPessoais 
         Caption         =   "Referências Pessoais"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   112
         Top             =   3540
         Width           =   9405
         Begin VB.Frame FraIndicacao 
            Caption         =   "Indicação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1155
            Left            =   5580
            TabIndex        =   118
            Top             =   300
            Width           =   3675
            Begin VB.TextBox txtIndicadoPor 
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   40
               Top             =   570
               Width           =   3345
            End
            Begin VB.Label Label12 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Indicado por"
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
               TabIndex        =   119
               Top             =   330
               Width           =   1215
            End
         End
         Begin VB.TextBox txtFoneContato2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4050
            MaxLength       =   13
            TabIndex        =   39
            Top             =   1080
            Width           =   1425
         End
         Begin VB.TextBox txtNomeContato2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   150
            MaxLength       =   50
            TabIndex        =   38
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtFoneContato1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4050
            MaxLength       =   13
            TabIndex        =   37
            Top             =   480
            Width           =   1425
         End
         Begin VB.TextBox txtNomeContato1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   150
            MaxLength       =   50
            TabIndex        =   36
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fone Contato 2"
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
            Left            =   4050
            TabIndex        =   116
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label lblFonteContato2 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fonte Contato 2"
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
            Left            =   150
            TabIndex        =   115
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label lblTelefoneContato1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fone Contato 1"
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
            Left            =   4050
            TabIndex        =   114
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label lblNomeContato 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Contato 1"
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
            Left            =   150
            TabIndex        =   113
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame FraContatos 
         Caption         =   "Telefones do Responsável Financeiro"
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
         Height          =   945
         Left            =   120
         TabIndex        =   108
         Top             =   2550
         Width           =   6345
         Begin VB.TextBox txtTelefone3 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3300
            MaxLength       =   13
            TabIndex        =   33
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox txtTelefone2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1740
            MaxLength       =   13
            TabIndex        =   32
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox txtTelefone1 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   13
            TabIndex        =   31
            Top             =   480
            Width           =   1485
         End
         Begin VB.Label lblTelefone3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 3"
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
            Left            =   3300
            TabIndex        =   111
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblTelefone2 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 2"
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
            Left            =   1740
            TabIndex        =   110
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblTelefone1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone 1"
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
            Left            =   180
            TabIndex        =   109
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame FraPropriedade 
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
         Height          =   945
         Left            =   6540
         TabIndex        =   105
         Top             =   2550
         Width           =   2985
         Begin VB.ComboBox cboTipoImovel 
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
            ItemData        =   "frm_crudEmpresaCliente.frx":05C2
            Left            =   120
            List            =   "frm_crudEmpresaCliente.frx":05C4
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   480
            Width           =   1365
         End
         Begin Calendario.ctlPicker CtlResideDesde 
            Height          =   315
            Left            =   1560
            TabIndex        =   35
            Top             =   480
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Text            =   "__/__/____"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   195
            Left            =   1560
            TabIndex        =   107
            Top             =   240
            Width           =   1335
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
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FraIdentificacaoResponsavel 
         Caption         =   "Identificação - Responsável Financeiro"
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
         Height          =   2145
         Left            =   120
         TabIndex        =   92
         Top             =   390
         Width           =   9405
         Begin VB.TextBox txtFiliacaoMae 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4020
            MaxLength       =   50
            TabIndex        =   30
            Top             =   1710
            Width           =   5295
         End
         Begin VB.TextBox txtNacionalidade 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   29
            Top             =   1710
            Width           =   1935
         End
         Begin VB.TextBox txtNaturalidade 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   20
            TabIndex        =   28
            Top             =   1710
            Width           =   1935
         End
         Begin VB.ComboBox cboEstadoCivil 
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
            ItemData        =   "frm_crudEmpresaCliente.frx":05C6
            Left            =   8070
            List            =   "frm_crudEmpresaCliente.frx":05C8
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1110
            Width           =   1245
         End
         Begin VB.ComboBox CboSexo 
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
            ItemData        =   "frm_crudEmpresaCliente.frx":05CA
            Left            =   7170
            List            =   "frm_crudEmpresaCliente.frx":05CC
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1110
            Width           =   765
         End
         Begin VB.CheckBox chkDesativado 
            Caption         =   "Desativado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7890
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox txtNomeResposavelFinanceiro 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   50
            TabIndex        =   20
            Top             =   480
            Width           =   5415
         End
         Begin VB.TextBox txtOrgaoEmissor 
            BackColor       =   &H00FFFFFF&
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
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   24
            Top             =   1080
            Width           =   1305
         End
         Begin VB.TextBox txtRG_Responsavel 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            MaxLength       =   20
            TabIndex        =   23
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtCPF_Responsavel 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   14
            TabIndex        =   22
            Top             =   1080
            Width           =   1725
         End
         Begin Calendario.ctlPicker CtlDataExpedicao 
            Height          =   315
            Left            =   5550
            TabIndex        =   25
            Top             =   1110
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Text            =   "__/__/____"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Calendario.ctlPicker CtlDataNascimento 
            Height          =   315
            Left            =   5550
            TabIndex        =   21
            Top             =   510
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Text            =   "__/__/____"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblNomeMae 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Mãe"
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
            Left            =   4020
            TabIndex        =   104
            Top             =   1470
            Width           =   1275
         End
         Begin VB.Label lblNacionalidade 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nacionalidade"
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
            Left            =   2040
            TabIndex        =   103
            Top             =   1470
            Width           =   1275
         End
         Begin VB.Label lblNaturalidade 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Naturalidade"
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
            Left            =   60
            TabIndex        =   102
            Top             =   1470
            Width           =   1185
         End
         Begin VB.Label lblEstadoCivil 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civíl"
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
            Left            =   8070
            TabIndex        =   101
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label lblDataNascimento 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data Nasc."
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
            Left            =   5550
            TabIndex        =   100
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblSexo 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
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
            Left            =   7170
            TabIndex        =   99
            Top             =   900
            Width           =   735
         End
         Begin VB.Label lblDataExpedicao 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Expedição"
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
            Left            =   5550
            TabIndex        =   98
            Top             =   870
            Width           =   1365
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
            Left            =   60
            TabIndex        =   96
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label lblOrgaoEmissor 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Orgão Emissor"
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
            Left            =   4170
            TabIndex        =   95
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label lblRG_Responsavel 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
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
            Left            =   1890
            TabIndex        =   94
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lblCPF_Responsavel 
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
            Left            =   60
            TabIndex        =   93
            Top             =   840
            Width           =   885
         End
      End
      Begin VB.Frame FraEnderecoResponsavel 
         Caption         =   "Endereço do Responsável Financeiro"
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
         Height          =   2205
         Left            =   90
         TabIndex        =   81
         Top             =   5160
         Width           =   9555
         Begin VB.TextBox txtTipoLogradouro_ 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   10
            TabIndex        =   41
            Top             =   510
            Width           =   975
         End
         Begin VB.TextBox txtLogradouro_ 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   42
            Top             =   510
            Width           =   6615
         End
         Begin VB.TextBox txtNumero_ 
            BackColor       =   &H00FFFFFF&
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
            Left            =   7860
            MaxLength       =   12
            TabIndex        =   43
            Top             =   510
            Width           =   1425
         End
         Begin VB.TextBox txtComplemento_ 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   25
            TabIndex        =   44
            Top             =   1110
            Width           =   2775
         End
         Begin VB.TextBox txtPontoReferencia_ 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3090
            MaxLength       =   25
            TabIndex        =   45
            Top             =   1110
            Width           =   6195
         End
         Begin VB.TextBox txtCEP_ 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   9
            TabIndex        =   46
            Top             =   1740
            Width           =   1065
         End
         Begin VB.ComboBox cboBairro_ 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1740
            Width           =   3525
         End
         Begin VB.ComboBox cboEstado_ 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1740
            Width           =   825
         End
         Begin VB.ComboBox cboCidade_ 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2190
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1740
            Width           =   3525
         End
         Begin VB.TextBox txtIDEnderecoResponsavel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   7740
            TabIndex        =   82
            Text            =   "id"
            Top             =   120
            Width           =   1755
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   7860
            TabIndex        =   89
            Top             =   270
            Width           =   1125
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   91
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Logradouro"
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
            Left            =   1230
            TabIndex        =   90
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            TabIndex        =   88
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ponto Referência"
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
            Left            =   3090
            TabIndex        =   87
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            TabIndex        =   86
            Top             =   1500
            Width           =   945
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Left            =   5760
            TabIndex        =   85
            Top             =   1500
            Width           =   945
         End
         Begin VB.Label Label2 
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
            Height          =   225
            Left            =   1320
            TabIndex        =   84
            Top             =   1500
            Width           =   915
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Left            =   2190
            TabIndex        =   83
            Top             =   1500
            Width           =   945
         End
      End
      Begin VB.Frame FraLOGO 
         Caption         =   "LOGOMARCA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   -68220
         TabIndex        =   77
         Top             =   2250
         Width           =   2865
         Begin VB.CommandButton cmdLimpar 
            Caption         =   "Limpar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   1380
            Width           =   915
         End
         Begin VB.CommandButton cmdTrocar 
            Caption         =   "Trocar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtImagePath 
            Height          =   315
            Left            =   30
            TabIndex        =   78
            Top             =   180
            Visible         =   0   'False
            Width           =   2415
         End
         Begin MSComDlg.CommonDialog dlgFOTO 
            Left            =   300
            Top             =   420
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image imgLogoEmpresaCliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   2265
            Left            =   1110
            Picture         =   "frm_crudEmpresaCliente.frx":05CE
            Stretch         =   -1  'True
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.Frame FraEndereco 
         Caption         =   "Endereço da Empresa Cliente"
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
         Height          =   2205
         Left            =   -74910
         TabIndex        =   66
         Top             =   5160
         Width           =   9555
         Begin VB.TextBox txtIDEndereco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   8070
            TabIndex        =   67
            Text            =   "id"
            Top             =   120
            Width           =   1425
         End
         Begin VB.ComboBox cboCidade 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1740
            Width           =   3465
         End
         Begin VB.ComboBox cboEstado 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1740
            Width           =   885
         End
         Begin VB.ComboBox cboBairro 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1740
            Width           =   3525
         End
         Begin VB.TextBox txtCEP 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   9
            TabIndex        =   16
            Top             =   1740
            Width           =   1035
         End
         Begin VB.TextBox txtPontoReferencia 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3090
            MaxLength       =   25
            TabIndex        =   15
            Top             =   1110
            Width           =   6165
         End
         Begin VB.TextBox txtComplemento 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   25
            TabIndex        =   14
            Top             =   1110
            Width           =   2745
         End
         Begin VB.TextBox txtNumero 
            BackColor       =   &H00FFFFFF&
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
            Left            =   7860
            MaxLength       =   12
            TabIndex        =   13
            Top             =   510
            Width           =   1395
         End
         Begin VB.TextBox txtLogradouro 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   12
            Top             =   510
            Width           =   6585
         End
         Begin VB.TextBox txtTipoLogradouro 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   10
            TabIndex        =   11
            Top             =   510
            Width           =   945
         End
         Begin VB.Label lblNumero 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   7860
            TabIndex        =   70
            Top             =   270
            Width           =   945
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
            Height          =   225
            Left            =   2220
            TabIndex        =   76
            Top             =   1470
            Width           =   945
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
            Height          =   225
            Left            =   1320
            TabIndex        =   75
            Top             =   1470
            Width           =   945
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
            Height          =   225
            Left            =   5760
            TabIndex        =   74
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label lblCEP 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            TabIndex        =   73
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label lblPontoReferencia 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ponto Referência"
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
            Left            =   3090
            TabIndex        =   72
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label lblComplemento 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            TabIndex        =   71
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label lblLogradouro 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Logradouro"
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
            Left            =   1230
            TabIndex        =   69
            Top             =   270
            Width           =   945
         End
         Begin VB.Label lblTipo 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   68
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.Frame FraIdentificacaoPJ 
         Caption         =   "Identificação - Pessoa Jurídica"
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
         Height          =   1005
         Left            =   -74850
         TabIndex        =   63
         Top             =   1230
         Width           =   5865
         Begin VB.TextBox txtCGC 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   18
            TabIndex        =   8
            Top             =   510
            Width           =   1875
         End
         Begin VB.TextBox txtRazaoSocial 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1980
            MaxLength       =   50
            TabIndex        =   9
            Top             =   510
            Width           =   3855
         End
         Begin VB.Label lblCG 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "C.G.C / C.N.P.J"
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
            Left            =   60
            TabIndex        =   65
            Top             =   270
            Width           =   1845
         End
         Begin VB.Label lblRazaoSocial 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Razão Social"
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
            Left            =   1980
            TabIndex        =   64
            Top             =   270
            Width           =   3825
         End
      End
      Begin VB.TextBox txtVendaDiaria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -74790
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   4020
         Width           =   1545
      End
      Begin VB.Frame FraIdentificacaoPF 
         Caption         =   "Identificação - Pessoa Física"
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
         Height          =   1005
         Left            =   -74850
         TabIndex        =   59
         Top             =   2430
         Width           =   2355
         Begin VB.TextBox txtCPF 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   14
            TabIndex        =   7
            Top             =   510
            Width           =   1695
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
            Height          =   195
            Left            =   60
            TabIndex        =   60
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame FraTipoEmpresa 
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
         Left            =   -74850
         TabIndex        =   3
         Top             =   420
         Width           =   4095
         Begin VB.OptionButton OptPessoaJuridica 
            Caption         =   "Pessoa jurídica"
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
            Left            =   2280
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptPessoaFisica 
            Caption         =   "Pessoa física"
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
            Left            =   720
            TabIndex        =   58
            Top             =   270
            Width           =   1485
         End
      End
      Begin VB.ComboBox cboRamo 
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
         Left            =   -70620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   690
         Width           =   3135
      End
      Begin Calendario.ctlPicker CtlDataInicioOperacao 
         Height          =   315
         Left            =   -67410
         TabIndex        =   6
         Top             =   660
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Text            =   "__/__/____"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image ImageDefault 
         BorderStyle     =   1  'Fixed Single
         Height          =   2265
         Left            =   -70170
         Picture         =   "frm_crudEmpresaCliente.frx":366B
         Stretch         =   -1  'True
         Top             =   2460
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblVendaDiaria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Faturamento / Dia"
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
         Left            =   -74790
         TabIndex        =   62
         Top             =   3810
         Width           =   1515
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
         Left            =   -67410
         TabIndex        =   61
         Top             =   420
         Width           =   1755
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
         Left            =   -70620
         TabIndex        =   57
         Top             =   420
         Width           =   1305
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10830
      Top             =   2700
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
            Picture         =   "frm_crudEmpresaCliente.frx":6708
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresaCliente.frx":6CA2
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresaCliente.frx":723C
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresaCliente.frx":77D6
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresaCliente.frx":7D70
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresaCliente.frx":830A
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroEmpresa 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
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
      MouseIcon       =   "frm_crudEmpresaCliente.frx":88A4
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   51
      Top             =   9555
      Width           =   9855
      _ExtentX        =   17383
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
Attribute VB_Name = "frm_crudEmpresaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCreditoFacil       As New ControladorCreditoFacil
Private oEmpresaCliente     As New clsEMPRESACLIENTE
Private oEndEmpCli          As New clsENDERECO
Private oResponsavel        As New clsResponsavel
Private oEndResFin          As New clsENDERECO
Private mstrTipoOperacao    As String
Private mCtLock             As Long
Private mCtLockEndereco     As Long
Private mCtLockEnderecoResponsavel     As Long
Private mCtLockResponsavel     As Long

Private Sub TabStrip1_Click()

End Sub

Private Sub cboBairro_LostFocus()
SSTClienteEmpresa.Tab = 1
End Sub

Private Sub cboCidade__Click()
    PopulaComboBairro_
End Sub

Private Sub cboCidade_Click()
    PopulaComboBairro
End Sub

Private Sub cboEstado__Click()
    PopulaComboCidade_
End Sub

Private Sub cboEstado_Click()
    PopulaComboCidade
End Sub

Private Sub cmdLimpar_Click()
    If MsgBox("Você deseja realmente limpar a imagem ?", vbYesNo, "CONFIRMAÇÃO") = vbYes Then
        imgLogoEmpresaCliente.Picture = ImageDefault.Picture
        txtImagePath = ""
    End If
End Sub

Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    If Screen.ActiveControl.Name = "cmdSelecaoEntidade" Then
    
        Set frmPesquisa.rsResultset = oCreditoFacil.oEmpresaCliente.recuperarEmpresasCliente()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_EMPRESACLIENTE"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "NOME"
        frmPesquisa.Caption = frmPesquisa.Caption & " Empresas Cliente Cadastradas"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtCodigoEntidade_LostFocus
    
    End If

End Sub

Private Sub cmdTrocar_Click()
    On Error GoTo erro
      dlgFOTO.DialogTitle = "Selecione uma imagem no formato indicado"
      dlgFOTO.InitDir = "C:"
      dlgFOTO.FileName = "*.jpg;*.jpeg;*.gif;*.bmp"
      dlgFOTO.Filter = "*.jpg;*.jpeg;*.gif;*.bmp"
      dlgFOTO.ShowOpen
      imgLogoEmpresaCliente.Picture = LoadPicture(dlgFOTO.FileName)
      txtImagePath.Text = dlgFOTO.FileName
        On Error GoTo 0
         Exit Sub
erro:
        If Err.Number = 75 Then
            'MsgBox "Cancelado pelo Usuário"
        End If
End Sub

Private Sub Form_Load()
    PopulaRamos
    PopulaComboSexo
    PopulaEstadoCivil
    PopulaTipoImovel
    PopulaComboEstado
    SSTClienteEmpresa.Tab = 0
    OptPessoaJuridica_Click
    
    oCreditoFacil.oEmpresaCliente.m_timeOut = gstrTimeOutGeral
    oCreditoFacil.oEmpresaCliente.m_stringConexao = gstrConexaoCreditoFacil
    
    oCreditoFacil.oEstado.mTIMEOUT = gstrTimeOutGeral
    oCreditoFacil.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    
    ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = False
    ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False

    
End Sub

Private Sub PopulaRamos()

Dim rs As ADODB.Recordset

oCreditoFacil.oRamo.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oRamo.mSTRING_CONEXAO = gstrConexaoCreditoFacil
Set rs = oCreditoFacil.oRamo.RecuperarRamos()
cboRamo.Clear
Do While Not rs.EOF
  cboRamo.AddItem rs("DESCRICAO")
  cboRamo.ItemData(cboRamo.NewIndex) = rs("ID_RAMO")
  rs.MoveNext
Loop

cboRamo.ListIndex = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set oCreditoFacil = Nothing
    Set oEmpresaCliente = Nothing
    Set oEndEmpCli = Nothing
    Set oResponsavel = Nothing
    Set oEndResFin = Nothing


End Sub

Private Sub OptPessoaFisica_Click()
    FraIdentificacaoPJ.Enabled = False
    FraIdentificacaoPF.Enabled = True
End Sub

Private Sub OptPessoaJuridica_Click()
    FraIdentificacaoPJ.Enabled = True
    FraIdentificacaoPF.Enabled = False
End Sub

Private Sub PopulaComboSexo()
    CboSexo.AddItem "M", 0
    CboSexo.AddItem "F", 1
End Sub
Private Sub PopulaEstadoCivil()
    cboEstadoCivil.AddItem "Solteiro", 0
    cboEstadoCivil.AddItem "Casado", 1
    cboEstadoCivil.AddItem "Viúvo", 2
    cboEstadoCivil.AddItem "Outro", 3
End Sub
Private Sub PopulaTipoImovel()
    cboTipoImovel.AddItem "Alugado", 0
    cboTipoImovel.AddItem "Financiado", 1
    cboTipoImovel.AddItem "Próprio", 2
End Sub
Private Sub PopulaComboEstado()

Dim rs As ADODB.Recordset

oCreditoFacil.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
oCreditoFacil.oEstado.mTIMEOUT = gstrTimeOutGeral
Set rs = oCreditoFacil.oEstado.RecuperaEstados()
cboEstado.Clear
cboEstado_.Clear
Do While Not rs.EOF
  cboEstado.AddItem rs("SIGLA")
  cboEstado_.AddItem rs("SIGLA")
  cboEstado.ItemData(cboEstado.NewIndex) = rs("ID_ESTADO")
  cboEstado_.ItemData(cboEstado_.NewIndex) = rs("ID_ESTADO")
  rs.MoveNext
Loop

cboEstado.ListIndex = -1
cboEstado_.ListIndex = -1

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
Private Sub PopulaComboCidade_()

Dim rs As ADODB.Recordset

If cboEstado_.ListIndex = -1 Then Exit Sub

Set rs = oCreditoFacil.oMunicipio.RecuperarMunicipios(gstrConexaoCreditoFacil, gstrTimeOutGeral, cboEstado.ItemData((cboEstado_.ListIndex)))
cboCidade_.Clear
Do While Not rs.EOF
  cboCidade_.AddItem rs("DESCRICAO")
  cboCidade_.ItemData(cboCidade_.NewIndex) = rs("ID_MUNICIPIO")
  rs.MoveNext
Loop

cboCidade_.ListIndex = -1

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
Private Sub PopulaComboBairro_()

Dim rs As ADODB.Recordset

If cboEstado_.ListIndex = -1 Or cboCidade_.ListIndex = -1 Then Exit Sub

Set rs = oCreditoFacil.oBairro.RecuperarBairros(gstrConexaoCreditoFacil, gstrTimeOutGeral, cboEstado.ItemData((cboEstado_.ListIndex)), cboCidade_.ItemData((cboCidade_.ListIndex)))
cboBairro_.Clear
Do While Not rs.EOF
  cboBairro_.AddItem rs("DESCRICAO_BAIRRO")
  cboBairro_.ItemData(cboBairro_.NewIndex) = rs("ID_BAIRRO")
  rs.MoveNext
Loop

cboBairro_.ListIndex = -1

End Sub

Private Sub ToolbarCadastroEmpresa_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
      Case 1 'Novo
        Novo_Click
        'PreencheCamposObrigatoriosTeste
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

SSTClienteEmpresa.Tab = 0
If Not ValidarClienteEmpresa Then
    Exit Sub
End If
If Not ValidarEnderecoEmpresaCliente Then
    MsgBox "O endereço da empresa cliente não está preenchido completamente. Complete o cadastro e tente novamente.", vbInformation, "MENSAGEM"
    Exit Sub
End If
SSTClienteEmpresa.Tab = 1
If Not ValidarResponsavelFinanceiro Then
    Exit Sub
End If
If Not ValidarEnderecoResponsavelFinanceiro Then
    MsgBox "O endereço do responável financeiro não está preenchido completamente. Complete o cadastro e tente novamente.", vbInformation, "MENSAGEM"
    Exit Sub
End If
SSTClienteEmpresa.Tab = 0

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
ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False

End Sub
'Validar Campos de ClienteEmpresa
Private Function ValidarClienteEmpresa() As Boolean

  If Trim(Len(txtDescricaoEntidade)) = 0 Then
     If OptPessoaFisica Then
        MsgBox "A nome do propietário da empresa é requerido.", vbInformation, "Mensagem"
     Else
        MsgBox "A nome fantasia da empresa é requerido.", vbInformation, "Mensagem"
     End If
     txtDescricaoEntidade.SetFocus
     ValidarClienteEmpresa = False
     Exit Function
  End If
  
  If cboRamo.ListIndex = -1 Then
    MsgBox "O ramo de atividade da empresa é requerido.", vbInformation, "Mensagem"
    cboRamo.SetFocus
    ValidarClienteEmpresa = False
    Exit Function
  End If
  
  If CtlDataInicioOperacao.Text = "__/__/____" Then
    MsgBox "O campo 'em atividade desde' da empresa é requerido.", vbInformation, "Mensagem"
    CtlDataInicioOperacao.SetFocus
    ValidarClienteEmpresa = False
    Exit Function
  End If
  
  Select Case OptPessoaFisica
  
    Case True
        If Len(txtCPF) = 0 Then
            MsgBox "O CPF do propietário da empresa é requerido.", vbInformation, "Mensagem"
            txtCPF.SetFocus
            ValidarClienteEmpresa = False
            Exit Function
        End If
    Case Else
        If Len(txtCGC) = 0 Then
            MsgBox "O CGC/CNPJ da empresa é requerido.", vbInformation, "Mensagem"
            txtCGC.SetFocus
            ValidarClienteEmpresa = False
            Exit Function
        End If
        If Len(txtRazaoSocial) = 0 Then
            MsgBox "A razão social da empresa é requerida.", vbInformation, "Mensagem"
            txtRazaoSocial.SetFocus
            ValidarClienteEmpresa = False
            Exit Function
        End If
  End Select
  
  If Not IsNumeric(txtVendaDiaria) Then
    MsgBox "O valor da venda diária está incorreto ou não foi preenchido, favor escreva no seguinte formato #.##", vbInformation, "INFORMACÃO"
    txtVendaDiaria.SetFocus
    txtVendaDiaria.SelStart = 0
    txtVendaDiaria.SelLength = Len(txtVendaDiaria)
    ValidarClienteEmpresa = False
    Exit Function
  End If
  
  If CCur(txtVendaDiaria) <= 0 Then
    MsgBox "O valor da venda diária não pode ser zero ou negativo, favor escreva no seguinte formato #.##", vbInformation, "INFORMACÃO"
    txtVendaDiaria.SetFocus
    ValidarClienteEmpresa = False
    Exit Function
  End If
  
  ValidarClienteEmpresa = True
  
End Function
'Validar Campos de responsavelFinanceiro
Private Function ValidarResponsavelFinanceiro() As Boolean

  If Len(txtNomeResposavelFinanceiro) = 0 Then
    MsgBox "O nome do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    txtNomeResposavelFinanceiro.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If CtlDataNascimento.Text = "__/__/____" Then
    MsgBox "A data de nascimento do responsável financeiro da empresa é requerida.", vbInformation, "Mensagem"
    CtlDataNascimento.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If Len(txtCPF_Responsavel) = 0 Then
    MsgBox "O CPF do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    txtCPF_Responsavel.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If Len(txtRG_Responsavel) = 0 Then
    MsgBox "O RG do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    txtRG_Responsavel.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If Len(txtOrgaoEmissor) = 0 Then
    MsgBox "O orgão emissor do RG do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    txtOrgaoEmissor.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If CtlDataExpedicao.Text = "__/__/____" Then
    MsgBox "A data de expedição do RG do responsável financeiro da empresa é requerida.", vbInformation, "Mensagem"
    CtlDataNascimento.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If CboSexo.ListIndex = -1 Then
    MsgBox "O sexo do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    CboSexo.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If cboEstadoCivil.ListIndex = -1 Then
    MsgBox "O estado civíl do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    cboEstadoCivil.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If Len(txtNaturalidade) = 0 Then
    MsgBox "A naturalidade do responsável financeiro da empresa é requerida.", vbInformation, "Mensagem"
    txtNaturalidade.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If Len(txtNacionalidade) = 0 Then
    MsgBox "A nacionalidade do responsável financeiro da empresa é requerida.", vbInformation, "Mensagem"
    txtNacionalidade.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If Len(txtFiliacaoMae) = 0 Then
    MsgBox "O nome da mãe do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    txtFiliacaoMae.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If cboTipoImovel.ListIndex = -1 Then
    MsgBox "O tipo de imóvel do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    cboTipoImovel.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  If CtlResideDesde.Text = "__/__/____" Then
    MsgBox "O campo 'reside desde' para o imóvel do responsável financeiro da empresa é requerido.", vbInformation, "Mensagem"
    CtlResideDesde.SetFocus
    ValidarResponsavelFinanceiro = False
    Exit Function
  End If
  
  ValidarResponsavelFinanceiro = True
  
End Function
'Validar Campos do Formulario
Private Function ValidarEnderecoEmpresaCliente() As Boolean

Dim rsFuncao As ADODB.Recordset
    
  If Trim(Len(txtTipoLogradouro)) = 0 Then
     ValidarEnderecoEmpresaCliente = False
     Exit Function
  End If
  
  If Trim(Len(txtLogradouro)) = 0 Then
     ValidarEnderecoEmpresaCliente = False
     Exit Function
  End If
  
  If Trim(Len(txtNumero)) = 0 Then
     ValidarEnderecoEmpresaCliente = False
     Exit Function
  End If
  
  If Trim(Len(txtTipoLogradouro)) = 0 Then
     ValidarEnderecoEmpresaCliente = False
     Exit Function
  End If
  
  If cboEstado.ListIndex = -1 Or cboCidade.ListIndex = -1 Or cboBairro.ListIndex = -1 Then
     ValidarEnderecoEmpresaCliente = False
     Exit Function
  End If
  
  ValidarEnderecoEmpresaCliente = True
  
End Function
'Validar Campos do Formulario
Private Function ValidarEnderecoResponsavelFinanceiro() As Boolean

Dim rsFuncao As ADODB.Recordset
    
  If Trim(Len(txtTipoLogradouro_)) = 0 Then
     ValidarEnderecoResponsavelFinanceiro = False
     Exit Function
  End If
  
  If Trim(Len(txtLogradouro_)) = 0 Then
     ValidarEnderecoResponsavelFinanceiro = False
     Exit Function
  End If
  
  If Trim(Len(txtNumero_)) = 0 Then
     ValidarEnderecoResponsavelFinanceiro = False
     Exit Function
  End If
  
  If Trim(Len(txtTipoLogradouro_)) = 0 Then
     ValidarEnderecoResponsavelFinanceiro = False
     Exit Function
  End If
  
  If cboEstado_.ListIndex = -1 Or cboCidade_.ListIndex = -1 Or cboBairro_.ListIndex = -1 Then
     ValidarEnderecoResponsavelFinanceiro = False
     Exit Function
  End If
  
  ValidarEnderecoResponsavelFinanceiro = True
  
End Function
Private Sub Novo_Click()

txtCodigoEntidade.Text = oCreditoFacil.oEmpresaCliente.getNovoIdEmpresaCliente
txtDescricaoEntidade.SetFocus
ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = True
mstrTipoOperacao = "I"
stbmsg.SimpleText = "Incluindo"
DoEvents

End Sub
Private Sub Excluir_Click()

If txtCodigoEntidade = "" Or txtCodigoEntidade = "0" Then Exit Sub

If MsgBox("Deseja realmente excluir o registro?", vbYesNo, "Pergunta") = vbNo Then Exit Sub

mstrTipoOperacao = "E"

MoveTelaParaObjetoCab mstrTipoOperacao

If txtID > 0 Then MsgBox "Registro excluido com sucesso!", vbInformation, "SUCESSO"

Call Limpacampos
txtCodigoEntidade = ""


End Sub
Private Sub Fechar_Click()
  Unload Me
End Sub
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
    On Error GoTo trataerro
    Dim rsEndEmpCli As New ADODB.Recordset
    Dim rsEndResFin As New ADODB.Recordset
    Dim rsResFin As New ADODB.Recordset
    
    
    
    'Atributos da empresaCliente
    With oEmpresaCliente
    
        .m_01_idEmpresaCliente = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
        .m_02_tipo = IIf(OptPessoaFisica, "F", "J")
        .m_03_idRamo = cboRamo.ItemData(cboRamo.ListIndex)
        .m_04_iniciouAtividade = CtlDataInicioOperacao.Text
        .m_05_vendaDiaria = txtVendaDiaria
        .m_06_idEndereco = IIf(txtIDEndereco = "", 0, txtIDEndereco)
        .m_07_cgc = txtCGC
        .m_08_razaoSocial = txtRazaoSocial
        .m_09_nomeFantasia = IIf(OptPessoaJuridica, txtDescricaoEntidade, "")
        .m_10_cpf = txtCPF
        .m_11_nomePessoaFisica = IIf(OptPessoaFisica, txtDescricaoEntidade, "")
        .m_12_blobLogoEmpresa = txtImagePath
        .m_13_dataInclusao = Now()
        .m_14_usuarioInclusao = LogInUserID
        .m_15_dataAlteracao = Now()
        .m_16_usuarioAlteracao = LogInUserID
        .m_17_ctLock = mCtLock
            
    End With
        
    'Atributos do endereco da empresaCliente
    oEndEmpCli.m_timeOut = gstrTimeOutGeral
    oEndEmpCli.m_stringConexao = gstrConexaoCreditoFacil
    
    oEndEmpCli.inicializaEndereco
    With oEndEmpCli.rsEndereco
        .Open
        .AddNew
        .Fields("ID_OBJECT_ENTIDADE") = oEndEmpCli.consultaIdObjectEntidade("empresaCliente")
        .Fields("ID_ENTIDADE") = CLng(txtCodigoEntidade)
        .Fields("ID_ENDERECO") = CLng(IIf(Len(txtIDEndereco) = 0, 0, txtIDEndereco))
        .Fields("TIPO_LOGRADOURO") = txtTipoLogradouro
        .Fields("LOGRADOURO") = txtLogradouro
        .Fields("NUMERO") = txtNumero
        .Fields("COMPLEMENTO") = txtComplemento
        .Fields("PONTO_REFERENCIA") = txtPontoReferencia
        .Fields("CEP") = txtCEP
        If cboBairro.ListIndex <> -1 Then
          .Fields("ID_BAIRRO") = cboBairro.ItemData(cboBairro.ListIndex)
        End If
        If cboCidade.ListIndex <> -1 Then
          .Fields("ID_MUNICIPIO") = cboCidade.ItemData(cboCidade.ListIndex)
        End If
        If cboEstado.ListIndex <> -1 Then
          .Fields("ID_ESTADO") = cboEstado.ItemData(cboEstado.ListIndex)
        End If
        If txtIDEndereco.Text = "" Then
          .Fields("USUARIO_INCLUSAO") = LogInUserID
          .Fields("DATA_INCLUSAO") = ""
        End If
        .Fields("USUARIO_ALTERACAO") = LogInUserID
        .Fields("DATA_ALTERACAO") = ""
        .Fields("CT_LOCK") = mCtLockEndereco
        .Update
    End With
   'Guarda o recordset que contém o endereço da empresaCliente
    Set rsEndEmpCli = oEndEmpCli.rsEndereco
    
    'Atributos do endereco da empresaCliente
    oResponsavel.m_timeOut = gstrTimeOutGeral
    oResponsavel.m_stringConexao = gstrConexaoCreditoFacil
    
    'Atributos do responsavelFinanceiro
    oResponsavel.inicializaResFin
    With oResponsavel.rsResFin
        .Open
        .AddNew
        .Fields("ID_RESPONSAVEL") = IIf(Trim(Len(txtIDResponsavelFinanceiro.Text)) = "", 0, txtIDResponsavelFinanceiro.Text)
        .Fields("ID_EMPRESACLIENTE") = txtCodigoEntidade
        .Fields("SITUACAO") = IIf(chkDesativado.value = 1, "D", "A")
        .Fields("NOME") = txtNomeResposavelFinanceiro
        .Fields("CPF") = txtCPF_Responsavel
        .Fields("RG") = txtRG_Responsavel
        .Fields("ORGAO_EMISSOR") = txtOrgaoEmissor
        .Fields("DATA_EXPEDICAO") = CtlDataExpedicao.Text
        .Fields("SEXO") = CboSexo.Text
        .Fields("DATA_NASCIMENTO") = CtlDataNascimento.Text
        .Fields("ESTADO_CIVIL") = Mid$(cboEstadoCivil.Text, 1, 1)
        .Fields("NATURALIDADE") = txtNaturalidade
        .Fields("NACIONALIDADE") = txtNacionalidade
        .Fields("NOME_MAE") = txtFiliacaoMae
        .Fields("ID_ENDERECO") = IIf(txtIDEnderecoResponsavel = "", 0, txtIDEnderecoResponsavel)
        .Fields("TIPO_IMOVEL") = Mid$(cboTipoImovel.Text, 1, 1)
        .Fields("RESIDE_DESDE") = CtlResideDesde.Text
        .Fields("TELEFONE1") = txtTelefone1
        .Fields("TELEFONE2") = txtTelefone2
        .Fields("TELEFONE3") = txtTelefone3
        .Fields("CONTATO_REFERENCIA1") = txtNomeContato1
        .Fields("TELEFONE_REFERENCIA1") = txtFoneContato1
        .Fields("CONTATO_REFERENCIA2") = txtNomeContato2
        .Fields("TELEFONE_REFERENCIA2") = txtFoneContato2
        .Fields("INDICADO_POR") = txtIndicadoPor
        .Fields("DATA_INCLUSAO") = ""
        .Fields("USUARIO_INCLUSAO") = LogInUserID
        .Fields("DATA_ALTERACAO") = ""
        .Fields("USUARIO_ALTERACAO") = LogInUserID
        .Fields("CT_LOCK") = mCtLockResponsavel
    End With
   'Guarda o recordset que contém o endereço do responsavelFinanceiro da empresaCliente
    Set rsResFin = oResponsavel.rsResFin
    
    
    
    'Atributos do endereco do responsavelFinanceiro
    oEndResFin.m_timeOut = gstrTimeOutGeral
    oEndResFin.m_stringConexao = gstrConexaoCreditoFacil
    oEndResFin.inicializaEndereco
    With oEndResFin.rsEndereco
        .Open
        .AddNew
        .Fields("ID_OBJECT_ENTIDADE") = oEndResFin.consultaIdObjectEntidade("responsavelFinanceiro")
        .Fields("ID_ENTIDADE") = txtIDResponsavelFinanceiro
        .Fields("ID_ENDERECO") = IIf(txtIDEnderecoResponsavel = "", 0, txtIDEnderecoResponsavel)
        .Fields("TIPO_LOGRADOURO") = txtTipoLogradouro_
        .Fields("LOGRADOURO") = txtLogradouro_
        .Fields("NUMERO") = txtNumero_
        .Fields("COMPLEMENTO") = txtComplemento_
        .Fields("PONTO_REFERENCIA") = txtPontoReferencia_
        .Fields("CEP") = txtCEP_
        If cboBairro_.ListIndex <> -1 Then
          .Fields("ID_BAIRRO") = cboBairro_.ItemData(cboBairro_.ListIndex)
        End If
        If cboCidade_.ListIndex <> -1 Then
          .Fields("ID_MUNICIPIO") = cboCidade_.ItemData(cboCidade_.ListIndex)
        End If
        If cboEstado_.ListIndex <> -1 Then
          .Fields("ID_ESTADO") = cboEstado_.ItemData(cboEstado_.ListIndex)
        End If
        If txtIDEnderecoResponsavel.Text = "" Then
          .Fields("USUARIO_INCLUSAO") = LogInUserID
          .Fields("DATA_INCLUSAO") = ""
        End If
        .Fields("USUARIO_ALTERACAO") = LogInUserID
        .Fields("DATA_ALTERACAO") = ""
        .Fields("CT_LOCK") = mCtLockEnderecoResponsavel
        .Update
    End With
   'Guarda o recordset que contém o endereço da empresaCliente
    Set rsEndResFin = oEndResFin.rsEndereco
    
    
    With oEmpresaCliente
        .m_timeOut = gstrTimeOutGeral
        .m_stringConexao = gstrConexaoCreditoFacil
    
        If strOperacao = "I" Then
            txtID.Text = .crudInsert(rsEndEmpCli, rsResFin, rsEndResFin)
            txtID.Text = .crudUpdate(rsEndEmpCli, rsResFin, rsEndResFin)
        ElseIf strOperacao = "A" Then
            txtID.Text = .crudUpdate(rsEndEmpCli, rsResFin, rsEndResFin)
        Else
            txtIDEndereco.Text = .crudDelete(rsEndEmpCli, rsResFin, rsEndResFin)
        End If
    
    End With
    
trataerro:
 If InStr(1, Err.Description, "FK_linhaCredito_empresaCliente") > 0 And Err.Number = -2147221503 Then
    MsgBox "Não é possível excluir a empresa, pois existe linha de crédito cadastrada para a mesma. Se desejar, exclua a linha de crédito e tente novamente.", vbInformation, "EXCLUSÃO NÃO FOI REALIZADA"
    txtID = 0
 End If
If InStr(1, Err.Description, "FK_endereco_bairro") > 0 And Err.Number = -2147221503 Then
    MsgBox "Não é possível excluir este bairro, pois está sendo usado em algum endereço.", vbInformation, "NÃO FOI POSSÍVEL EXCLUIR"
    txtID = 0
End If

End Sub
Private Sub Limpacampos()
        
    'Campos da empresaCliente
    txtDescricaoEntidade.Text = ""
    OptPessoaJuridica.value = 1
    cboRamo.ListIndex = -1
    CtlDataInicioOperacao.Text = "__/__/____"
    txtCPF = ""
    txtCGC = ""
    txtRazaoSocial = ""
    txtVendaDiaria = "0,00"
    imgLogoEmpresaCliente.Picture = ImageDefault.Picture
    txtImagePath = ""
    
    LimpaCamposEnderecoEmpresaCliente
    
    'Campos do responsavel Financeiro
    txtNomeResposavelFinanceiro = ""
    CtlDataNascimento.Text = "__/__/____"
    chkDesativado.value = 0
    txtCPF_Responsavel = ""
    txtRG_Responsavel = ""
    txtOrgaoEmissor = ""
    CtlDataExpedicao.Text = "__/__/____"
    CboSexo.ListIndex = -1
    cboEstadoCivil.ListIndex = -1
    txtNaturalidade = ""
    txtNacionalidade = ""
    txtFiliacaoMae = ""
    txtTelefone1 = ""
    txtTelefone2 = ""
    txtTelefone3 = ""
    cboTipoImovel.ListIndex = -1
    CtlResideDesde.Text = "__/__/____"
    txtNomeContato1 = ""
    txtFoneContato1 = ""
    txtNomeContato2 = ""
    txtFoneContato2 = ""
    txtIndicadoPor = ""
    
    LimpaCamposEnderecoResponsavel

End Sub

Private Sub LimpaCamposEnderecoEmpresaCliente()

txtTipoLogradouro = ""
txtLogradouro = ""
txtNumero = ""
txtComplemento = ""
txtPontoReferencia = ""
txtCEP = ""
cboBairro.ListIndex = -1
cboCidade.ListIndex = -1
cboEstado.ListIndex = -1

End Sub
Private Sub LimpaCamposEnderecoResponsavel()

txtTipoLogradouro_ = ""
txtLogradouro_ = ""
txtNumero_ = ""
txtComplemento_ = ""
txtPontoReferencia_ = ""
txtCEP_ = ""
cboBairro_.ListIndex = -1
cboCidade_.ListIndex = -1
cboEstado_.ListIndex = -1

End Sub

Private Sub txtCodigoEntidade_Change()

    Limpacampos
    ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = False
    ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False
    txtID.Text = ""
    txtIDEndereco = ""
    txtIDResponsavelFinanceiro = ""
    txtIDEnderecoResponsavel = ""
    stbmsg.SimpleText = ""
    
End Sub

Private Sub txtCodigoEntidade_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCodigoEntidade_LostFocus()

Dim rsEmpCli As ADODB.Recordset
Dim rsEndEmpCli As ADODB.Recordset
Dim rsResFin As ADODB.Recordset
Dim rsEndResFin As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
        
    'Recupera a empresaCliente
    Set rsEmpCli = oCreditoFacil.oEmpresaCliente.consulta(txtCodigoEntidade)
    
    'Caso não exista a empresaCliente, sair
    If rsEmpCli.EOF Then
      If oCreditoFacil.oEmpresaCliente.getNovoIdEmpresaCliente <> txtCodigoEntidade Then
        ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False
      End If
      Exit Sub
    End If
    
    'Recupera o endereco da empresaCliente
    With oCreditoFacil.oEndereco
        If Not rsEmpCli.EOF Then
            .m_timeOut = gstrTimeOutGeral
            .m_stringConexao = gstrConexaoCreditoFacil
            Set rsEndEmpCli = .recuperarEndereco(.consultaIdObjectEntidade("empresaCliente"), rsEmpCli("ID_EMPRESACLIENTE"))
        End If
    End With
    
    'Recupera o responsavelFinanceiro
    With oResponsavel
        If Not rsEmpCli.EOF Then
            .m_timeOut = gstrTimeOutGeral
            .m_stringConexao = gstrConexaoCreditoFacil
            Set rsResFin = .recuperarResponsavelFicanceiro(rsEmpCli("ID_EMPRESACLIENTE"))
        End If
    End With
       
    'Recupera o endereco do responsavelFinanceiro
    With oEndResFin
        If Not rsResFin.EOF Then
            .m_timeOut = gstrTimeOutGeral
            .m_stringConexao = gstrConexaoCreditoFacil
            Set rsEndResFin = .recuperarEndereco(.consultaIdObjectEntidade("responsavelFinanceiro"), rsResFin("ID_RESPONSAVEL"))
        End If
    End With
    
    ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = True
    ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rsEmpCli, rsEndEmpCli, rsResFin, rsEndResFin
    txtDescricaoEntidade.SetFocus
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rsEmpCli As ADODB.Recordset, _
                                  ByRef rsEndEmpCli As ADODB.Recordset, _
                                  ByRef rsResFin As ADODB.Recordset, _
                                  ByRef rsEndResFin As ADODB.Recordset)
    
    Dim i As Integer
    
    'Descarregando dados da empresa
    txtID.Text = rsEmpCli("ID_EMPRESACLIENTE")
    mCtLock = rsEmpCli("CT_LOCK")
    
    Select Case rsEmpCli("TIPO")
    
        Case "F" 'Pessoa Física
            OptPessoaFisica.value = True
            txtCPF = IIf(IsNull(rsEmpCli("CPF")), "", rsEmpCli("CPF"))
            txtDescricaoEntidade = rsEmpCli("NOME_PESSOA_FISICA")
        Case "J" 'Pessoa Jurídica
            OptPessoaJuridica.value = True
            txtCGC = IIf(IsNull(rsEmpCli("CGC")), "", rsEmpCli("CGC"))
            txtRazaoSocial = IIf(IsNull(rsEmpCli("RAZAO_SOCIAL")), "", rsEmpCli("RAZAO_SOCIAL"))
            txtDescricaoEntidade = rsEmpCli("NOME_FANTASIA")
            
    End Select
        
    For i = 1 To cboRamo.ListCount
        If cboRamo.ItemData(i - 1) = rsEmpCli("ID_RAMO") Then
            cboRamo.ListIndex = i - 1
        End If
    Next
    CtlDataInicioOperacao.Text = Format(rsEmpCli("INICIOU_ATIVIDADE"), "dd/mm/yyyy")
    txtVendaDiaria = Format(rsEmpCli("VENDA_DIARIA"), "0.00")
    
    'Descarregando a LOGO
    If Not IsNull(rsEmpCli("BLOB_LOGO_EMPRESA")) Then
        txtImagePath = oCreditoFacil.oEmpresaCliente.carregarImagem(rsEmpCli)
        imgLogoEmpresaCliente.Picture = LoadPicture(txtImagePath)
    End If
        
    
    If Not rsEndEmpCli.EOF Then
    
     'Descarregando Endereço da empresa
      txtIDEndereco.Text = rsEndEmpCli("ID_ENDERECO")
      mCtLockEndereco = rsEndEmpCli("CT_LOCK")
      txtTipoLogradouro = rsEndEmpCli("TIPO_LOGRADOURO")
      txtLogradouro = rsEndEmpCli("LOGRADOURO")
      txtNumero = rsEndEmpCli("NUMERO")
      txtComplemento = IIf(IsNull(rsEndEmpCli("COMPLEMENTO")), "", rsEndEmpCli("COMPLEMENTO"))
      txtPontoReferencia = IIf(IsNull(rsEndEmpCli("PONTO_REFERENCIA")), "", rsEndEmpCli("PONTO_REFERENCIA"))
      txtCEP = IIf(IsNull(rsEndEmpCli("CEP")), "", rsEndEmpCli("CEP"))
      For i = 1 To cboEstado.ListCount
        If cboEstado.ItemData(i - 1) = rsEndEmpCli("ID_ESTADO") Then
          cboEstado.ListIndex = i - 1
        End If
      Next
      For i = 1 To cboCidade.ListCount
        If cboCidade.ItemData(i - 1) = rsEndEmpCli("ID_MUNICIPIO") Then
          cboCidade.ListIndex = i - 1
        End If
      Next
      For i = 1 To cboBairro.ListCount
        If cboBairro.ItemData(i - 1) = rsEndEmpCli("ID_BAIRRO") Then
          cboBairro.ListIndex = i - 1
        End If
      Next
    End If
    
    'Descarregando dados do responsavel
    If Not rsResFin.EOF Then
      
      txtIDResponsavelFinanceiro.Text = rsResFin("ID_RESPONSAVEL")
      mCtLockResponsavel = rsResFin("CT_LOCK")
      
      txtNomeResposavelFinanceiro = rsResFin("NOME")
      CtlDataNascimento.Text = Format(rsResFin("DATA_NASCIMENTO"), "dd/mm/yyyy")
      chkDesativado.value = IIf(rsResFin("SITUACAO") = "A", 0, 1)
      txtCPF_Responsavel = rsResFin("CPF")
      txtRG_Responsavel = rsResFin("RG")
      txtOrgaoEmissor = rsResFin("ORGAO_EMISSOR")
      CtlDataExpedicao.Text = Format(rsResFin("DATA_EXPEDICAO"), "dd/mm/yyyy")
      CboSexo.ListIndex = IIf(rsResFin("SEXO") = "M", 0, 1)
      Select Case rsResFin("ESTADO_CIVIL")
      
        Case "S" 'Solteiro
            cboEstadoCivil.ListIndex = 0
        Case "C" 'Casado
            cboEstadoCivil.ListIndex = 1
        Case "V" 'Viuvo
            cboEstadoCivil.ListIndex = 2
        Case "O" 'Outro
            cboEstadoCivil.ListIndex = 3
      
      End Select
      txtNaturalidade = rsResFin("NATURALIDADE")
      txtNacionalidade = rsResFin("NACIONALIDADE")
      txtFiliacaoMae = rsResFin("NOME_MAE")
      txtTelefone1 = IIf(IsNull(rsResFin("TELEFONE1")), "", rsResFin("TELEFONE1"))
      txtTelefone2 = IIf(IsNull(rsResFin("TELEFONE2")), "", rsResFin("TELEFONE2"))
      txtTelefone3 = IIf(IsNull(rsResFin("TELEFONE3")), "", rsResFin("TELEFONE3"))
      Select Case rsResFin("TIPO_IMOVEL")
      
        Case "A" 'Alugado
            cboTipoImovel.ListIndex = 0
        Case "F" 'Financiado
            cboTipoImovel.ListIndex = 1
        Case "P" 'Próprio
            cboTipoImovel.ListIndex = 2
      
      End Select
      CtlResideDesde.Text = Format(rsResFin("RESIDE_DESDE"), "dd/mm/yyyy")
      txtNomeContato1 = IIf(IsNull(rsResFin("CONTATO_REFERENCIA1")), "", rsResFin("CONTATO_REFERENCIA1"))
      txtFoneContato1 = IIf(IsNull(rsResFin("TELEFONE_REFERENCIA1")), "", rsResFin("TELEFONE_REFERENCIA1"))
      txtNomeContato2 = IIf(IsNull(rsResFin("CONTATO_REFERENCIA2")), "", rsResFin("CONTATO_REFERENCIA2"))
      txtFoneContato2 = IIf(IsNull(rsResFin("TELEFONE_REFERENCIA2")), "", rsResFin("TELEFONE_REFERENCIA2"))
      txtIndicadoPor = IIf(IsNull(rsResFin("INDICADO_POR")), "", rsResFin("INDICADO_POR"))
      
    End If
    
    'Descarregando Endereço do responsavel
    If Not rsEndResFin.EOF Then
     
      txtIDEnderecoResponsavel.Text = rsEndResFin("ID_ENDERECO")
      mCtLockEnderecoResponsavel = rsEndResFin("CT_LOCK")
      
      txtTipoLogradouro_ = rsEndResFin("TIPO_LOGRADOURO")
      txtLogradouro_ = rsEndResFin("LOGRADOURO")
      txtNumero_ = rsEndResFin("NUMERO")
      txtComplemento_ = IIf(IsNull(rsEndResFin("COMPLEMENTO")), "", rsEndResFin("COMPLEMENTO"))
      txtPontoReferencia_ = IIf(IsNull(rsEndResFin("PONTO_REFERENCIA")), "", rsEndResFin("PONTO_REFERENCIA"))
      txtCEP_ = IIf(IsNull(rsEndResFin("CEP")), "", rsEndResFin("CEP"))
      For i = 1 To cboEstado_.ListCount
        If cboEstado_.ItemData(i - 1) = rsEndResFin("ID_ESTADO") Then
          cboEstado_.ListIndex = i - 1
        End If
      Next
      For i = 1 To cboCidade_.ListCount
        If cboCidade_.ItemData(i - 1) = rsEndResFin("ID_MUNICIPIO") Then
          cboCidade_.ListIndex = i - 1
        End If
      Next
      For i = 1 To cboBairro_.ListCount
        If cboBairro_.ItemData(i - 1) = rsEndResFin("ID_BAIRRO") Then
          cboBairro_.ListIndex = i - 1
        End If
      Next
    End If
    
End Sub

Private Sub PreencheCamposObrigatoriosTeste()

    'DADOS EMPRESA
    txtDescricaoEntidade = "Empreendimentos Claudio S/A"
    OptPessoaFisica.value = True
    OptPessoaFisica_Click
    txtCPF = "617.684.723-00"
    'OptPessoaJuridica.value = True
    'OptPessoaJuridica_Click
    'txtCGC = "1.585.854/541-00"
    'txtRazaoSocial = "CLAUDIO DE SA INOVACOES M.E"
    cboRamo.ListIndex = 2
    CtlDataInicioOperacao.Text = "15/05/1995"
    txtVendaDiaria = "575,96"
        'ENDERECO
        txtTipoLogradouro = "RUA"
        txtLogradouro = "ARARA AZUL"
        txtNumero = "1521"
        txtComplemento = "AP 521"
        txtPontoReferencia = "DO LADO DA BANQUINHA"
        txtCEP = "60841-560"
        cboEstado.ListIndex = 0
        cboEstado_Click
        cboCidade.ListIndex = 1
        cboCidade_Click
        cboBairro.ListIndex = 0
    'DADOS DO RESPONSAVEL
    txtNomeResposavelFinanceiro = "FRANCISCO GUSTAVO"
    CtlDataNascimento.Text = "14/09/1978"
    txtCPF_Responsavel = "852.963.741-54"
    txtRG_Responsavel = "98745001"
    txtOrgaoEmissor = "SSP-CE"
    CtlDataExpedicao.Text = "15/04/2001"
    CboSexo.ListIndex = 0
    cboEstadoCivil.ListIndex = 1
    txtNaturalidade = "Fortaleza"
    txtNacionalidade = "Brasileiro"
    txtFiliacaoMae = "Zuila Paiva de Sa"
    txtTelefone1 = "8803-7269"
    txtTelefone2 = "8803-7270"
    txtTelefone3 = "8803-7271"
    cboTipoImovel.ListIndex = 2
    CtlResideDesde.Text = "01/03/1994"
    txtNomeContato1 = "Maria Ivone"
    txtFoneContato1 = "3276-5947"
    txtNomeContato2 = "Jose Messias"
    txtFoneContato2 = "3275-5948"
    txtIndicadoPor = "Claudio de Sa"
        'ENDERECO
        txtTipoLogradouro_ = "AV"
        txtLogradouro_ = "GODOFREDO MACIEL"
        txtNumero_ = "5200"
        txtComplemento_ = "CASA AMARELA"
        txtPontoReferencia_ = "EM FRENTE A PADARIA"
        txtCEP_ = "60200-750"
        cboEstado_.ListIndex = 0
        cboEstado__Click
        cboCidade_.ListIndex = 1
        cboCidade__Click
        cboBairro_.ListIndex = 2
    
End Sub

Private Sub txtVendaDiaria_Change()
    txtVendaDiaria = Replace(txtVendaDiaria, ".", ",")
    txtVendaDiaria.SelStart = Len(txtVendaDiaria)
End Sub

