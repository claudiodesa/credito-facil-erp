VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_crudEmpresa 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Empresa Mestre"
   ClientHeight    =   7215
   ClientLeft      =   1200
   ClientTop       =   3540
   ClientWidth     =   10080
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_crudEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraLogomarca 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6930
      TabIndex        =   40
      Top             =   2010
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
         TabIndex        =   43
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
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txtImagePath 
         Height          =   315
         Left            =   30
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSComDlg.CommonDialog dlgFOTO 
         Left            =   300
         Top             =   420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgLogoEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   2265
         Left            =   1080
         Picture         =   "frm_crudEmpresa.frx":058A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FraEndereco 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Endereço"
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
      Left            =   240
      TabIndex        =   26
      Top             =   4710
      Width           =   9555
      Begin VB.TextBox txtIDEndereco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   8940
         TabIndex        =   39
         Top             =   120
         Width           =   555
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
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1710
         Width           =   3525
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
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1710
         Width           =   855
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
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1710
         Width           =   3525
      End
      Begin VB.TextBox txtCEP 
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
         Left            =   210
         MaxLength       =   9
         TabIndex        =   16
         Top             =   1710
         Width           =   1065
      End
      Begin VB.TextBox txtPontoReferencia 
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
         Left            =   3090
         MaxLength       =   25
         TabIndex        =   15
         Top             =   1110
         Width           =   6165
      End
      Begin VB.TextBox txtComplemento 
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
         Left            =   210
         MaxLength       =   25
         TabIndex        =   14
         Top             =   1110
         Width           =   2745
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   7860
         MaxLength       =   12
         TabIndex        =   13
         Top             =   510
         Width           =   1395
      End
      Begin VB.TextBox txtLogradouro 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   12
         Top             =   510
         Width           =   6585
      End
      Begin VB.TextBox txtTipoLogradouro 
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
         Left            =   210
         MaxLength       =   10
         TabIndex        =   11
         Top             =   510
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
         Left            =   2190
         TabIndex        =   35
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
         Left            =   1290
         TabIndex        =   34
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
         Height          =   195
         Left            =   5760
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   1470
         Width           =   945
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   870
         Width           =   1575
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
         TabIndex        =   29
         Top             =   270
         Width           =   945
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
         TabIndex        =   28
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblTipoDeLogradouro 
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
         TabIndex        =   27
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame FraCampos 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   240
      TabIndex        =   23
      Top             =   1920
      Width           =   5475
      Begin VB.TextBox txtTelefone3 
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
         Left            =   3270
         MaxLength       =   13
         TabIndex        =   10
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txtTelefone2 
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
         Left            =   1770
         MaxLength       =   13
         TabIndex        =   9
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txtTelefone1 
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
         Left            =   210
         MaxLength       =   13
         TabIndex        =   8
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txtRazaoSocial 
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
         Left            =   210
         MaxLength       =   50
         TabIndex        =   7
         Top             =   960
         Width           =   5115
      End
      Begin VB.TextBox txtCGC 
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
         Left            =   210
         MaxLength       =   18
         TabIndex        =   6
         Top             =   360
         Width           =   2445
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
         Left            =   3270
         TabIndex        =   38
         Top             =   1290
         Width           =   1065
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
         Left            =   1770
         TabIndex        =   37
         Top             =   1290
         Width           =   1065
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
         Left            =   210
         TabIndex        =   36
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Label lblCG 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social / Nome do Responsável"
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
         TabIndex        =   25
         Top             =   720
         Width           =   3705
      End
      Begin VB.Label lblSigla 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "C.G.C / C.P.F do Responsável"
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
         Left            =   240
         TabIndex        =   24
         Top             =   120
         Width           =   2985
      End
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Empresa Mestre"
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
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   9555
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   375
         Left            =   8010
         TabIndex        =   22
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
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   570
         Width           =   6795
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   8580
         TabIndex        =   3
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblCódigo 
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
         TabIndex        =   21
         Top             =   330
         Width           =   885
      End
      Begin VB.Label lblDescricao 
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
         TabIndex        =   20
         Top             =   330
         Width           =   1935
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1950
      Top             =   6390
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
            Picture         =   "frm_crudEmpresa.frx":3627
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresa.frx":3BC1
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresa.frx":415B
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresa.frx":46F5
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresa.frx":4C8F
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudEmpresa.frx":5229
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroEmpresa 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
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
      MouseIcon       =   "frm_crudEmpresa.frx":57C3
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6960
      Width           =   10080
      _ExtentX        =   17780
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
   Begin VB.Image ImageDefault 
      BorderStyle     =   1  'Fixed Single
      Height          =   2265
      Left            =   4980
      Picture         =   "frm_crudEmpresa.frx":5D5D
      Stretch         =   -1  'True
      Top             =   2220
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frm_crudEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCreditoFacil       As New ControladorCreditoFacil
Private oEndereco           As New clsENDERECO
Private mstrTipoOperacao    As String
Private mCtLock             As Long
Private mCtLockEndereco     As Long

Private Sub cboCidade_Click()
  PopulaComboBairro
End Sub

Private Sub cboEstado_Click()
  PopulaComboCidade
End Sub

Private Sub cmdLimpar_Click()
    If MsgBox("Você deseja realmente limpar a imagem ?", vbYesNo, "CONFIRMAÇÃO") = vbYes Then
        imgLogoEmpresa.Picture = ImageDefault.Picture
        txtImagePath = ""
    End If
End Sub

Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    If Screen.ActiveControl.Name = "cmdSelecaoEntidade" Then
    
        Set frmPesquisa.rsResultset = oCreditoFacil.oEmpresa.recuperarEmpresas()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_EMPRESA"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "NOME_FANTASIA"
        frmPesquisa.Caption = frmPesquisa.Caption & " Empresas Cadastradas"
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
      imgLogoEmpresa.Picture = LoadPicture(dlgFOTO.FileName)
      txtImagePath.Text = dlgFOTO.FileName
        On Error GoTo 0
         Exit Sub
erro:
        If Err.Number = 75 Then
            'MsgBox "Cancelado pelo Usuário"
        End If
End Sub

Private Sub Form_Load()

FraCampos.Visible = False
oCreditoFacil.oEmpresa.m_timeOut = gstrTimeOutGeral
oCreditoFacil.oEmpresa.m_stringConexao = gstrConexaoCreditoFacil
oCreditoFacil.oEstado.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil

mstrTipoOperacao = ""
ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False

'Popula estados da federação no combo
PopulaComboEstado

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set oCreditoFacil = Nothing
    Set oEndereco = Nothing

End Sub

Private Sub ToolbarCadastroEmpresa_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
  Case 1 'Novo
    Novo_Click
    'PreencharCamposTeste
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

If Not ValidarInsertEndereco Then
    MsgBox "O endereço não está preenchido completamente. Complete o cadastro e tente novamente.", vbInformation, "MENSAGEM"
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
ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False

End Sub
Private Sub Novo_Click()

txtDescricaoEntidade.SetFocus
txtCodigoEntidade.Text = oCreditoFacil.oEmpresa.getNovoIdEmpresa
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
'Validar Campos do Formulario
Private Function ValidarInsert() As Boolean

Dim rsFuncao As ADODB.Recordset
    
  If Trim(Len(txtDescricaoEntidade)) = 0 Then
     MsgBox "A nome fantasia da empresa é requerido.", vbInformation, "Mensagem"
     txtDescricaoEntidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtRazaoSocial)) = 0 Then
     MsgBox "A razão social da empresa é requerida. Sugestão: Você pode repetir o nome fantasia", vbInformation, "Mensagem"
     txtRazaoSocial.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  ValidarInsert = True
  
End Function
'Validar Campos do Formulario
Private Function ValidarInsertEndereco() As Boolean

Dim rsFuncao As ADODB.Recordset
    
  If Trim(Len(txtTipoLogradouro)) = 0 Then
     ValidarInsertEndereco = False
     Exit Function
  End If
  
  If Trim(Len(txtLogradouro)) = 0 Then
     ValidarInsertEndereco = False
     Exit Function
  End If
  
  If Trim(Len(txtNumero)) = 0 Then
     ValidarInsertEndereco = False
     Exit Function
  End If
  
  If Trim(Len(txtTipoLogradouro)) = 0 Then
     ValidarInsertEndereco = False
     Exit Function
  End If
  
  If cboEstado.ListIndex = -1 Or cboCidade.ListIndex = -1 Or cboBairro.ListIndex = -1 Then
     ValidarInsertEndereco = False
     Exit Function
  End If
  
  ValidarInsertEndereco = True
  
End Function

Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
On Error GoTo trataerro
    
    'Atributos da empresa
    oCreditoFacil.oEmpresa.m_01_idEmpresa = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
    oCreditoFacil.oEmpresa.m_02_cgcEmpresa = txtCGC.Text
    oCreditoFacil.oEmpresa.m_03_razaoSocial = txtRazaoSocial
    oCreditoFacil.oEmpresa.m_04_nomeFantasia = txtDescricaoEntidade
    oCreditoFacil.oEmpresa.m_05_idEndereco = IIf(txtIDEndereco = "", 0, txtIDEndereco)
    oCreditoFacil.oEmpresa.m_06_telefone1 = txtTelefone1
    oCreditoFacil.oEmpresa.m_07_telefone2 = txtTelefone2
    oCreditoFacil.oEmpresa.m_08_telefone3 = txtTelefone3
    oCreditoFacil.oEmpresa.m_09_blobLogEmpresa = txtImagePath
    If txtID.Text = "" Then
      oCreditoFacil.oEmpresa.m_10_dataInclusao = Now
      oCreditoFacil.oEmpresa.m_11_usuarioInclusao = LogInUserID
    End If
    oCreditoFacil.oEmpresa.m_12_dataAlteracao = Now
    oCreditoFacil.oEmpresa.m_13_usuarioAlteracao = LogInUserID
    oCreditoFacil.oEmpresa.m_14_ctLock = mCtLock
    
    'Atributos de endereço
    oEndereco.inicializaEndereco
    oEndereco.m_timeOut = gstrTimeOutGeral
    oEndereco.m_stringConexao = gstrConexaoCreditoFacil
    With oEndereco.rsEndereco
        .Open
        .AddNew
        .Fields("ID_OBJECT_ENTIDADE") = oEndereco.consultaIdObjectEntidade("empresa")
        .Fields("ID_ENTIDADE") = CLng(txtCodigoEntidade)
        .Fields("ID_ENDERECO") = IIf(txtIDEndereco = "", 0, txtIDEndereco)
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
    
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oEmpresa.crudInsert(oEndereco.rsEndereco)
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oEmpresa.crudUpdate(oEndereco.rsEndereco)
    Else
        txtID.Text = oCreditoFacil.oEmpresa.crudDelete(oEndereco.rsEndereco)
    End If
    
trataerro:
    If InStr(1, Err.Description, "FK_funcionario_empresa") > 0 And Err.Number = -2147221503 Then
        MsgBox "Não é possível excluir a empresa pois já possui funcionários cadastrados", vbInformation, "NÃO FOI POSSÍVEL EXCLUIR"
        txtID.Text = 0
    End If
End Sub
Private Sub Limpacampos()

FraCampos.Visible = True
txtDescricaoEntidade.Text = ""
txtCGC.Text = ""
txtRazaoSocial = ""
txtTelefone1 = ""
txtTelefone2 = ""
txtTelefone3 = ""
imgLogoEmpresa.Picture = ImageDefault.Picture
txtImagePath = ""

End Sub

Private Sub LimpaCamposEndereco()

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

Private Sub txtCEP_Change()
    If Len(txtCEP) = 5 Then
    txtCEP = txtCEP + "-"
    txtCEP.SelStart = 7
    End If
End Sub

Private Sub txtCodigoEntidade_Change()

Limpacampos
LimpaCamposEndereco
ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = False
txtID.Text = ""
txtIDEndereco = ""
stbmsg.SimpleText = ""

End Sub

Private Sub txtCodigoEntidade_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCodigoEntidade_LostFocus()

Dim rsEmpresa As ADODB.Recordset
Dim rsEndereco As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
        
    Set rsEmpresa = oCreditoFacil.oEmpresa.consulta(txtCodigoEntidade)
    oEndereco.m_timeOut = gstrTimeOutGeral
    oEndereco.m_stringConexao = gstrConexaoCreditoFacil
    Set rsEndereco = oEndereco.recuperarEndereco(oEndereco.consultaIdObjectEntidade("empresa"), txtCodigoEntidade)
    
    If rsEmpresa.EOF Then
      If oCreditoFacil.oEmpresa.getNovoIdEmpresa <> txtCodigoEntidade Then
        ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False
      End If
      Exit Sub
    End If
    
    FraCampos.Visible = True
    ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = True
    ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rsEmpresa, rsEndereco
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rsEmpresa, ByRef rsEndereco As ADODB.Recordset)
    
    Dim i As Integer
    
    txtID.Text = rsEmpresa("ID_EMPRESA")
    mCtLock = rsEmpresa("CT_LOCK")
        
    txtDescricaoEntidade = rsEmpresa("NOME_FANTASIA")
    txtCGC = rsEmpresa("CGC_EMPRESA")
    txtRazaoSocial = rsEmpresa("RAZAO_SOCIAL")
    
    'IdEndereco
    txtIDEndereco = IIf(IsNull(rsEmpresa("ID_ENDERECO")), "", rsEmpresa("ID_ENDERECO"))
    
    txtTelefone1 = rsEmpresa("TELEFONE1")
    txtTelefone2 = rsEmpresa("TELEFONE2")
    txtTelefone3 = rsEmpresa("TELEFONE3")
    
    'Carregamento da imagem 3x4
    If Not IsNull(rsEmpresa("BLOB_LOGO_EMPRESA")) Then
        txtImagePath = oCreditoFacil.oEmpresa.carregarImagem(rsEmpresa)
        imgLogoEmpresa.Picture = LoadPicture(txtImagePath)
    End If
    
    
    If Not rsEndereco.EOF Then
    
      txtIDEndereco.Text = rsEndereco("ID_ENDERECO")
      mCtLockEndereco = rsEndereco("CT_LOCK")
      
      txtTipoLogradouro = rsEndereco("TIPO_LOGRADOURO")
      txtLogradouro = rsEndereco("LOGRADOURO")
      txtNumero = rsEndereco("NUMERO")
      txtComplemento = IIf(IsNull(rsEndereco("COMPLEMENTO")), "", rsEndereco("COMPLEMENTO"))
      txtPontoReferencia = IIf(IsNull(rsEndereco("PONTO_REFERENCIA")), "", rsEndereco("PONTO_REFERENCIA"))
      txtCEP = IIf(IsNull(rsEndereco("CEP")), "", rsEndereco("CEP"))
      For i = 1 To cboEstado.ListCount
        If cboEstado.ItemData(i - 1) = rsEndereco("ID_ESTADO") Then
          cboEstado.ListIndex = i - 1
        End If
      Next
      For i = 1 To cboCidade.ListCount
        If cboCidade.ItemData(i - 1) = rsEndereco("ID_MUNICIPIO") Then
          cboCidade.ListIndex = i - 1
        End If
      Next
      For i = 1 To cboBairro.ListCount
        If cboBairro.ItemData(i - 1) = rsEndereco("ID_BAIRRO") Then
          cboBairro.ListIndex = i - 1
        End If
      Next
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

Private Sub PreencharCamposTeste()

    txtDescricaoEntidade = "AQUAT-888 ENTERPRISE"
    txtCGC = "1.585.654/5011-56"
    txtRazaoSocial = "FABRICA DE SOFTWARE"
    txtTelefone1 = "8803-7269"
    txtTelefone2 = "8803-7270"
    txtTelefone3 = "8803-7271"
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
    
    
        
End Sub
