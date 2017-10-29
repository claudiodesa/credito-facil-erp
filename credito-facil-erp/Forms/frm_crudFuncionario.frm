VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{21F8D070-6EA8-40F7-8555-9E5FA3E03CB5}#1.0#0"; "calendario.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_crudFuncionario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Funcionários"
   ClientHeight    =   8175
   ClientLeft      =   5775
   ClientTop       =   2910
   ClientWidth     =   9810
   Icon            =   "frm_crudFuncionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraEndereço 
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
      Height          =   2325
      Left            =   120
      TabIndex        =   38
      Top             =   5490
      Width           =   9555
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
         TabIndex        =   15
         Top             =   510
         Width           =   915
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
         TabIndex        =   16
         Top             =   510
         Width           =   6555
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
         TabIndex        =   17
         Top             =   510
         Width           =   1365
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
         TabIndex        =   18
         Top             =   1140
         Width           =   2715
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
         TabIndex        =   19
         Top             =   1140
         Width           =   6135
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
         TabIndex        =   20
         Top             =   1740
         Width           =   1035
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
         TabIndex        =   23
         Top             =   1740
         Width           =   3495
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1740
         Width           =   825
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
         TabIndex        =   22
         Top             =   1740
         Width           =   3495
      End
      Begin VB.TextBox txtIDEndereco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   8850
         TabIndex        =   39
         Top             =   150
         Width           =   645
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
         TabIndex        =   48
         Top             =   240
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
         TabIndex        =   47
         Top             =   240
         Width           =   1035
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
         TabIndex        =   46
         Top             =   240
         Width           =   945
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
         TabIndex        =   45
         Top             =   870
         Width           =   1575
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
         TabIndex        =   44
         Top             =   870
         Width           =   1575
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
         TabIndex        =   43
         Top             =   1470
         Width           =   1035
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   1470
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
         TabIndex        =   40
         Top             =   1470
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
      Height          =   3495
      Left            =   120
      TabIndex        =   30
      Top             =   1920
      Width           =   9555
      Begin VB.ComboBox cboEmpresa 
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
         TabIndex        =   3
         Top             =   360
         Width           =   4845
      End
      Begin VB.Frame fraFoto3x4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Foto3x4"
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
         Left            =   6390
         TabIndex        =   52
         Top             =   810
         Width           =   2865
         Begin VB.TextBox txtImagePath 
            Height          =   315
            Left            =   30
            TabIndex        =   57
            Top             =   180
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton cmdTrocar 
            Caption         =   "Trocar"
            Height          =   405
            Left            =   90
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   960
            Width           =   915
         End
         Begin VB.CommandButton cmdLimpar 
            Caption         =   "Limpar"
            Height          =   405
            Left            =   90
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1380
            Width           =   915
         End
         Begin MSComDlg.CommonDialog dlgFOTO 
            Left            =   300
            Top             =   420
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image imgFotoFuncionario 
            BorderStyle     =   1  'Fixed Single
            Height          =   2265
            Left            =   1110
            Picture         =   "frm_crudFuncionario.frx":058A
            Stretch         =   -1  'True
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.ComboBox cboSituacao 
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
         ItemData        =   "frm_crudFuncionario.frx":29A3
         Left            =   1980
         List            =   "frm_crudFuncionario.frx":29A5
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1185
      End
      Begin Calendario.ctlPicker CtlDataAdmissao 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   960
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
      Begin VB.ComboBox cboFuncao 
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
         Left            =   6390
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cboSexo 
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
         ItemData        =   "frm_crudFuncionario.frx":29A7
         Left            =   4200
         List            =   "frm_crudFuncionario.frx":29A9
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1650
         Width           =   795
      End
      Begin VB.TextBox txtRG 
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
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1650
         Width           =   2295
      End
      Begin VB.TextBox txtCPF 
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
         Left            =   180
         MaxLength       =   14
         TabIndex        =   8
         Top             =   1650
         Width           =   1635
      End
      Begin VB.TextBox txtObservacoes 
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
         Left            =   180
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3000
         Width           =   4695
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
         TabIndex        =   11
         Top             =   2370
         Width           =   1485
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
         TabIndex        =   12
         Top             =   2370
         Width           =   1485
      End
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
         Left            =   3330
         MaxLength       =   13
         TabIndex        =   13
         Top             =   2370
         Width           =   1485
      End
      Begin Calendario.ctlPicker CtlDataDemissao 
         Height          =   315
         Left            =   3330
         TabIndex        =   6
         Top             =   960
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
         Left            =   6060
         Picture         =   "frm_crudFuncionario.frx":29AB
         Stretch         =   -1  'True
         Top             =   1020
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblEmpresa 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         TabIndex        =   56
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblDataDemissao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Demissão"
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
         Left            =   3330
         TabIndex        =   55
         Top             =   720
         Width           =   1635
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
         Left            =   1980
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDataAdmissao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Admissão"
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
         TabIndex        =   50
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label lblFuncao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Função"
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
         Left            =   6390
         TabIndex        =   49
         Top             =   120
         Width           =   1125
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
         Left            =   4200
         TabIndex        =   37
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label lblRG 
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
         Left            =   1860
         TabIndex        =   36
         Top             =   1380
         Width           =   1185
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
         Left            =   180
         TabIndex        =   35
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label lblObservacoes 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
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
         TabIndex        =   34
         Top             =   2730
         Width           =   1305
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
         TabIndex        =   33
         Top             =   2100
         Width           =   1005
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
         TabIndex        =   32
         Top             =   2100
         Width           =   1005
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
         Left            =   3330
         TabIndex        =   31
         Top             =   2100
         Width           =   1005
      End
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Funcionário"
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
      TabIndex        =   25
      Top             =   720
      Width           =   9555
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   8580
         TabIndex        =   27
         Top             =   150
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
         TabIndex        =   2
         Top             =   570
         Width           =   6765
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
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   885
      End
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   375
         Left            =   7980
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   570
         Width           =   465
      End
      Begin VB.Label lblDescricao 
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
         Left            =   1200
         TabIndex        =   29
         Top             =   330
         Width           =   1965
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
         TabIndex        =   28
         Top             =   330
         Width           =   915
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   2640
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
            Picture         =   "frm_crudFuncionario.frx":4DC4
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFuncionario.frx":535E
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFuncionario.frx":58F8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFuncionario.frx":5E92
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFuncionario.frx":642C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFuncionario.frx":69C6
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
      Width           =   9810
      _ExtentX        =   17304
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
      MouseIcon       =   "frm_crudFuncionario.frx":6F60
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   7920
      Width           =   9810
      _ExtentX        =   17304
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
Attribute VB_Name = "frm_crudFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCreditoFacil       As New ControladorCreditoFacil
Private mstrTipoOperacao    As String
Private mCtLock             As Long
Private mCtLockEndereco     As Long
Private Sub PopulaComboSexo()
    CboSexo.AddItem "M", 0
    CboSexo.AddItem "F", 1
End Sub
Private Sub PopulaComboFuncao()
    Dim rs As ADODB.Recordset
    Set rs = oCreditoFacil.oFuncao.RecuperarFuncoes(gstrConexaoCreditoFacil, gstrTimeOutGeral)
    If Not rs.EOF Then
        cboFuncao.Clear
        Do While Not rs.EOF
            cboFuncao.AddItem rs("DESCRICAO_FUNCAO")
            cboFuncao.ItemData(cboFuncao.NewIndex) = rs("ID_FUNCAO")
            rs.MoveNext
        Loop
    End If
    cboFuncao.ListIndex = -1
End Sub
Private Sub PopulaComboEmpresa()
    Dim rs As ADODB.Recordset
    oCreditoFacil.oEmpresa.m_timeOut = gstrTimeOutGeral
    oCreditoFacil.oEmpresa.m_stringConexao = gstrConexaoCreditoFacil
    Set rs = oCreditoFacil.oEmpresa.recuperarEmpresas()
    If Not rs.EOF Then
        cboEmpresa.Clear
        Do While Not rs.EOF
            cboEmpresa.AddItem rs("NOME_FANTASIA")
            cboEmpresa.ItemData(cboEmpresa.NewIndex) = rs("ID_EMPRESA")
            rs.MoveNext
        Loop
    End If
    cboEmpresa.ListIndex = -1
End Sub
Private Sub PopulaComboSituacao()
    cboSituacao.AddItem "Ativo", 0
    cboSituacao.AddItem "Desligado", 1
End Sub

Private Sub cboCidade_Click()
    PopulaComboBairro
End Sub

Private Sub cboEstado_Click()
    PopulaComboCidade
End Sub

Private Sub cmdLimpar_Click()
    If MsgBox("Você deseja realmente limpar a imagem ?", vbYesNo, "CONFIRMAÇÃO") = vbYes Then
        imgFotoFuncionario.Picture = ImageDefault.Picture
        txtImagePath = ""
    End If
End Sub

Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    If Screen.ActiveControl.Name = "cmdSelecaoEntidade" Then
    
        oCreditoFacil.oFuncionario.m_timeOut = gstrTimeOutGeral
        oCreditoFacil.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
        Set frmPesquisa.rsResultset = oCreditoFacil.oFuncionario.RecuperarFuncionarios()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_FUNCIONARIO"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "NOME"
        frmPesquisa.Caption = frmPesquisa.Caption & " Funcionários Cadastrados"
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
      imgFotoFuncionario.Picture = LoadPicture(dlgFOTO.FileName)
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
    
    oCreditoFacil.oFuncionario.m_timeOut = gstrTimeOutGeral
    oCreditoFacil.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
    
    oCreditoFacil.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oCreditoFacil.oEstado.mTIMEOUT = gstrTimeOutGeral
    
    mstrTipoOperacao = ""
    ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False
    ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = False
    
    PopulaComboEmpresa
    PopulaComboEstado
    PopulaComboSexo
    PopulaComboFuncao
    PopulaComboSituacao
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoFacil = Nothing
End Sub

Private Sub ToolbarCadastroEmpresa_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
      Case 1 'Novo
        Novo_Click
        'PreencheCamposTeste
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
'Validar Campos do Formulario
Private Function ValidarInsert() As Boolean

  If Trim(Len(txtDescricaoEntidade)) = 0 Then
     MsgBox "A nome completo do funcionário é requerido.", vbInformation, "Mensagem"
     txtDescricaoEntidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtCPF)) = 0 Then
     MsgBox "O CPF do funcionário é requerido.", vbInformation, "Mensagem"
     txtCPF.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtRG)) = 0 Then
     MsgBox "O RG do funcionário é requerido.", vbInformation, "Mensagem"
     txtRG.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If CboSexo.ListIndex = -1 Then
     MsgBox "O sexo do funcionário é requerido.", vbInformation, "Mensagem"
     CboSexo.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If cboFuncao.ListIndex = -1 Then
     MsgBox "A função do funcionário deve ser informada.", vbInformation, "Mensagem"
     cboFuncao.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If CtlDataAdmissao.Text = "__/__/____" Then
     MsgBox "A data de admissão do funcionário é requerida.", vbInformation, "Mensagem"
     CtlDataAdmissao.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If cboSituacao.ListIndex = -1 Then
     MsgBox "A situação do funcionário é requerida.", vbInformation, "Mensagem"
     cboFuncao.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
   If cboEmpresa.ListIndex = -1 Then
     MsgBox "A empresa do funcionário é requerida.", vbInformation, "Mensagem"
     cboEmpresa.SetFocus
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
    oCreditoFacil.oFuncionario.m_01_idFuncionario = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
    oCreditoFacil.oFuncionario.m_02_nome = txtDescricaoEntidade.Text
    oCreditoFacil.oFuncionario.m_03_cpf = txtCPF
    oCreditoFacil.oFuncionario.m_04_rg = txtRG
    oCreditoFacil.oFuncionario.m_05_sexo = CboSexo.Text
    oCreditoFacil.oFuncionario.m_06_idEndereco = IIf(txtIDEndereco = "", 0, txtIDEndereco)
    oCreditoFacil.oFuncionario.m_07_telefone1 = txtTelefone1
    oCreditoFacil.oFuncionario.m_08_telefone2 = txtTelefone2
    oCreditoFacil.oFuncionario.m_09_telefone3 = txtTelefone3
    oCreditoFacil.oFuncionario.m_10_blobFoto = txtImagePath
    oCreditoFacil.oFuncionario.m_11_idFuncao = cboFuncao.ItemData(cboFuncao.ListIndex)
    oCreditoFacil.oFuncionario.m_12_dataAdmissao = CtlDataAdmissao.Text
    oCreditoFacil.oFuncionario.m_13_situacao = Mid(cboSituacao.Text, 1, 1)
    oCreditoFacil.oFuncionario.m_14_observacoes = txtObservacoes
    If txtID.Text = "" Then
      oCreditoFacil.oFuncionario.m_16_dataInclusao = Now
      oCreditoFacil.oFuncionario.m_17_usuarioInclusao = LogInUserID
    End If
    oCreditoFacil.oFuncionario.m_18_dataAlteracao = Now
    oCreditoFacil.oFuncionario.m_19_usuarioAlteracao = LogInUserID
    oCreditoFacil.oFuncionario.m_20_ctLock = mCtLock
    oCreditoFacil.oFuncionario.m_21_idEmpresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)
    oCreditoFacil.oFuncionario.m_22_dataDemissao = IIf(CtlDataDemissao.Text <> "__/__/____", CtlDataDemissao.Text, "00:00:00")
    
    'Atributos de endereço
    oCreditoFacil.oEndereco.inicializaEndereco
    With oCreditoFacil.oEndereco.rsEndereco
        .Open
        .AddNew
        .Fields("ID_OBJECT_ENTIDADE") = oCreditoFacil.oEndereco.consultaIdObjectEntidade("funcionario")
        .Fields("ID_ENTIDADE") = txtCodigoEntidade
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
        End If
        .Fields("USUARIO_ALTERACAO") = LogInUserID
        .Fields("DATA_ALTERACAO") = CStr(Now)
        .Fields("CT_LOCK") = mCtLockEndereco
        .Update
    End With
    
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oFuncionario.crudInsert(oCreditoFacil.oEndereco.rsEndereco)
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oFuncionario.crudUpdate(oCreditoFacil.oEndereco.rsEndereco)
    Else
        txtIDEndereco.Text = oCreditoFacil.oFuncionario.crudDelete(oCreditoFacil.oEndereco.rsEndereco)
    End If
    
trataerro:
 If InStr(1, Err.Description, "FK_usuario_funcionario") > 0 And Err.Number = -2147221503 Then
    MsgBox "Não é possível excluir o funcionário, pois o mesmo ainda está ligado a um usuário/login.", vbInformation, "EXCLUSÃO NÃO FOI REALIZADA"
    txtID = 0
 End If
 If InStr(1, Err.Description, "FK_rota_funcionario") > 0 And Err.Number = -2147221503 Then
    MsgBox "Não é possível excluir o funcionário, pois o mesmo ainda está ligado a uma rota.", vbInformation, "EXCLUSÃO NÃO FOI REALIZADA"
    txtID = 0
 End If
    
    
End Sub
Private Sub Limpacampos()

    FraCampos.Visible = True
    txtDescricaoEntidade.Text = ""
    cboEmpresa.ListIndex = -1
    CtlDataAdmissao.Text = "__/__/____"
    cboSituacao.ListIndex = -1
    CtlDataDemissao.Text = "__/__/____"
    cboFuncao.ListIndex = -1
    txtCPF = ""
    txtRG = ""
    CboSexo.ListIndex = -1
    txtTelefone1 = ""
    txtTelefone2 = ""
    txtTelefone3 = ""
    txtObservacoes = ""
    imgFotoFuncionario.Picture = ImageDefault.Picture
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
ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False
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

Dim rsFuncionario As ADODB.Recordset
Dim rsEndereco As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
        
    Set rsFuncionario = oCreditoFacil.oFuncionario.consulta(txtCodigoEntidade)
    oCreditoFacil.oEndereco.m_timeOut = gstrTimeOutGeral
    oCreditoFacil.oEndereco.m_stringConexao = gstrConexaoCreditoFacil
    Set rsEndereco = oCreditoFacil.oEndereco.recuperarEndereco(oCreditoFacil.oEndereco.consultaIdObjectEntidade("funcionario"), txtCodigoEntidade)
    
    If rsFuncionario.EOF Then
      If oCreditoFacil.oFuncionario.GetNovoIDFuncionario <> txtCodigoEntidade Then
        ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = False
      End If
      Exit Sub
    End If
    
    FraCampos.Visible = True
    ToolbarCadastroEmpresa.Buttons("Excluir").Enabled = True
    ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rsFuncionario, rsEndereco
    txtDescricaoEntidade.SetFocus
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rsFuncionario, ByRef rsEndereco As ADODB.Recordset)
    
    Dim i As Integer
    
    txtID.Text = rsFuncionario("ID_FUNCIONARIO")
    mCtLock = rsFuncionario("CT_LOCK")
       
    txtDescricaoEntidade = rsFuncionario("NOME")
    For i = 1 To cboEmpresa.ListCount
        If cboEmpresa.ItemData(i - 1) = rsFuncionario("ID_EMPRESA") Then
            cboEmpresa.ListIndex = i - 1
        End If
    Next
    CtlDataAdmissao.Text = Format(rsFuncionario("DATA_ADMISSAO"), "dd/mm/yyyy")
    cboSituacao.ListIndex = IIf(rsFuncionario("SITUACAO") = "A", 0, 1)
    CtlDataDemissao.Text = IIf(IsNull(rsFuncionario("DATA_DEMISSAO")), "__/__/____", Format(rsFuncionario("DATA_DEMISSAO"), "dd/mm/yyyy"))
    For i = 1 To cboFuncao.ListCount
        If cboFuncao.ItemData(i - 1) = rsFuncionario("ID_FUNCAO") Then
            cboFuncao.ListIndex = i - 1
        End If
    Next
    txtCPF = rsFuncionario("CPF")
    txtRG = rsFuncionario("RG")
    CboSexo.ListIndex = IIf(rsFuncionario("SEXO") = "M", 0, 1)
    txtTelefone1 = rsFuncionario("TELEFONE1")
    txtTelefone2 = rsFuncionario("TELEFONE2")
    txtTelefone3 = rsFuncionario("TELEFONE3")
    txtObservacoes = rsFuncionario("OBSERVACOES")
    'IdEndereco
    txtIDEndereco = IIf(IsNull(rsFuncionario("ID_ENDERECO")), "", rsFuncionario("ID_ENDERECO"))
    
    'Carregamento da imagem 3x4
    If Not IsNull(rsFuncionario("BLOB_FOTO")) Then
        txtImagePath = oCreditoFacil.oFuncionario.carregarImagem(rsFuncionario)
        imgFotoFuncionario.Picture = LoadPicture(txtImagePath)
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
Private Sub Novo_Click()

txtCodigoEntidade.Text = oCreditoFacil.oFuncionario.GetNovoIDFuncionario
txtDescricaoEntidade.SetFocus
ToolbarCadastroEmpresa.Buttons("Salvar").Enabled = True
mstrTipoOperacao = "I"
stbmsg.SimpleText = "Incluindo"
DoEvents

End Sub
Private Sub Excluir_Click()

If txtCodigoEntidade = "" Or txtCodigoEntidade = "0" Then Exit Sub

    Dim rs As ADODB.Recordset
    
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

Private Sub txtCPF_Change()
    If Len(txtCPF) = 3 Then
    txtCPF = txtCPF + "."
    txtCPF.SelStart = 5
    End If
    If Len(txtCPF) = 7 Then
    txtCPF = txtCPF + "."
    txtCPF.SelStart = 9
    End If
    If Len(txtCPF) = 11 Then
    txtCPF = txtCPF + "-"
    txtCPF.SelStart = 14
    End If
End Sub
Private Sub PreencheCamposTeste()
    'FUNCIONARIO
    txtDescricaoEntidade = "MATHEUS GIRAO"
    cboEmpresa.ListIndex = 0
    cboFuncao.ListIndex = 0
    CtlDataAdmissao.Text = "15/02/2000"
    cboSituacao.ListIndex = 0
    'CtlDataDemissao.Text = "31/12/2500"
    txtCPF = "617.985.985-88"
    txtRG = "12345600"
    CboSexo.ListIndex = 0
    txtTelefone1 = "8803-7269"
    txtTelefone2 = "8803-7270"
    txtTelefone3 = "8803-7271"
    txtObservacoes = "De ferias"
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
