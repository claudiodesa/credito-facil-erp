VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_processoGerarFinanciamento 
   Caption         =   "Gerar Financiamento de Crédito"
   ClientHeight    =   7185
   ClientLeft      =   4110
   ClientTop       =   855
   ClientWidth     =   10335
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_processoGerarFinanciamento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Financiamento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   60
      TabIndex        =   0
      Top             =   750
      Width           =   10215
      Begin VB.TextBox txtJurosMora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6690
         MaxLength       =   6
         TabIndex        =   37
         Text            =   "0,75"
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Frame FraCobranca 
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
         Height          =   885
         Left            =   5040
         TabIndex        =   35
         Top             =   4380
         Width           =   4875
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
            TabIndex        =   36
            Top             =   360
            Width           =   4545
         End
      End
      Begin VB.Frame FraPeriodoFinanciamento 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Período do Financiamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   5340
         TabIndex        =   28
         Top             =   2880
         Width           =   4575
         Begin MSComCtl2.DTPicker DTPickerInicio 
            Height          =   345
            Left            =   1140
            TabIndex        =   29
            Top             =   510
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   17629185
            CurrentDate     =   40758
         End
         Begin MSComCtl2.DTPicker DTPickerFim 
            Height          =   345
            Left            =   2790
            TabIndex        =   30
            Top             =   510
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   17629185
            CurrentDate     =   40758
         End
         Begin VB.Label lblUltimaParcela 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Última parcela"
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
            Left            =   2790
            TabIndex        =   32
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label lblParcela1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "1ª Parcela"
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
            Left            =   1140
            TabIndex        =   31
            Top             =   300
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Valores e Cálculos do Empréstimo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   390
         TabIndex        =   17
         Top             =   2880
         Width           =   4605
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
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
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   3030
            MaxLength       =   10
            TabIndex        =   26
            Text            =   "0,00"
            Top             =   1770
            Width           =   1485
         End
         Begin VB.TextBox txtTaxa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3030
            MaxLength       =   6
            TabIndex        =   24
            Text            =   "15,00"
            Top             =   1020
            Width           =   1485
         End
         Begin VB.TextBox txtParcelasValor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
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
            Height          =   390
            Left            =   3030
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "0,00"
            Top             =   1380
            Width           =   1485
         End
         Begin VB.TextBox txtParcelas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3030
            MaxLength       =   2
            TabIndex        =   20
            Text            =   "25"
            Top             =   660
            Width           =   1485
         End
         Begin VB.TextBox txtValorFinanciado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3030
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Total à Pagar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   210
            TabIndex        =   27
            Top             =   1800
            Width           =   2745
         End
         Begin VB.Label lblTaxa 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Taxa (%)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   210
            TabIndex        =   25
            Top             =   1050
            Width           =   2745
         End
         Begin VB.Label lblParcelasValor 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Valor Por Parcela"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   210
            TabIndex        =   23
            Top             =   1410
            Width           =   2745
         End
         Begin VB.Label lblParcelas 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Número de Parcelas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   210
            TabIndex        =   21
            Top             =   690
            Width           =   2745
         End
         Begin VB.Label lblValorFianciamento 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Valor do empréstimo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   210
            TabIndex        =   19
            Top             =   300
            Width           =   2745
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dados da Pré-aprovação do crédito"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   390
         TabIndex        =   8
         Top             =   1620
         Width           =   9525
         Begin VB.TextBox txtIDLinhaCredito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   8580
            TabIndex        =   38
            Text            =   "ID"
            Top             =   150
            Width           =   885
         End
         Begin VB.TextBox txtSituacao 
            Alignment       =   1  'Right Justify
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
            Left            =   7110
            MaxLength       =   14
            TabIndex        =   15
            Top             =   630
            Width           =   2265
         End
         Begin VB.TextBox txtDataLiberacao 
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
            Left            =   3750
            MaxLength       =   20
            TabIndex        =   13
            Top             =   630
            Width           =   1935
         End
         Begin VB.TextBox txtNomeAprovador 
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
            Left            =   180
            MaxLength       =   50
            TabIndex        =   11
            Top             =   630
            Width           =   3525
         End
         Begin VB.TextBox txtLimite 
            Alignment       =   1  'Right Justify
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
            Left            =   5730
            MaxLength       =   10
            TabIndex        =   9
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lblSituacao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Avaliação "
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
            Left            =   7380
            TabIndex        =   16
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDataLiberacao 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data da Liberação"
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
            Left            =   3750
            TabIndex        =   14
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label lblAprovador 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Gerente de Crédito"
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
            TabIndex        =   12
            Top             =   360
            Width           =   3825
         End
         Begin VB.Label lblLimiteAprovado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de Crédito"
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
            Left            =   5430
            TabIndex        =   10
            Top             =   360
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   9555
         Begin VB.CommandButton cmdSelecaoEntidade 
            Caption         =   "[...]"
            Height          =   405
            Left            =   8010
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   540
            Width           =   495
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
            TabIndex        =   3
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
            Left            =   8610
            TabIndex        =   2
            Top             =   120
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
            TabIndex        =   7
            Top             =   330
            Width           =   885
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
            Left            =   1230
            TabIndex        =   6
            Top             =   330
            Width           =   1935
         End
      End
      Begin VB.Label lblTaxaMora 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Se ultrapassar o limite de parcelas, cobrar juros de mora de (%)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         TabIndex        =   39
         Top             =   5370
         Width           =   6075
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6030
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":058A
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":0B24
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":10BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":1658
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":1BF2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":218C
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":2726
            Key             =   "Gerar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":2CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":325A
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":37F4
            Key             =   "Consultar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":3D8E
            Key             =   "Alterar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":4328
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":48C2
            Key             =   "Calc"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGerarFinanciamento.frx":4E5C
            Key             =   "Emprestimo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroLinhaCredito 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1111
      ButtonWidth     =   1402
      ButtonHeight    =   953
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Key             =   "Novo"
            Description     =   "Novo Financiamento"
            Object.ToolTipText     =   "Novo Financiamento"
            ImageIndex      =   14
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar"
            Key             =   "Gerar"
            Description     =   "Gerar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Description     =   "Imprimir Nota Promissória"
            Object.ToolTipText     =   "Imprimir Nota Promissória"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Alterar"
            Description     =   "Salvar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            Description     =   "Cancela o financiamento atual"
            Object.ToolTipText     =   "Cancela o financiamento atual"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "Consultar"
            Description     =   "Consultar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calc."
            Key             =   "Calc"
            Description     =   "Calculadora"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Description     =   "Fecha a janela atual"
            Object.ToolTipText     =   "Fecha a janela atual"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_processoGerarFinanciamento.frx":53F6
      OLEDropMode     =   1
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   6930
      Width           =   10335
      _ExtentX        =   18230
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
Attribute VB_Name = "frm_processoGerarFinanciamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private oFinanciamento As New clsFinanciamento
Private oParcelas As New clsFinanciamentoParcela
Private oControle As New ControladorCreditoFacil
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private rsVencimentos As ADODB.Recordset
Private ocaixa As New clsCaixa
Private mstrIdCaixa As Long

Private mstrTipoOperacao As String

Private Sub PrepararVencimentos()
    Set rsVencimentos = New ADODB.Recordset
    rsVencimentos.Fields.Append "PARCELA", adInteger
    rsVencimentos.Fields.Append "VENCIMENTO", adDate
End Sub
Private Function System32Path() As String

Dim lngRet As Long
    
    System32Path = Space$(255)
    lngRet = GetSystemDirectory(System32Path, 255)
    System32Path = Left$(System32Path, lngRet)
    
End Function
Private Sub GerarVencimentos(ByVal parcela As Integer, ByVal vencimento As Date)
    If rsVencimentos.State = 0 Then rsVencimentos.Open
    rsVencimentos.AddNew
    With rsVencimentos
        .Fields("PARCELA") = parcela
        .Fields("VENCIMENTO") = vencimento
        .Update
    End With
End Sub

Private Sub cmdCacularDataFim_Click()
    
End Sub

Private Function ValidarValores() As Boolean

    If Not IsNumeric(txtValorFinanciado) Then
    End If
    
End Function


Private Sub cmdSelecaoEntidade_Click()
    
    mstrTipoOperacao = "S"
    oControle.oLinhaCred.mTIMEOUT = gstrTimeOutGeral
    oControle.oLinhaCred.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    Set frmPesquisa.rsResultset = oControle.oLinhaCred.recuperarEmpresasComLinhadeCreditoAprovada()
    'Campo chave
    frmPesquisa.FieldsKey = "ID_EMPRESACLIENTE"
    'Campo a ser listado no resultado da pesquisa
    frmPesquisa.FieldsList = "NOME"
    frmPesquisa.Caption = frmPesquisa.Caption & " Empresas Com Crédito Pré-Aprovado"
    frmPesquisa.Show 1
    'Recebe retorno da pesquisa
    txtCodigoEntidade = frmPesquisa.FieldsReturn
    txtCodigoEntidade_LostFocus

End Sub

Private Sub DTPickerInicio_Change()
    Call CalcularDataFim
End Sub

Private Sub Form_Load()

oFinanciamento.m_timeOut = gstrTimeOutGeral
oFinanciamento.m_stringConexao = gstrConexaoCreditoFacil


ocaixa.mTIMEOUT = gstrTimeOutGeral
ocaixa.mSTRING_CONEXAO = gstrConexaoCreditoFacil

PopulaRotas

DTPickerInicio = Now()
DTPickerFim = Now()

Call CalcularDataFim

ToolbarCadastroLinhaCredito.Buttons("Cancelar").Enabled = False
ToolbarCadastroLinhaCredito.Buttons("Gerar").Enabled = False
ToolbarCadastroLinhaCredito.Buttons("Alterar").Enabled = False
ToolbarCadastroLinhaCredito.Buttons("Imprimir").Enabled = False

End Sub
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

Private Sub Form_Unload(Cancel As Integer)
    
    Set ocaixa = Nothing
    Set oControle = Nothing
    Set oFinanciamento = Nothing
    
End Sub

Private Sub ToolbarCadastroLinhaCredito_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
      Case "Novo" 'Novo Financiamento
        Click_Novo
      Case "Gerar" 'Gerar Parcelas
        Click_Gerar
      Case "Alterar" 'Salvar alterações do financiamento
        Click_Gerar
      Case "Cancelar" 'Cancelar o financiamento
        Click_Cancelar
      Case "Consultar" 'Consultar
        Click_Consultar
      Case "Sair" 'Fechar
        Click_Fechar
      Case "Imprimir"
        Click_Imprimir
      Case "Calc"
        Shell System32Path & "\Calc.exe"
    End Select
End Sub
Private Sub Click_Imprimir()

    If txtCodigoEntidade = "" Then
        Exit Sub
    End If
    
    If mstrTipoOperacao = "C" Or mstrTipoOperacao = "A" Then
        
        mstrTipoOperacao = "Imprimir"
        txtCodigoEntidade_LostFocus
        
    End If

End Sub
Private Sub Click_Consultar()
        
    If txtCodigoEntidade = "" Then
        MsgBox "Selecione a empresa que deseja consultar, clicando no botão [...]", vbInformation, "SELECIONE UM CLIENTE"
        Exit Sub
    End If
    
    ToolbarCadastroLinhaCredito.Buttons("Gerar").Enabled = False
    mstrTipoOperacao = "C"
    txtCodigoEntidade_LostFocus

End Sub
Private Sub Click_Cancelar()

If txtCodigoEntidade = "" Or txtCodigoEntidade = "0" Then Exit Sub
If MsgBox("Deseja realmente cancelar o financiamento? Esta opção é irreversível! ", vbYesNo, "Pergunta") = vbNo Then Exit Sub
mstrTipoOperacao = "E"
MoveTelaParaObjetoCab mstrTipoOperacao
If txtID >= 0 Then MsgBox "O financiamento foi cancelado com sucesso!", vbInformation, "SUCESSO"

Call Limpacampos
ToolbarCadastroLinhaCredito.Buttons("Imprimir").Enabled = False
ToolbarCadastroLinhaCredito.Buttons("Alterar").Enabled = False
ToolbarCadastroLinhaCredito.Buttons("Cancelar").Enabled = False
txtCodigoEntidade = ""

End Sub
Private Sub Click_Fechar()
  Unload Me
End Sub
Private Sub Click_Gerar()
    
    If Not ValidarEmprestimo Then
        Exit Sub
    End If
    
    If mstrTipoOperacao = "A" Then
        If MsgBox("Deseja realmente alterar o registro?", vbYesNo, "Pergunta") = vbNo Then Exit Sub
    End If
    
    If mstrTipoOperacao = "I" Then
        If txtValorFinanciado > ocaixa.getSaldo(ocaixa.IdUltimoCaixaAberto) Then MsgBox "Não há saldo em caixa para realizar a operação!" & " O saldo atual é de " & Format(ocaixa.getSaldo(ocaixa.IdUltimoCaixaAberto), "Currency"): Exit Sub
        If MsgBox("Confirma os dados informados para gerar este empréstimo?", vbYesNo, "Pergunta") = vbNo Then Exit Sub
    End If
    
    'Move dados para tela para o objeto
    MoveTelaParaObjetoCab mstrTipoOperacao
    
    If txtID.Text = "" Then Exit Sub
    If txtID.Text = 0 Then Exit Sub
    
    If mstrTipoOperacao = "I" Then
        MsgBox "Registro incluido com sucesso!", vbInformation, "SUCESSO"
        mstrTipoOperacao = ""
        NotaPromissoria
    ElseIf mstrTipoOperacao = "A" Then
        MsgBox "Registro alterado com sucesso!", vbInformation, "SUCESSO"
    End If
    
    Call Limpacampos
    txtCodigoEntidade.Text = ""
    ToolbarCadastroLinhaCredito.Buttons("Gerar").Enabled = False
    ToolbarCadastroLinhaCredito.Buttons("Alterar").Enabled = False
    
End Sub
Private Sub Limpacampos()

    txtValorFinanciado = ""
    txtParcelas = "25"
    txtTaxa = "15,00"
    txtParcelasValor = "0,00"
    txtTotal = "0,00"
    txtNomeAprovador = ""
    txtDataLiberacao = ""
    txtLimite = "0,00"
    txtJurosMora = "0,75"
    txtSituacao = ""
    DTPickerInicio = Date
    DTPickerFim = Date
    CalcularDataFim
    cboRotas.ListIndex = -1

End Sub
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
Dim i As Integer
    
On Error GoTo trataerro
    
    'Atributos da empresa
    oFinanciamento.m_01_ID_FINANCIAMENTO = txtID
    oFinanciamento.m_02_ID_EMPRESACLIENTE = txtCodigoEntidade
    oFinanciamento.m_03_ID_LINHACREDITO = txtIDLinhaCredito
    oFinanciamento.m_04_ID_ROTA = cboRotas.ItemData(cboRotas.ListIndex)
    oFinanciamento.m_05_VALOR_SACADO = txtValorFinanciado
    oFinanciamento.m_06_TAXA = txtTaxa
    oFinanciamento.m_07_QTD_PARCELAS = txtParcelas
    oFinanciamento.m_08_VALOR_PARCELA = txtParcelasValor
    oFinanciamento.m_09_DATA_PRIMEIRA_PARCELA = DTPickerInicio
    oFinanciamento.m_10_DATA_ULTIMA_PARCELA = DTPickerFim
    oFinanciamento.m_11_SALDO_DEVEDOR = txtTotal
    oFinanciamento.m_12_DATA_INCLUSAO = ""
    oFinanciamento.m_13_USUARIO_INCLUSAO = LogInUserID
    oFinanciamento.m_14_DATA_ALTERACAO = ""
    oFinanciamento.m_15_USUARIO_ALTERACAO = LogInUserID
    oFinanciamento.m_16_CT_LOCK = mCtLock
    oFinanciamento.m_17_ID_CAIXA = ocaixa.IdUltimoCaixaAberto
    oFinanciamento.m_18_TAXA_JUROS_MORA = txtJurosMora
    
    If mstrTipoOperacao = "I" Then
    'Parcelas
    rsVencimentos.MoveFirst
    oParcelas.inicializaParcela
    oParcelas.rsFinanciamentoParcela.Open
    For i = 1 To txtParcelas
        With oParcelas.rsFinanciamentoParcela
            .AddNew
            .Fields("ID_FINANCIAMENTO") = oFinanciamento.m_01_ID_FINANCIAMENTO
            .Fields("NUM_PARCELA") = i
            .Fields("DATA_VENCIMENTO") = Format(rsVencimentos("VENCIMENTO"), "dd/mm/yyyy")
            .Fields("VALOR_COBRADO") = Replace(txtParcelasValor, ",", ".")
            .Fields("SALDO_DEVEDOR") = Replace(txtTotal, ",", ".")
            .Fields("USUARIO_INCLUSAO") = LogInUserID
            .Fields("USUARIO_ALTERACAO") = LogInUserID
            .Update
        End With
        rsVencimentos.MoveNext
    Next
    End If
    
    If strOperacao = "I" Then
        txtID.Text = oFinanciamento.crudInsert(oParcelas.rsFinanciamentoParcela)
    ElseIf strOperacao = "A" Then
        txtID.Text = oFinanciamento.crudUpdate()
    Else
        txtID.Text = oFinanciamento.crudDelete()
    End If
    
    Exit Sub
    
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
Private Sub NotaPromissoria()

    If MsgBox("Deseja imprimir a nota promissória agora?", vbYesNo) = vbYes Then
        mstrTipoOperacao = "Imprimir"
        txtCodigoEntidade_LostFocus
    End If
    
End Sub
Private Sub Click_Novo()
    
    If txtCodigoEntidade = "" Then
        MsgBox "Selecione a empresa para a qual deseja gerar o financiamento, clicando no botão [...]", vbInformation, "SELECIONE UM CLIENTE"
        Exit Sub
    End If
    
    'Não permitir que sejam realizados empréstimos quando o caixa estiver fechado!
    If Not ocaixa.CaixaAberto Then
        MsgBox "Não é permitido gerar financiamento quando o caixa está fechado!"
        Exit Sub
    Else
        'Recupera o id do caixa aberto
        mstrIdCaixa = ocaixa.IdUltimoCaixaAberto
    End If
    
    Limpacampos
    
    mstrTipoOperacao = "I"
    txtID.Text = oFinanciamento.GetNovoIDFinanciamento
    txtCodigoEntidade_LostFocus

End Sub

Private Sub txtCodigoEntidade_Change()
    
    Limpacampos
    ToolbarCadastroLinhaCredito.Buttons("Gerar").Enabled = False
    ToolbarCadastroLinhaCredito.Buttons("Cancelar").Enabled = False
    ToolbarCadastroLinhaCredito.Buttons("Imprimir").Enabled = False
    txtID.Text = ""
    txtIDLinhaCredito = ""
    txtDescricaoEntidade = ""
    stbmsg.SimpleText = ""

End Sub

Private Sub txtCodigoEntidade_LostFocus()

Dim rsEmpCli As ADODB.Recordset
Dim rsLinha  As ADODB.Recordset
Dim i As Integer

    If txtCodigoEntidade = "" Then Exit Sub
    If mstrTipoOperacao = "" Then Exit Sub
        
    'Recupera a empresaCliente
    oControle.oEmpresaCliente.m_timeOut = gstrTimeOutGeral
    oControle.oEmpresaCliente.m_stringConexao = gstrConexaoCreditoFacil
    Set rsEmpCli = oControle.oEmpresaCliente.consulta(txtCodigoEntidade)
    
    'Caso não exista a empresaCliente, sair
    If rsEmpCli.EOF Then
        Exit Sub
    End If
    
    txtDescricaoEntidade = rsEmpCli("NOME")
    If mstrTipoOperacao = "S" Then Exit Sub
    
    'Recupera linha de crédito
    oControle.oLinhaCred.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oControle.oLinhaCred.mTIMEOUT = gstrTimeOutGeral
    Set rsLinha = oControle.oLinhaCred.consulta(rsEmpCli("ID_EMPRESACLIENTE"))
        
    MoveObjParaTela rsEmpCli, rsLinha
        
End Sub
Private Sub MoveObjParaTela(ByVal rsEmpCli As ADODB.Recordset, ByVal rsLinha As ADODB.Recordset)
        
Dim rsFinanciamentoEmpresa As ADODB.Recordset
Dim PossuiFinancimentoPendente As Boolean

    'Dados da liberação do limite de crédito ao cliente
    oControle.oFuncionario.m_timeOut = gstrTimeOutGeral
    oControle.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
    txtNomeAprovador = oControle.oFuncionario.RecuperarNomePeloLogin(rsLinha("USUARIO_ALTERACAO"))
    txtDataLiberacao = rsLinha("DATA_ALTERACAO")
    txtLimite = Format(rsLinha("LIMITE"), "0.00")
    txtSituacao = "Aprovado"
    txtIDLinhaCredito = rsLinha("ID_LINHACREDITO")
    
    'Verifica se esta empresa possui algum financiamento pendente
    If oFinanciamento.EmpresaPossuiFinanciamentoPendente(rsEmpCli("ID_EMPRESACLIENTE")) Then
        PossuiFinancimentoPendente = True
        If mstrTipoOperacao <> "Imprimir" Then stbmsg.SimpleText = "Financiamento pendente, não é possível gerar um novo neste momento."
        txtSituacao.MaxLength = 50
        txtSituacao = "Empréstimo Pendente"
        
        Set rsFinanciamentoEmpresa = oFinanciamento.consulta(rsEmpCli("ID_EMPRESACLIENTE"))
        txtID = rsFinanciamentoEmpresa("ID_FINANCIAMENTO")
        DesabilitaCamposFinanciamento
        
        txtValorFinanciado = Format(rsFinanciamentoEmpresa("VALOR_SACADO"), "0.00")
        txtParcelas = rsFinanciamentoEmpresa("QTD_PARCELAS")
        txtTaxa = Format(rsFinanciamentoEmpresa("TAXA"), "0.00")
        txtJurosMora = Format(rsFinanciamentoEmpresa("TAXA_JUROS_MORA"), "0.00")
        txtParcelasValor = Format(rsFinanciamentoEmpresa("VALOR_PARCELA"), "0.00")
        txtTotal = Format(txtParcelasValor * txtParcelas, "0.00")
        
        For i = 1 To cboRotas.ListCount
            If rsFinanciamentoEmpresa("ID_ROTA") = cboRotas.ItemData(i - 1) Then
                cboRotas.ListIndex = i - 1
                Exit For
            End If
        Next
        If mstrTipoOperacao = "I" And PossuiFinancimentoPendente Then
            
            ToolbarCadastroLinhaCredito.Buttons("Imprimir").Enabled = False
            ToolbarCadastroLinhaCredito.Buttons("Alterar").Enabled = False
            ToolbarCadastroLinhaCredito.Buttons("Gerar").Enabled = False
            Exit Sub
            
        ElseIf mstrTipoOperacao = "C" Then
        
            If MsgBox("Você deseja editar este financiamento? Você poderá: " & Chr(13) & Chr(13) & "[-] Cancelar o financimento" & Chr(13) & "[-] Alterar a rota da cobrança" & Chr(13) & "[-] Imprimir Nota Promissória", vbYesNo, "FINANCIAMENTO PENDENTE") = vbYes Then
                mstrTipoOperacao = "A"
                stbmsg.SimpleText = "Alterando"
                ToolbarCadastroLinhaCredito.Buttons("Alterar").Enabled = True
                ToolbarCadastroLinhaCredito.Buttons("Cancelar").Enabled = True
                ToolbarCadastroLinhaCredito.Buttons("Imprimir").Enabled = True
                Exit Sub
            Else
                stbmsg.SimpleText = "Consultando"
                ToolbarCadastroLinhaCredito.Buttons("Alterar").Enabled = False
                ToolbarCadastroLinhaCredito.Buttons("Cancelar").Enabled = False
                ToolbarCadastroLinhaCredito.Buttons("Imprimir").Enabled = True
                Exit Sub
            End If
        
        End If
    Else
        If mstrTipoOperacao = "C" Then
            stbmsg.SimpleText = "Nenhum empréstimo em aberto"
            Exit Sub
        End If
    End If
        
    Select Case mstrTipoOperacao
    
        Case "Imprimir"
            MoveObjParaImpressao rsFinanciamentoEmpresa
        Case "I"
            stbmsg.SimpleText = "Gerando financiamento"
            txtID = oFinanciamento.GetNovoIDFinanciamento
            ToolbarCadastroLinhaCredito.Buttons("Gerar").Enabled = True
            HabilitaCamposFinanciamento
            txtValorFinanciado.SetFocus
        Case Else
            
    End Select
    
End Sub
Private Sub MoveObjParaImpressao(ByVal rsFinanciamentoEmpresa As ADODB.Recordset)

Dim rsResponsavel As ADODB.Recordset
Dim rsEndResFin As ADODB.Recordset
Dim rsEmpCli As ADODB.Recordset
Dim rsEndEmpCli As ADODB.Recordset
Dim rsFinancimentoDetalhado As ADODB.Recordset
Dim NotaPromissoria As rel_notaPromissoria
Dim rsFuncionarioRota As ADODB.Recordset
Dim rsRota As ADODB.Recordset

    stbmsg.SimpleText = "Gerando Nota Promissória..."
    '##Campos da Nota Promissória##
    'Pega o endereço do responsável financeiro
    '@Valor tem
    '@Vencimento tem
    '@ValorExterso tem
    '@Responsavel tem
    '@Cpf tem
    '@Endereco tem
    
    oControle.oResponsavel.m_stringConexao = gstrConexaoCreditoFacil
    oControle.oResponsavel.m_timeOut = gstrTimeOutGeral
    Set rsResponsavel = oControle.oResponsavel.recuperarResponsavelFicanceiro(rsFinanciamentoEmpresa("ID_EMPRESACLIENTE"))
    oControle.oEndereco.m_stringConexao = gstrConexaoCreditoFacil
    oControle.oEndereco.m_timeOut = gstrTimeOutGeral
    Set rsEndResFin = oControle.oEndereco.recuperarEndereco(oControle.oEndereco.consultaIdObjectEntidade("responsavelFinanceiro"), rsResponsavel("ID_RESPONSAVEL"))
    
    'Recupera as parcelas
    Set rsFinancimentoDetalhado = oFinanciamento.consultaDetalhada(rsFinanciamentoEmpresa("ID_EMPRESACLIENTE"))

    oControle.oRota.m_stringConexao = gstrConexaoCreditoFacil
    oControle.oRota.m_timeOut = gstrTimeOutGeral
    Set rsRota = oControle.oRota.consulta(cboRotas.ItemData(cboRotas.ListIndex))
    oControle.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
    oControle.oFuncionario.m_timeOut = gstrTimeOutGeral
    Set rsFuncionarioRota = oControle.oFuncionario.consulta(rsRota("ID_FUNCIONARIO"))
    
    Set NotaPromissoria = New rel_notaPromissoria
    NotaPromissoria.ParameterFields(1).AddCurrentValue CStr(txtTotal)
    NotaPromissoria.ParameterFields(2).AddCurrentValue Day(DTPickerFim) & " de " & Format(DTPickerFim, "mmmm") & " de " & Year(DTPickerFim)
    NotaPromissoria.ParameterFields(3).AddCurrentValue extenso(txtTotal, "Reais", "Real")
    NotaPromissoria.ParameterFields(4).AddCurrentValue CStr(rsResponsavel("NOME"))
    NotaPromissoria.ParameterFields(5).AddCurrentValue CStr(rsResponsavel("CPF"))
    NotaPromissoria.ParameterFields(8).AddCurrentValue CStr(Mid(cboRotas.Text, 1, InStr(1, cboRotas.Text, " "))) & " / " & rsFuncionarioRota("TELEFONE1")
    'Recupera dados da empresacliente
    oControle.oEmpresaCliente.m_stringConexao = gstrConexaoCreditoFacil
    oControle.oEmpresaCliente.m_timeOut = gstrTimeOutGeral
    Set rsEmpCli = oControle.oEmpresaCliente.consulta(rsFinanciamentoEmpresa("ID_EMPRESACLIENTE"))
    If rsEmpCli("TIPO") = "F" Then
        NotaPromissoria.ParameterFields(9).AddCurrentValue rsEmpCli("NOME_PESSOA_FISICA") & "(" & rsResponsavel("TELEFONE1") & ")"
    Else
        NotaPromissoria.ParameterFields(9).AddCurrentValue rsEmpCli("RAZAO_SOCIAL") & "(" & rsResponsavel("TELEFONE1") & ")"
    End If
    'Recupera endereco da empresaCliente
    Set rsEndEmpCli = oControle.oEndereco.recuperarEndereco(oControle.oEndereco.consultaIdObjectEntidade("empresaCliente"), rsFinanciamentoEmpresa("ID_EMPRESACLIENTE"))
    oControle.oBairro.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oControle.oBairro.mTIMEOUT = gstrTimeOutGeral
    oControle.oMunicipio.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oControle.oMunicipio.mTIMEOUT = gstrTimeOutGeral
    NotaPromissoria.ParameterFields(10).AddCurrentValue rsEndEmpCli("TIPO_LOGRADOURO") & " " & rsEndEmpCli("LOGRADOURO") & " Nº" & rsEndEmpCli("NUMERO") & " - " & oControle.oBairro.RecuperaNomeBairro(rsEndEmpCli("ID_BAIRRO")) & " - " & oControle.oMunicipio.RecuperaNomeMunicipio(rsEndEmpCli("ID_MUNICIPIO"))
    
    oControle.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oControle.oEstado.mTIMEOUT = gstrTimeOutGeral
    oControle.oMunicipio.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oControle.oMunicipio.mTIMEOUT = gstrTimeOutGeral
    oControle.oBairro.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    oControle.oBairro.mTIMEOUT = gstrTimeOutGeral
    
    NotaPromissoria.ParameterFields(6).AddCurrentValue rsEndResFin("TIPO_LOGRADOURO") & " " _
                                        & rsEndResFin("LOGRADOURO") & ", " _
                                        & rsEndResFin("NUMERO") & Chr(13) _
                                        & oControle.oBairro.RecuperaNomeBairro(rsEndResFin("ID_BAIRRO")) & " - " _
                                        & oControle.oMunicipio.RecuperaNomeMunicipio(rsEndResFin("ID_MUNICIPIO")) & "-" _
                                        & oControle.oEstado.RecuperaNomeEstado(rsEndResFin("ID_ESTADO")) & " " & _
                                        rsEndResFin("CEP")
                                        
    NotaPromissoria.ParameterFields(7).AddCurrentValue oControle.oMunicipio.RecuperaNomeMunicipio(rsEndResFin("ID_MUNICIPIO")) & ", " & Day(Now) & " de " & Format(Now, "mmmm") & " de " & Year(Now)
            
    NotaPromissoria.Database.SetDataSource rsFinancimentoDetalhado
    
    frmRelGenerico.crGenerico.ReportSource = NotaPromissoria
    frmRelGenerico.Caption = "NOTA PROMISSÓRIA - " & txtDescricaoEntidade
    txtCodigoEntidade = ""
    mstrTipoOperacao = ""
    frmRelGenerico.crGenerico.ViewReport
    frmRelGenerico.Show vbModal


End Sub
Private Sub DesabilitaCamposFinanciamento()
    txtValorFinanciado.Enabled = False
    txtParcelas.Enabled = False
    txtTaxa.Enabled = False
    DTPickerInicio.Enabled = False
    DTPickerFim.Enabled = False
End Sub
Private Sub HabilitaCamposFinanciamento()
    txtValorFinanciado.Enabled = True
    txtParcelas.Enabled = True
    txtTaxa.Enabled = True
    DTPickerInicio.Enabled = True
End Sub

Private Sub txtJurosMora_GotFocus()
    txtJurosMora.SelStart = 0
    txtJurosMora.SelLength = Len(txtJurosMora)
End Sub

Private Sub txtJurosMora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub txtParcelas_Change()
    Call CalcularDataFim
    Call CalcularFinanciamento
End Sub

Private Sub UpDown1_DownClick()
DTPickerInicio = DTPickerInicio + 1
End Sub
Private Sub CalcularDataFim()

Dim i As Integer
Dim setouDataPrimeira As Boolean

PrepararVencimentos
DTPickerInicio = DTPickerInicio + 1
'Testa se a data da 1ªParcela cairá um final de semana ou feriado, sempre que for acrescenta + 1dia
oControle.oFeriado.mTIMEOUT = gstrTimeOutGeral
oControle.oFeriado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
Do While (FinalDeSemana(Format(DTPickerInicio, "dd/mm/yyyy")) Or oControle.oFeriado.DataFeriado(Format(DTPickerInicio, "dd/mm/yyyy")))
   DTPickerInicio = DTPickerInicio + 1
Loop

'Seta a última parcela para o dia inicio
DTPickerFim = DTPickerInicio

If Not IsNumeric(txtParcelas) Then txtParcelas = 1: txtParcelas.SelStart = 1
If txtParcelas = 0 Then txtParcelas = 1: txtParcelas.SelStart = 1

For i = 1 To txtParcelas

   'Testa se a data que foi atualizada é um final de semana ou feriado, sempre que for acrescenta + 1dia
   Do While (FinalDeSemana(Format(DTPickerFim, "dd/mm/yyyy")) Or oControle.oFeriado.DataFeriado(Format(DTPickerFim, "dd/mm/yyyy")))
      'VencimentoNovaParcela = VencimentoNovaParcela + 1
      DTPickerFim = DTPickerFim + 1
   Loop
        
    'Do While FinalDeSemana(DTPickerFim) 'Testa se a data que foi atualizada é um final de semana, sempre que for acrescenta + 1dia
    '    DTPickerFim = DTPickerFim + 1
    '    DoEvents
    'Loop
    
    'Do While oControle.oFeriado.DataFeriado(DTPickerFim) 'Testa se é um feriado, sempre que for acrescenta + 1dia
    '    DTPickerFim = DTPickerFim + 1
    '    DoEvents
    'Loop
    
    GerarVencimentos i, DTPickerFim 'A data de vencimento será armazenada
    If Not setouDataPrimeira Then 'Testa se a data do 1° vencimento foi guardada, senão guarda
        DTPickerInicio = DTPickerFim 'Guarda a data do 1º vencimento
        setouDataPrimeira = True 'Registra que o 1° vencimento já foi calculado
        If txtParcelas = 1 Then
            DTPickerFim = DTPickerInicio
            Exit For
        End If
    End If
    If i = txtParcelas Then Exit For
    DTPickerFim = DTPickerFim + 1 'Acrescenta + 1dia à data do último pagamento
    DoEvents
Next

End Sub

Private Function FinalDeSemana_(Data As Date) As Boolean

    If Weekday(Data) = vbSunday Or Weekday(Data) = vbSaturday Then
        FinalDeSemana_ = True
    End If

End Function

Private Sub txtParcelas_GotFocus()
    txtParcelas.SelStart = 0
    txtParcelas.SelLength = Len(txtParcelas)
End Sub

Private Sub txtParcelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtTaxa_Change()
    Call CalcularFinanciamento
End Sub

Private Sub txtTaxa_GotFocus()
    txtTaxa.SelStart = 0
    txtTaxa.SelLength = Len(txtTaxa)
End Sub

Private Sub txtTaxa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub txtTaxa_LostFocus()
    txtTaxa = Format(txtTaxa, "0.00")
End Sub

Private Sub txtValorFinanciado_Change()
    Call CalcularFinanciamento
End Sub

Private Sub txtValorFinanciado_GotFocus()
    txtValorFinanciado.SelStart = 0
    txtValorFinanciado.SelLength = Len(txtValorFinanciado)
End Sub

Private Sub txtValorFinanciado_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{TAB}"
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 44 Then KeyAscii = 0
    
End Sub
Private Sub CalcularFinanciamento()
    
If mstrTipoOperacao <> "I" Then Exit Sub
    
    If Not IsNumeric(txtValorFinanciado) Then txtValorFinanciado = "0,00": txtValorFinanciado.SelStart = Len(txtValorFinanciado)
    If Not IsNumeric(txtTaxa) Then txtTaxa = "15,00": txtTaxa.SelStart = Len(txtTaxa)
    
    If txtValorFinanciado = "0,00" Then txtTotal = "0,00": txtParcelasValor = "0,00": Exit Sub
    txtTotal = Format((txtValorFinanciado * txtTaxa / 100) + txtValorFinanciado, "0.00")
    txtParcelasValor = Format(txtTotal / txtParcelas, "0.00")
    txtTotal = Format(txtParcelasValor * txtParcelas, "0.00")

End Sub

Private Sub txtValorFinanciado_LostFocus()
    txtValorFinanciado = Format(txtValorFinanciado, "0.00")
End Sub
Private Function ValidarEmprestimo() As Boolean

If txtValorFinanciado = "" Then
    MsgBox "Valor do empréstimo incorreto", vbInformation, "VALOR INVÁLIDO."
    txtValorFinanciado.SetFocus
    Exit Function
End If

If txtValorFinanciado <= 0 Then
    MsgBox "Valor do empréstimo incorreto", vbInformation, "VALOR INVÁLIDO."
    txtValorFinanciado.SetFocus
    Exit Function
End If

If CCur(txtValorFinanciado) > CCur(txtLimite) Then
    MsgBox "Valor do empréstimo não pode ultrapassar ao limite aprovado", vbInformation, "LIMITE DE CRÉDITO ULTRAPASSADO."
    txtValorFinanciado.SetFocus: txtValorFinanciado.SelStart = 0: txtValorFinanciado.SelLength = Len(txtValorFinanciado)
    Exit Function
End If

If txtTaxa = 0 Then
    If MsgBox("Valor da taxa zerado, é isto mesmo ?", vbYesNo, "VALOR ZERO NA TAXA.") = vbNo Then
        txtTaxa.SetFocus
        Exit Function
    End If
End If

If cboRotas.ListIndex = -1 Then
    MsgBox "Selecione o agente responsável pela cobrança", vbInformation, "AGENTE NÃO INFORMADO."
    cboRotas.SetFocus
    Exit Function
End If

ValidarEmprestimo = True

End Function
