VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_relatorioLancamentosFuturos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Relatório Lançamentos Futuros"
   ClientHeight    =   2235
   ClientLeft      =   2610
   ClientTop       =   1755
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   1890
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd_emitir_relatorio 
      Caption         =   "Emitir Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame fra_classifica_relarorio 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ordem do relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   3300
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton opt_class_vencimento 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vencimeno"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   2535
      End
      Begin VB.OptionButton opt_class_nomeCliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ordem alfabética de cliente"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame fra_tipo_relatorio 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filtro - Tipo de Lançamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton opt_tipo_foraPrazo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fora do prazo"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton opt_tipo_dentroPrazo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dentro do prazo"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton opt_tipo_todos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frm_relatorioLancamentosFuturos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oRelatorio As New clsFinanciamento

Private Sub cmd_emitir_relatorio_Click()

Dim relLancamentosFuturos As New rel_lancamentosFuturos
Dim rs As ADODB.Recordset

oRelatorio.m_stringConexao = gstrConexaoCreditoFacil
oRelatorio.m_timeOut = gstrTimeOutGeral
Set rs = oRelatorio.RelatorioLancamentosFuturos(IIf(opt_tipo_todos, todos, IIf(opt_tipo_dentroPrazo, dentro_prazo, fora_prazo)), _
                                                IIf(opt_class_nomeCliente, clientes, vencimento))

If rs.EOF Then
    stbmsg.SimpleText = "Não há lançamentos a exibir."
Else
    relLancamentosFuturos.ParameterFields(1).AddCurrentValue IIf(opt_tipo_todos, opt_tipo_todos.Caption, IIf(opt_tipo_dentroPrazo, opt_tipo_dentroPrazo.Caption, opt_tipo_foraPrazo.Caption))
    relLancamentosFuturos.ParameterFields(2).AddCurrentValue IIf(opt_class_nomeCliente, opt_class_nomeCliente.Caption, opt_class_vencimento.Caption)
    relLancamentosFuturos.ParameterFields(3).AddCurrentValue CStr(rs.RecordCount)
    relLancamentosFuturos.Database.SetDataSource rs
    frmRelGenerico.crGenerico.ReportSource = relLancamentosFuturos
    frmRelGenerico.Caption = "RELATÓRIO LANÇAMENTOS FUTUROS"
    stbmsg.SimpleText = ""
    frmRelGenerico.crGenerico.ViewReport
    frmRelGenerico.Show vbModal
End If

End Sub
