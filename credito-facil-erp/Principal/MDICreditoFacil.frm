VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDICreditoFacil 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Crédito Fácil - ERP - Versão "
   ClientHeight    =   7395
   ClientLeft      =   1980
   ClientTop       =   1035
   ClientWidth     =   6585
   Icon            =   "MDICreditoFacil.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   2145
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   6525
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6585
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   16005
         Left            =   240
         Picture         =   "MDICreditoFacil.frx":058A
         ScaleHeight     =   16005
         ScaleWidth      =   24000
         TabIndex        =   3
         Top             =   120
         Width           =   24000
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2040
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   767
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuOpt_Sair 
      Caption         =   "Sair"
   End
   Begin VB.Menu menuSeguranca 
      Caption         =   "Segurança"
      Begin VB.Menu menuOpt_Funcao 
         Caption         =   "Função"
      End
      Begin VB.Menu menuOpt_Usuario 
         Caption         =   "Usuários"
      End
      Begin VB.Menu menuOpt_Permissoes 
         Caption         =   "Permissões"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu menuOpt_Empresa 
         Caption         =   "Empresa Mestre"
      End
      Begin VB.Menu menuOpt_Bairros 
         Caption         =   "Bairros"
      End
      Begin VB.Menu menuOpt_Feriados 
         Caption         =   "Feriados"
      End
      Begin VB.Menu menuOpt_Rotas 
         Caption         =   "Rotas"
      End
      Begin VB.Menu menuOpt_Ramo 
         Caption         =   "Ramos de Atividades Comerciais"
      End
      Begin VB.Menu menuOpt_Clientes 
         Caption         =   "Empresas Clientes"
      End
   End
   Begin VB.Menu menuAdministrativo 
      Caption         =   "Administrativo Pessoal"
      Begin VB.Menu menuOpt_Funcionarios 
         Caption         =   "Funcionários"
      End
   End
   Begin VB.Menu menuAdminFinanceiro 
      Caption         =   "Administrativo Financeiro"
      Begin VB.Menu menuOpt_Caixa 
         Caption         =   "Caixa"
      End
      Begin VB.Menu menuOpt_Cobranca 
         Caption         =   "Cobrança"
      End
      Begin VB.Menu menuOpt_Baixa 
         Caption         =   "Baixa"
      End
      Begin VB.Menu menuOpt_BaixaVencimentoRota 
         Caption         =   "Baixa Por Agente / Vencimento"
      End
      Begin VB.Menu menuOpt_Lancamentos 
         Caption         =   "Lançamentos"
      End
   End
   Begin VB.Menu menuGestaoEmprestimos 
      Caption         =   "Gestão de Empréstimos"
      Begin VB.Menu menuOpt_PreAprovacao 
         Caption         =   "Linhas de Crédito - Pré-Aprovação"
      End
      Begin VB.Menu menuOpt_Emprestimos 
         Caption         =   "Concessão de Empréstimos"
      End
      Begin VB.Menu Separador2 
         Caption         =   "-"
      End
      Begin VB.Menu menuOpt_Auditoria 
         Caption         =   "Auditar Financiamentos"
      End
   End
   Begin VB.Menu menuRelatorios 
      Caption         =   "Relatórios"
      Begin VB.Menu menuOpt_ResumoCaixaAtual 
         Caption         =   "Resumo do Caixa Atualmente Aberto"
      End
      Begin VB.Menu menuOpt_ResumoCaixaFechado 
         Caption         =   "Resumo do Caixa Fechado / Por Data"
      End
      Begin VB.Menu mnuOpt_LancamentosFuturos 
         Caption         =   "Lançamentos Futuros"
      End
      Begin VB.Menu Separador 
         Caption         =   "-"
      End
      Begin VB.Menu menuOpt_FormularioCadastralCliente 
         Caption         =   "Formulário Cadastral de Cliente"
      End
      Begin VB.Menu mnuRelatorioListagemEmpresas 
         Caption         =   "Lista de Empresas Cadastradas"
      End
   End
End
Attribute VB_Name = "MDICreditoFacil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oControle As New ControladorCreditoFacil
Private rsEmpresaMestre As ADODB.Recordset
Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & App.Major & "." & App.Minor & "." & App.Revision
    
    'Recupere a logo da empresa
    oControle.oEmpresa.m_stringConexao = gstrConexaoCreditoFacil
    oControle.oEmpresa.m_timeOut = gstrTimeOutGeral
    Set rsEmpresaMestre = oControle.oEmpresa.consulta(gstrEmpresaMestre)
    
    If gstrEmpresaMestre > 0 Then gstrLogoRel = oControle.oEmpresa.carregarImagem(rsEmpresaMestre)
    
End Sub

Private Sub MDIForm_Resize()

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
        
    ' Set the MDI form's picture.
    Picture = picStretched.Image


End Sub

Private Sub menuOpt_Auditoria_Click()

Dim relAuditoria As rel_auditoriaFinanciamentosAtuais
Dim rsAuditoria As ADODB.Recordset
Dim oFinanc As New clsFinanciamento

'Recupera resumo para auditoria atual
oFinanc.m_stringConexao = gstrConexaoCreditoFacil
oFinanc.m_timeOut = gstrTimeOutGeral
Set rsAuditoria = oFinanc.ResumoAuditoriaFinanciamentosAtuais

Set relAuditoria = New rel_auditoriaFinanciamentosAtuais

    relAuditoria.Database.SetDataSource rsAuditoria
    
    frmRelGenerico.crGenerico.ReportSource = relAuditoria
    frmRelGenerico.Caption = "AUDITORIA FINANCIAMENTOS ATUAIS"
    frmRelGenerico.crGenerico.ViewReport
    frmRelGenerico.Show vbModal

End Sub

Private Sub menuOpt_Bairros_Click()
frm_crudBairro.Show 'vbModal
End Sub

Private Sub menuOpt_Baixa_Click()
frm_processoBaixa.Show vbModal
End Sub

Private Sub menuOpt_BaixaVencimentoRota_Click()
frm_processoBaixaDiariaPorRota.Show vbModal
End Sub

Private Sub menuOpt_Caixa_Click()
frm_processoGestaoCaixa.Show vbModal
End Sub

Private Sub menuOpt_Clientes_Click()
frm_crudEmpresaCliente.Show vbModal
End Sub

Private Sub menuOpt_Cobranca_Click()
frm_relatorioCobranca.Show vbModal
End Sub

Private Sub menuOpt_Empresa_Click()
frm_crudEmpresa.Show vbModal
End Sub

Private Sub menuOpt_Emprestimos_Click()
frm_processoGerarFinanciamento.Show vbModal
End Sub

Private Sub menuOpt_Feriados_Click()
frm_crudFeriado.Show vbModal
End Sub

Private Sub menuOpt_FormularioCadastralCliente_Click()
Dim relFicha As rel_formularioCadastral

Set relFicha = New rel_formularioCadastral

    frmRelGenerico.crGenerico.ReportSource = relFicha
    frmRelGenerico.Caption = "FICHA CADASTRAL DE CLIENTES"
    frmRelGenerico.crGenerico.ViewReport
    frmRelGenerico.Show vbModal

End Sub

Private Sub menuOpt_Funcao_Click()
frm_crudFuncao.Show vbModal
End Sub

Private Sub menuOpt_Funcionarios_Click()
frm_crudFuncionario.Show vbModal
End Sub

Private Sub menuOpt_Lancamentos_Click()
frm_processoLancamentos.Show vbModal
End Sub

Private Sub menuOpt_PreAprovacao_Click()
frm_processoLinhaCredito.Show vbModal
End Sub

Private Sub menuOpt_Ramo_Click()
frm_crudRamoAtividade.Show vbModal
End Sub

Private Sub menuOpt_ResumoCaixaAtual_Click()

Dim relResumo As rel_resumoCaixaAtual
Dim rs As ADODB.Recordset
Dim ocaixa As clsCaixa

Set ocaixa = New clsCaixa
Set relResumo = New rel_resumoCaixaAtual
ocaixa.mTIMEOUT = gstrTimeOutGeral
ocaixa.mSTRING_CONEXAO = gstrConexaoCreditoFacil

If Not ocaixa.CaixaAberto Then
    MsgBox "Não há dados a exibir, pois o caixa está fechado"
    Exit Sub
End If

Set rs = ocaixa.consultaResumoCaixaAtual()

If rs.EOF Then
    MsgBox "Não há dados a exibir, pois não há nenhum lançamento dentro do caixa atual"
    Exit Sub
End If

relResumo.Database.SetDataSource rs
relResumo.ParameterFields(1).AddCurrentValue "RELATÓRIO DE RESUMO DO CAIXA ATUAL (aberto)"
frmRelGenerico.crGenerico.ReportSource = relResumo
frmRelGenerico.Caption = "RESUMO DO CAIXA ATUAL"
frmRelGenerico.crGenerico.ViewReport
frmRelGenerico.Show vbModal

End Sub

Private Sub menuOpt_ResumoCaixaFechado_Click()
frm_relatorioResumoCaixaFechado.Show vbModal
End Sub

Private Sub menuOpt_Rotas_Click()
frm_crudRotas.Show vbModal
End Sub

Private Sub menuOpt_Sair_Click()
If MsgBox("O sistema será encerrado, confirma?", vbYesNo, "CONFIRMAÇÃO DE SAÍDA") = vbYes Then
  End
End If
End Sub

Private Sub menuOpt_Usuario_Click()
frm_crudUsuario.Show vbModal
End Sub

Private Sub Centraliza(Parent As Form, Child As Form)
Dim iTop As Integer
Dim iLeft As Integer
If Parent.WindowState <> 0 Then Exit Sub
  iTop = ((Parent.Height - Child.Height) \ 2)
  iLeft = ((Parent.Width - Child.Width) \ 2)
  Child.Move iLeft, iTop
End Sub

Private Sub mnuOpt_LancamentosFuturos_Click()
frm_relatorioLancamentosFuturos.Show vbModal
End Sub

Private Sub mnuRelatorioListagemEmpresas_Click()

Dim relListagem As rel_listagemEmpresasClientes
Dim rsRelacao As ADODB.Recordset
Dim oEmpresa As New clsEMPRESA

'Recupera resumo para auditoria atual
oEmpresa.m_stringConexao = gstrConexaoCreditoFacil
oEmpresa.m_timeOut = gstrTimeOutGeral
Set rsRelacao = oEmpresa.listagemEmpresasParaRelatorio

Set relListagem = New rel_listagemEmpresasClientes

    relListagem.Database.SetDataSource rsRelacao
    
    frmRelGenerico.crGenerico.ReportSource = relListagem
    frmRelGenerico.Caption = "RELATÓRIO LISTAGEM DE EMPRESAS CLIENTES"
    frmRelGenerico.crGenerico.ViewReport
    frmRelGenerico.Show vbModal
End Sub

