VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_crudUsuario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Usuários"
   ClientHeight    =   4590
   ClientLeft      =   6855
   ClientTop       =   4890
   ClientWidth     =   6510
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_crudUsuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6510
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAtivo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pode fazer login"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3540
      Width           =   1725
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
      Height          =   1635
      Left            =   90
      TabIndex        =   11
      Top             =   1830
      Width           =   6285
      Begin VB.TextBox txtIDFuncionario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   5310
         TabIndex        =   16
         Text            =   "IDFuncionario"
         Top             =   150
         Width           =   915
      End
      Begin VB.TextBox txtConfirmacaoSenha 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   32
         PasswordChar    =   "="
         TabIndex        =   5
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   210
         MaxLength       =   32
         PasswordChar    =   "="
         TabIndex        =   4
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txtLogin 
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
         MaxLength       =   30
         TabIndex        =   3
         Top             =   420
         Width           =   2085
      End
      Begin VB.Label lblConfirmacaoDa 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmação da Senha"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
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
         TabIndex        =   13
         Top             =   810
         Width           =   885
      End
      Begin VB.Label lblSigla 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Top             =   180
         Width           =   2055
      End
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuário / Login"
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
      Left            =   90
      TabIndex        =   7
      Top             =   690
      Width           =   6285
      Begin VB.CommandButton cmdSelecionarEntidade 
         Caption         =   "[...]"
         Height          =   405
         Left            =   5640
         TabIndex        =   17
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
         TabIndex        =   1
         Top             =   570
         Width           =   945
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
         TabIndex        =   2
         Top             =   570
         Width           =   4425
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   5310
         TabIndex        =   8
         Text            =   "ID"
         Top             =   150
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
         TabIndex        =   10
         Top             =   330
         Width           =   885
      End
      Begin VB.Label lblDescricao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
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
         Width           =   1125
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   3600
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
            Picture         =   "frm_crudUsuario.frx":058A
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudUsuario.frx":0B24
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudUsuario.frx":10BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudUsuario.frx":1658
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudUsuario.frx":1BF2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudUsuario.frx":218C
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroFuncao 
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
      MouseIcon       =   "frm_crudUsuario.frx":2726
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4335
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
Attribute VB_Name = "frm_crudUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oUsuario            As New clsUSUARIO
Private oFuncionario        As New clsFUNCIONARIO
Private mstrTipoOperacao    As String
Private mCtLock             As Long


Private Sub cmdSelecionarEntidade_Click()

    'Verifica o objeto atualmente com o foco
    'If Screen.ActiveControl.Name = "cmdSelecionarEntidade" Then
    
        oFuncionario.m_timeOut = gstrTimeOutGeral
        oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
        Set frmPesquisa.rsResultset = oFuncionario.RecuperarFuncionarios()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_FUNCIONARIO"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "NOME"
        frmPesquisa.Caption = frmPesquisa.Caption & "Selecione um Funcionário Cadastrado"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtIDFuncionario = txtCodigoEntidade
        txtCodigoEntidade_LostFocus
    
    'End If

End Sub

Private Sub Form_Load()

oUsuario.m_timeOut = gstrTimeOutGeral
oUsuario.m_stringConexao = gstrConexaoCreditoFacil
mstrTipoOperacao = ""

ToolbarCadastroFuncao.Buttons("Salvar").Enabled = False
ToolbarCadastroFuncao.Buttons("Excluir").Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set oUsuario = Nothing
Set oFuncionario = Nothing

End Sub

Private Sub ToolbarCadastroFuncao_ButtonClick(ByVal Button As MSComctlLib.Button)

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
ToolbarCadastroFuncao.Buttons("Salvar").Enabled = False

End Sub
Private Sub Novo_Click()

txtLogin.SetFocus
cmdSelecionarEntidade_Click
ToolbarCadastroFuncao.Buttons("Salvar").Enabled = True
mstrTipoOperacao = "I"
stbmsg.SimpleText = "Incluindo"
DoEvents
'txtCodigoEntidade = ""


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
    
  If Trim(Len(txtSenha)) = 0 Then
     MsgBox "O login ainda não foi definido.", vbInformation, "Mensagem"
     txtLogin.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtSenha)) = 0 Then
     MsgBox "A senha ainda não foi definida.", vbInformation, "Mensagem"
     txtSenha.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtConfirmacaoSenha)) = 0 Then
     MsgBox "A confirmação de senha não foi preenchida.", vbInformation, "Mensagem"
     txtConfirmacaoSenha.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If txtConfirmacaoSenha <> txtSenha Then
     MsgBox "As senha digitadas não coincidem.", vbInformation, "Mensagem"
     txtSenha.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If chkAtivo.value = 0 Then
    If MsgBox("O atual usuário será desativado e não poderá efetuar login, confirma?", vbYesNo, "DESATIVAÇÃO DE USUÁRIO") = vbNo Then
        chkAtivo.SetFocus
        ValidarInsert = False
        Exit Function
    End If
  End If
    
  ValidarInsert = True
  
End Function
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
    'Atributos
    oUsuario.m_01_idUsuario = txtID.Text
    oUsuario.m_02_idFuncionario = txtCodigoEntidade
    oUsuario.m_03_status = IIf(chkAtivo.value = 1, "A", "D")
    oUsuario.m_04_login = txtLogin
    oUsuario.m_05_senha = CriptSenha(txtSenha)
    If txtID.Text = "" Then
      oUsuario.m_06_usuarioInclusao = LogInUserID
      oUsuario.m_07_dataInclusao = Now
    End If
    oUsuario.m_08_usuarioAlteracao = LogInUserID
    oUsuario.m_09_dataAlteracao = Now
    oUsuario.m_10_ctLock = mCtLock
        
    If strOperacao = "I" Then
        txtID.Text = oUsuario.crudInsert
    ElseIf strOperacao = "A" Then
        txtID.Text = oUsuario.crudUpdate
    Else
        txtID.Text = oUsuario.crudDelete
    End If
    
End Sub
Private Sub Limpacampos()

txtDescricaoEntidade.Text = ""
txtLogin.Text = ""
txtSenha = ""
txtConfirmacaoSenha = ""
chkAtivo.value = 0

End Sub

Private Sub txtCodigoEntidade_Change()

Limpacampos
ToolbarCadastroFuncao.Buttons("Excluir").Enabled = False
stbmsg.SimpleText = ""
txtID = ""
txtIDFuncionario = ""

End Sub

Private Sub txtCodigoEntidade_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCodigoEntidade_LostFocus()

Dim rs As ADODB.Recordset
Dim rsFuncionario As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
    
    Set rsFuncionario = oFuncionario.consulta(txtCodigoEntidade)
    If rsFuncionario.EOF Then Exit Sub
    txtDescricaoEntidade = rsFuncionario("NOME")
    Set rsFuncionario = Nothing
            
    Set rs = oUsuario.ConsultaIdFuncionario(txtCodigoEntidade)
    
    If rs.EOF Then
      txtID = oUsuario.GetNovoIdUsuario
      'txtLogin.SetFocus
      Exit Sub
    End If
    
    FraCampos.Visible = True
    ToolbarCadastroFuncao.Buttons("Excluir").Enabled = True
    ToolbarCadastroFuncao.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rs
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rs As ADODB.Recordset)
    
    Dim rsFuncionario As ADODB.Recordset
    
    oFuncionario.m_timeOut = gstrTimeOutGeral
    oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
    Set rsFuncionario = oFuncionario.consulta(txtCodigoEntidade)
    
    txtID.Text = rs("ID_USUARIO")
    txtIDFuncionario = rsFuncionario("ID_FUNCIONARIO")
    mCtLock = rs("CT_LOCK")
    
    txtDescricaoEntidade = rsFuncionario("NOME")
    txtLogin = rs("LOGIN")
    txtSenha = DeCriptSenha(rs("SENHA"))
    txtConfirmacaoSenha = txtSenha
    
    If rs("STATUS") = "A" Then
        chkAtivo.value = 1
    Else
        chkAtivo.value = 0
    End If
    
    
End Sub

Private Sub txtSenha_Change()
txtConfirmacaoSenha = ""
End Sub
