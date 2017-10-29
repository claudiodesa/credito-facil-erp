VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login do usuário"
   ClientHeight    =   4215
   ClientLeft      =   5295
   ClientTop       =   5745
   ClientWidth     =   6750
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490.361
   ScaleMode       =   0  'User
   ScaleWidth      =   6337.885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   1455
      Width           =   2355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      TabIndex        =   3
      Top             =   2850
      Width           =   2340
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   -150
      TabIndex        =   4
      Top             =   2610
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "="
      TabIndex        =   2
      Top             =   2175
      Width           =   2355
   End
   Begin VB.Label lblCreditoFacilCima 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito Fácil"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   30
      Width           =   6570
   End
   Begin VB.Label lblEmpresa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito Fácil"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   570
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   60
      Width           =   6570
   End
   Begin VB.Label lblVersao 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "<<Versão>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4350
      TabIndex        =   8
      Top             =   570
      Width           =   2340
   End
   Begin VB.Label lblSenha 
      BackColor       =   &H00808080&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   1950
      Width           =   2340
   End
   Begin VB.Label lblMensagem 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   300
      TabIndex        =   6
      Top             =   3780
      Width           =   6765
   End
   Begin VB.Label lblLabel1 
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   3660
      Width           =   6765
   End
   Begin VB.Image Image2 
      Height          =   2460
      Left            =   3960
      Picture         =   "frmLogin.frx":058A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2490
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00808080&
      Caption         =   "Usuário "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1230
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   -30
      Picture         =   "frmLogin.frx":4786
      Stretch         =   -1  'True
      Top             =   -1110
      Width           =   8475
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Private oUsuario As New ControladorCreditoFacil
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim rsUsuario As ADODB.Recordset
    Dim rsFuncionario As ADODB.Recordset
    Dim rsFuncao As ADODB.Recordset
    
    blnLoginOK = False
    blnLoginAdmin = False
    Dim strMesage As String

    strMesage = "Combinação errada de Nome de Usuário e Senha."
    
    'Caso seja o admin que estiver se logando
    If txtUserName = "ADMIN" And txtPassword = "ADMIN" Then
        blnLoginAdmin = True
        MDICreditoFacil.menuAdminFinanceiro.Enabled = False
        MDICreditoFacil.menuCadastros.Enabled = False
        MDICreditoFacil.menuRelatorios.Enabled = False
        MDICreditoFacil.menuGestaoEmprestimos.Enabled = False
    End If

    If Not blnLoginAdmin Then
        'Verificar se o usuário existe no cadastro
        Set rsUsuario = oUsuario.RecuperarUsuario(txtUserName, gstrConexaoCreditoFacil, gstrTimeOutGeral)
          
        If rsUsuario.EOF Then
            'MsgBox "Usuário digitado não corresponde a um usuário válido.", vbCritical, "Segurança"
            lblMensagem.Caption = strMesage
            txtPassword.Text = ""
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName)
            txtUserName.SetFocus
            Exit Sub
        End If
          
        'Valida a senha informada
        If CriptSenha(txtPassword) = (rsUsuario("SENHA")) Then
    
            'Valida se o login do usuário não foi desativado
            If rsUsuario("STATUS") = "D" Then
                strMesage = "Este login foi desativado pelo administrador."
            Else
                blnLoginOK = True
            End If
        End If
    End If
    
'Se o admin se logou
If blnLoginAdmin = True Then
    MDICreditoFacil.Show
    MDICreditoFacil.StatusBar1.SimpleText = "Usuário : " & "ADMIN"
    'Fecha a tela de login
    cmdCancel_Click
    Exit Sub
End If
    
'Se o login foi efetuado com sucesso
If blnLoginOK = True Then
    oUsuario.oFuncionario.m_timeOut = gstrTimeOutGeral
    oUsuario.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
    Set rsFuncionario = oUsuario.oFuncionario.consulta(rsUsuario("ID_FUNCIONARIO"))
    'a senha está correta, salvar nas variáveis globais de usuário logado
    LogInUserID = rsUsuario("LOGIN").value
    LogInUserName = rsFuncionario("NOME").value
    gstrEmpresaMestre = rsFuncionario("ID_EMPRESA")
    ' sage the username in the Registry
    SaveSetting App.EXEName, "Settings", "LastLogIn", rsUsuario("LOGIN")
    'Libera os menus de acordo com a função do usuário logado
    Set rsFuncao = oUsuario.oFuncao.Consulta_By_Codigo(rsFuncionario("ID_FUNCAO"), gstrConexaoCreditoFacil, gstrTimeOutGeral)
    If InStr(1, rsFuncao("DESCRICAO_FUNCAO"), "Gerente") > 0 Then
    Else
        MDICreditoFacil.menuGestaoEmprestimos = False
        MDICreditoFacil.menuCadastros = False
        MDICreditoFacil.menuSeguranca = False
        MDICreditoFacil.menuRelatorios = False
        MDICreditoFacil.menuAdministrativo = False
        MDICreditoFacil.menuOpt_Caixa = False
        MDICreditoFacil.menuOpt_Lancamentos = False
    End If
    Set rsUsuario = Nothing
    'Abre o MDI
    MDICreditoFacil.Show
    MDICreditoFacil.StatusBar1.SimpleText = "Usuário Logado : " & LogInUserID & " | " & rsFuncionario("NOME")
    Set rsFuncionario = Nothing
    'Fecha a tela de login
    cmdCancel_Click
Else
    txtPassword.Text = ""
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
    txtUserName.SetFocus
    lblMensagem.Caption = strMesage
End If
    
End Sub

Private Sub Form_Activate()

    'Define o último usuário logado no campo usuário
    ' Get the user that last logged in from the registry
    txtUserName = GetSetting(App.EXEName, "Settings", "LastLogIn", "")

    If Len(txtUserName) > 0 Then
        txtPassword.SetFocus
    End If

    lblVersao.Caption = "Versão " & gstrVersao

    'Provisório
    'txtPassword = "12345"
    'cmdOK_Click
    'txtUserName = "ADMIN"
    'txtPassword = "ADMIN"
    'cmdOK_Click

End Sub

Private Sub Form_Load()

    'Definindo a versão do sistema
    gstrVersao = App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub txtPassword_Change()
    lblMensagem = ""
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword = CriptSenha(txtPassword)
End Sub

Private Sub txtUserName_Change()
    lblMensagem = ""
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
End Sub

