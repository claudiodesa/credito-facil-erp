VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_crudFuncao 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Funções de Usuário"
   ClientHeight    =   3495
   ClientLeft      =   4200
   ClientTop       =   4605
   ClientWidth     =   6600
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmFuncao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
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
      Height          =   1125
      Left            =   180
      TabIndex        =   8
      Top             =   1980
      Width           =   2235
      Begin VB.CheckBox chkAtiva 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ativa"
         Height          =   285
         Left            =   810
         TabIndex        =   12
         Top             =   690
         Width           =   885
      End
      Begin VB.TextBox txtSigla 
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
         Left            =   810
         MaxLength       =   3
         TabIndex        =   9
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla"
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
         TabIndex        =   11
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Top             =   720
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdSelecaoEntidade 
      Caption         =   "[...]"
      Height          =   375
      Left            =   5850
      TabIndex        =   5
      Top             =   1380
      Width           =   435
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1155
      Left            =   180
      TabIndex        =   2
      Top             =   810
      Width           =   6285
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   5310
         TabIndex        =   13
         Text            =   "ID"
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
         TabIndex        =   4
         Top             =   570
         Width           =   4425
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
         TabIndex        =   3
         Top             =   570
         Width           =   945
      End
      Begin VB.Label lblDescricao 
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
         TabIndex        =   7
         Top             =   330
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
         TabIndex        =   6
         Top             =   330
         Width           =   885
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   3210
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
            Picture         =   "frmFuncao.frx":058A
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncao.frx":0B24
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncao.frx":10BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncao.frx":1658
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncao.frx":1BF2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncao.frx":218C
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
      Width           =   6600
      _ExtentX        =   11642
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
      MouseIcon       =   "frmFuncao.frx":2726
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   6600
      _ExtentX        =   11642
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
Attribute VB_Name = "frm_crudFuncao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCreditoFacil       As New ControladorCreditoFacil
Private mstrTipoOperacao    As String
Private mCtLock             As Long

Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    If Screen.ActiveControl.Name = "cmdSelecaoEntidade" Then
    
        Set frmPesquisa.rsResultset = oCreditoFacil.oFuncao.RecuperarFuncoes(gstrConexaoCreditoFacil, CLng(gstrTimeOutGeral))
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_FUNCAO"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "DESCRICAO_FUNCAO"
        frmPesquisa.Caption = frmPesquisa.Caption & " Funções Cadastradas"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtCodigoEntidade_LostFocus
    
    End If


End Sub

Private Sub Form_Load()

FraCampos.Visible = False
oCreditoFacil.oFuncao.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oFuncao.mSTRING_CONEXAO = gstrConexaoCreditoFacil
mstrTipoOperacao = ""

ToolbarCadastroFuncao.Buttons("Salvar").Enabled = False
ToolbarCadastroFuncao.Buttons("Excluir").Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoFacil = Nothing
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

txtDescricaoEntidade.SetFocus
txtCodigoEntidade.Text = oCreditoFacil.oFuncao.GetNovoIDFuncao
ToolbarCadastroFuncao.Buttons("Salvar").Enabled = True
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
     MsgBox "A descrição da função é requerida.", vbInformation, "Mensagem"
     txtDescricaoEntidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtSigla)) = 0 Then
     MsgBox "A sigla da função é requerida.", vbInformation, "Mensagem"
     txtSigla.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If Trim(Len(txtSigla)) < 3 Then
     MsgBox "A sigla da função deve ter no mínimo 3 caracteres.", vbInformation, "Mensagem"
     txtSigla.SetFocus
     ValidarInsert = False
     Exit Function
  End If
    
  If Len(Trim(txtSigla)) > 3 Then
     MsgBox "A sigla da função deve ter no máximo 3 caracteres.", vbInformation, "Mensagem"
     txtSigla.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If chkAtiva.value = 0 Then
     MsgBox "A função inserida não está definida para ser ativa, se necessário altere para ativá-la.", vbInformation, "Mensagem"
  End If
  
  ValidarInsert = True
  
End Function
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
On Error GoTo trataerro
    
    'Atributos
    oCreditoFacil.oFuncao.m_01_ID_FUNCAO = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
    oCreditoFacil.oFuncao.m_02_DESCRICAO_FUNCAO = txtDescricaoEntidade.Text
    oCreditoFacil.oFuncao.m_03_STATUS_FUNCAO = IIf(chkAtiva.value = 1, "A", "D")
    oCreditoFacil.oFuncao.m_04_SIGLA_FUNCAO = Trim(txtSigla.Text)
    If txtID.Text = "" Then
      oCreditoFacil.oFuncao.m_05_USUARIO_INCLUSAO = LogInUserID
      oCreditoFacil.oFuncao.m_06_DATA_INCLUSAO = Now
    End If
    oCreditoFacil.oFuncao.m_07_USUARIO_ALTERACAO = LogInUserID
    oCreditoFacil.oFuncao.m_08_DATA_ALTERACAO = Now
    oCreditoFacil.oFuncao.m_09_CT_LOCK = mCtLock
        
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oFuncao.crudInsert
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oFuncao.crudUpdate
    Else
        txtID.Text = oCreditoFacil.oFuncao.crudDelete
    End If
    
trataerro:
 If InStr(1, Err.Description, "FK_funcionario_funcao") > 0 And Err.Number = -2147221503 Then
    MsgBox "Não é possível excluir a função, quando existem funcionários que a exercem.", vbInformation, "EXCLUSÃO NÃO FOI REALIZADA"
    txtID = 0
 End If
    
End Sub
Private Sub Limpacampos()

FraCampos.Visible = True
txtDescricaoEntidade.Text = ""
txtSigla.Text = ""
chkAtiva.value = 0

End Sub

Private Sub txtCodigoEntidade_Change()

Limpacampos
ToolbarCadastroFuncao.Buttons("Excluir").Enabled = False
ToolbarCadastroFuncao.Buttons("Salvar").Enabled = False
txtID.Text = ""
stbmsg.SimpleText = ""

End Sub

Private Sub txtCodigoEntidade_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCodigoEntidade_LostFocus()

Dim rs As ADODB.Recordset

    If txtCodigoEntidade = "" Then Exit Sub
        
    Set rs = oCreditoFacil.oFuncao.Consulta_By_Codigo(txtCodigoEntidade, gstrConexaoCreditoFacil, gstrTimeOutGeral)
    
    If rs.EOF Then
      'ToolbarCadastroFuncao.Buttons("Excluir").Enabled = False
      'FraCampos.Visible = False
      If oCreditoFacil.oFuncao.GetNovoIDFuncao <> txtCodigoEntidade Then
        'txtCodigoEntidade = oCreditoFacil.oFuncao.GetNovoIDFuncao
        ToolbarCadastroFuncao.Buttons("Salvar").Enabled = False
      End If
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
    
    txtID.Text = rs("ID_FUNCAO")
    mCtLock = rs("CT_LOCK")
    
    txtDescricaoEntidade = rs("DESCRICAO_FUNCAO")
    
    If rs("STATUS_FUNCAO") = "A" Then
        chkAtiva.value = 1
    Else
        chkAtiva.value = 0
    End If
    
    txtSigla = rs("SIGLA_FUNCAO")
    
End Sub
