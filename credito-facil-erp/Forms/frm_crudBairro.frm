VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_crudBairro 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Bairros"
   ClientHeight    =   2805
   ClientLeft      =   8715
   ClientTop       =   6675
   ClientWidth     =   6330
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_crudBairro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   6330
   StartUpPosition =   1  'CenterOwner
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
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2100
      Width           =   975
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
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2100
      Width           =   3525
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1155
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   6285
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   375
         Left            =   5610
         TabIndex        =   9
         TabStop         =   0   'False
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
         TabIndex        =   6
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
         TabIndex        =   2
         Top             =   570
         Width           =   4395
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   285
         Left            =   5310
         TabIndex        =   5
         Top             =   120
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
         TabIndex        =   8
         Top             =   330
         Width           =   885
      End
      Begin VB.Label lblDescricao 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descricao"
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
         Width           =   915
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5610
      Top             =   1920
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
            Picture         =   "frm_crudBairro.frx":058A
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudBairro.frx":0B24
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudBairro.frx":10BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudBairro.frx":1658
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudBairro.frx":1BF2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudBairro.frx":218C
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroBairro 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
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
      MouseIcon       =   "frm_crudBairro.frx":2726
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2550
      Width           =   6330
      _ExtentX        =   11165
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
      Height          =   195
      Left            =   210
      TabIndex        =   12
      Top             =   1860
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
      Height          =   195
      Left            =   1230
      TabIndex        =   11
      Top             =   1860
      Width           =   945
   End
End
Attribute VB_Name = "frm_crudBairro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCreditoFacil       As New ControladorCreditoFacil
Private mstrTipoOperacao    As String
Private mCtLock             As Long

Private Sub cboEstado_Click()
  PopulaComboCidade
End Sub

Private Sub cmdSelecaoEntidade_Click()

    'Verifica o objeto atualmente com o foco
    If Screen.ActiveControl.Name = "cmdSelecaoEntidade" Then
    
        If cboEstado.ListIndex = -1 Or cboCidade.ListIndex = -1 Then Exit Sub
        
        Set frmPesquisa.rsResultset = oCreditoFacil.oBairro.RecuperarBairros(gstrConexaoCreditoFacil, CLng(gstrTimeOutGeral), cboEstado.ItemData(cboEstado.ListIndex), cboCidade.ItemData(cboCidade.ListIndex))
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_BAIRRO"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "DESCRICAO_BAIRRO"
        frmPesquisa.Caption = frmPesquisa.Caption & " Bairros Cadastrados"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtCodigoEntidade_LostFocus
    End If


End Sub

Private Sub Form_Load()

oCreditoFacil.oBairro.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oBairro.mSTRING_CONEXAO = gstrConexaoCreditoFacil

oCreditoFacil.oEstado.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oEstado.mSTRING_CONEXAO = gstrConexaoCreditoFacil

mstrTipoOperacao = ""

ToolbarCadastroBairro.Buttons("Salvar").Enabled = False
ToolbarCadastroBairro.Buttons("Excluir").Enabled = False

'Popula estados da federação no combo
PopulaComboEstado

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
ToolbarCadastroBairro.Buttons("Salvar").Enabled = False

End Sub
Private Sub Novo_Click()

txtDescricaoEntidade.SetFocus
txtCodigoEntidade.Text = oCreditoFacil.oBairro.GetNovoIDBairro
ToolbarCadastroBairro.Buttons("Salvar").Enabled = True
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
     MsgBox "A descrição do bairro é requerida.", vbInformation, "Mensagem"
     txtDescricaoEntidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If cboEstado.ListIndex = -1 Then
     MsgBox "Informe o estado.", vbInformation, "Mensagem"
     cboEstado.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  If cboCidade.ListIndex = -1 Then
     MsgBox "Informe a cidade.", vbInformation, "Mensagem"
     cboCidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  
  ValidarInsert = True
  
End Function
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
On Error GoTo trataerro
    
    'Atributos
    oCreditoFacil.oBairro.m_01_ID_BAIRRO = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
    oCreditoFacil.oBairro.m_02_DESCRICAO_BAIRRO = txtDescricaoEntidade.Text
    oCreditoFacil.oBairro.m_03_ID_MUNICIPIO = cboCidade.ItemData(cboCidade.ListIndex)
    oCreditoFacil.oBairro.m_04_ID_ESTADO = cboEstado.ItemData(cboEstado.ListIndex)
    If txtID.Text = "" Then
      oCreditoFacil.oBairro.m_05_USUARIO_INCLUSAO = LogInUserID
      oCreditoFacil.oBairro.m_06_DATA_INCLUSAO = Now
    End If
    oCreditoFacil.oBairro.m_07_USUARIO_ALTERACAO = LogInUserID
    oCreditoFacil.oBairro.m_08_DATA_ALTERACAO = Now
    oCreditoFacil.oBairro.m_09_CT_LOCK = mCtLock
        
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oBairro.crudInsert
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oBairro.crudUpdate
    Else
        txtID.Text = oCreditoFacil.oBairro.crudDelete
    End If
    
trataerro:
    If InStr(1, Err.Description, "FK_endereco_bairro") > 0 And Err.Number = -2147221503 Then
        MsgBox "Não é possível excluir este bairro, pois está sendo usado em algum endereço.", vbInformation, "NÃO FOI POSSÍVEL EXCLUIR"
        txtID = 0
    End If
    If InStr(1, Err.Description, "FK_rota_det_bairro") > 0 And Err.Number = -2147221503 Then
        MsgBox "Não é possível excluir este bairro, pois está sendo usado em alguma rota.", vbInformation, "NÃO FOI POSSÍVEL EXCLUIR"
        txtID = 0
    End If
    
End Sub
Private Sub Limpacampos()

txtDescricaoEntidade.Text = ""
cboCidade.ListIndex = -1
cboEstado.ListIndex = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoFacil = Nothing
End Sub

Private Sub ToolbarCadastroBairro_ButtonClick(ByVal Button As MSComctlLib.Button)

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

Private Sub txtCodigoEntidade_Change()

Limpacampos
ToolbarCadastroBairro.Buttons("Excluir").Enabled = False
ToolbarCadastroBairro.Buttons("Salvar").Enabled = False
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
        
    Set rs = oCreditoFacil.oBairro.consulta(txtCodigoEntidade)
    
    If rs.EOF Then
      'ToolbarCadastroBairro.Buttons("Excluir").Enabled = False
      'FraCampos.Visible = False
      If oCreditoFacil.oBairro.GetNovoIDBairro <> txtCodigoEntidade Then
        'txtCodigoEntidade = oCreditoFacil.oBairro.GetNovoIDFuncao
        ToolbarCadastroBairro.Buttons("Salvar").Enabled = False
      End If
      Exit Sub
    End If
    
    ToolbarCadastroBairro.Buttons("Excluir").Enabled = True
    ToolbarCadastroBairro.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rs
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rs As ADODB.Recordset)
    
    Dim i As Integer
    
    txtID.Text = rs("ID_BAIRRO")
    mCtLock = rs("CT_LOCK")
    
    txtDescricaoEntidade = rs("DESCRICAO_BAIRRO")
    
    For i = 1 To cboEstado.ListCount
      If cboEstado.ItemData(i - 1) = rs("ID_ESTADO") Then
        cboEstado.ListIndex = i - 1
      End If
    Next
    
    For i = 1 To cboCidade.ListCount
      If cboCidade.ItemData(i - 1) = rs("ID_MUNICIPIO") Then
        cboCidade.ListIndex = i - 1
      End If
    Next
    
    
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
