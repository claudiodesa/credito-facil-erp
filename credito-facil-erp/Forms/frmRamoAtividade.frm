VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_crudRamoAtividade 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cadastro: Ramos de Atividades Comerciais"
   ClientHeight    =   2325
   ClientLeft      =   4035
   ClientTop       =   4620
   ClientWidth     =   6600
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmRamoAtividade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelecaoEntidade 
      Caption         =   "[...]"
      Height          =   375
      Left            =   5820
      TabIndex        =   5
      Top             =   1380
      Width           =   435
   End
   Begin VB.Frame FraCamposChave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atividade ou Ramo de Negócio"
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
         TabIndex        =   8
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
         Top             =   300
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
         Top             =   300
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
            Picture         =   "frmRamoAtividade.frx":058A
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRamoAtividade.frx":0B24
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRamoAtividade.frx":10BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRamoAtividade.frx":1658
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRamoAtividade.frx":1BF2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRamoAtividade.frx":218C
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroRamo 
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
      MouseIcon       =   "frmRamoAtividade.frx":2726
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2070
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
Attribute VB_Name = "frm_crudRamoAtividade"
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
    
        Set frmPesquisa.rsResultset = oCreditoFacil.oRamo.RecuperarRamos()
    
        'Campo chave
        frmPesquisa.FieldsKey = "ID_RAMO"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "DESCRICAO"
        frmPesquisa.Caption = frmPesquisa.Caption & " Ramos de Atividade Cadastrados"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        txtCodigoEntidade = frmPesquisa.FieldsReturn
        txtCodigoEntidade_LostFocus
    
    End If


End Sub

Private Sub Form_Load()

oCreditoFacil.oRamo.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oRamo.mSTRING_CONEXAO = gstrConexaoCreditoFacil
mstrTipoOperacao = ""

ToolbarCadastroRamo.Buttons("Salvar").Enabled = False
ToolbarCadastroRamo.Buttons("Excluir").Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoFacil = Nothing
End Sub

Private Sub ToolbarCadastroRamo_ButtonClick(ByVal Button As MSComctlLib.Button)

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
ToolbarCadastroRamo.Buttons("Salvar").Enabled = False

End Sub
Private Sub Novo_Click()

txtDescricaoEntidade.SetFocus
txtCodigoEntidade.Text = oCreditoFacil.oRamo.GetNovoIDRamo
ToolbarCadastroRamo.Buttons("Salvar").Enabled = True
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
     MsgBox "A descrição do ramo de atividade é requerida.", vbInformation, "Mensagem"
     txtDescricaoEntidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  ValidarInsert = True
  
End Function
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)

On Error GoTo trataerro
    
    'Atributos
    oCreditoFacil.oRamo.m_01_ID_RAMO = IIf(Trim(Len(txtID.Text)) = "", 0, txtCodigoEntidade.Text)
    oCreditoFacil.oRamo.m_02_DESCRICAO = txtDescricaoEntidade.Text
        
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oRamo.crudInsert
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oRamo.crudUpdate
    Else
        txtID.Text = oCreditoFacil.oRamo.crudDelete
    End If
    
trataerro:
If InStr(1, Err.Description, "FK_empresaCliente_ramoAtividade") > 0 And Err.Number = -2147221503 Then
    MsgBox "Exclusão não permitida, pois já existem empresas cadastradas neste ramo de atividade", vbInformation, "NÃO FOI POSSÍVEL EXCLUIR"
    txtID.Text = 0
End If
    
End Sub
Private Sub Limpacampos()

txtDescricaoEntidade.Text = ""

End Sub

Private Sub txtCodigoEntidade_Change()

Limpacampos
ToolbarCadastroRamo.Buttons("Excluir").Enabled = False
ToolbarCadastroRamo.Buttons("Salvar").Enabled = False
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
        
    Set rs = oCreditoFacil.oRamo.Consulta_By_Codigo(txtCodigoEntidade, gstrConexaoCreditoFacil, gstrTimeOutGeral)
    
    If rs.EOF Then
      If oCreditoFacil.oRamo.GetNovoIDRamo <> txtCodigoEntidade Then
        ToolbarCadastroRamo.Buttons("Salvar").Enabled = False
      End If
      Exit Sub
    End If
        
    ToolbarCadastroRamo.Buttons("Excluir").Enabled = True
    ToolbarCadastroRamo.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rs
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents

End Sub
Private Sub MoveObjetoParaTelaCab(ByRef rs As ADODB.Recordset)
    
    txtID.Text = rs("ID_RAMO")
    
    txtDescricaoEntidade = rs("DESCRICAO")
    
End Sub
