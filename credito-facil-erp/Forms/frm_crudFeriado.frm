VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{21F8D070-6EA8-40F7-8555-9E5FA3E03CB5}#1.0#0"; "Calendario.ocx"
Begin VB.Form frm_crudFeriado 
   Caption         =   "Cadastro: Feriados"
   ClientHeight    =   3300
   ClientLeft      =   6870
   ClientTop       =   5235
   ClientWidth     =   5460
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_crudFeriado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkExcepcionalmente 
      Caption         =   "Excepcionalmente"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   2700
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FraCamposChave 
      Caption         =   "Feriado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   5205
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   5340
         TabIndex        =   9
         Text            =   "id"
         Top             =   120
         Width           =   885
      End
      Begin Calendario.ctlPicker CtlDataFeriado 
         Height          =   315
         Left            =   210
         TabIndex        =   2
         Top             =   570
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Text            =   "__/__"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatDate      =   "dd/mm"
      End
      Begin VB.CommandButton cmdSelecaoEntidade 
         Caption         =   "[...]"
         Height          =   375
         Left            =   4650
         TabIndex        =   7
         Top             =   1320
         Width           =   465
      End
      Begin VB.TextBox txtDescricaoEntidade 
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
         Left            =   210
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   4425
      End
      Begin VB.Label lblCódigo 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data (dd/mm)"
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
         Width           =   1335
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
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   1050
         Width           =   1365
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
            Picture         =   "frm_crudFeriado.frx":058A
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFeriado.frx":0B24
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFeriado.frx":10BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFeriado.frx":1658
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFeriado.frx":1BF2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_crudFeriado.frx":218C
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroFeriado 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
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
      MouseIcon       =   "frm_crudFeriado.frx":2726
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3045
      Width           =   5460
      _ExtentX        =   9631
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
Attribute VB_Name = "frm_crudFeriado"
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
    
        Set frmPesquisa.rsResultset = oCreditoFacil.oFeriado.RecuperarFeriados(gstrConexaoCreditoFacil, CLng(gstrTimeOutGeral))
    
        'Campo chave
        frmPesquisa.FieldsKey = "DATA_FERIADO"
        'Campo a ser listado no resultado da pesquisa
        frmPesquisa.FieldsList = "DESCRICAO"
        frmPesquisa.Caption = frmPesquisa.Caption & " Feriados Cadastrados"
        frmPesquisa.Show 1
        'Recebe retorno da pesquisa
        CtlDataFeriado.Text = Format(frmPesquisa.FieldsReturn, "dd/mm")
        CtlDataFeriado_LostFocus
    
    End If


End Sub

Private Sub CtlDataFeriado_Change()
    txtID = ""
    Limpacampos
    ToolbarCadastroFeriado.Buttons("Excluir").Enabled = False
    'ToolbarCadastroFeriado.Buttons("Salvar").Enabled = False
End Sub

Private Sub CtlDataFeriado_LostFocus()

Dim rs As ADODB.Recordset

    If CtlDataFeriado.Text = "__/__" Or CtlDataFeriado.Text = "" Then Exit Sub
        
    Set rs = oCreditoFacil.oFeriado.Consulta_By_Data(CtlDataFeriado.Text, gstrConexaoCreditoFacil, gstrTimeOutGeral)
    
    If rs.EOF Then
      mstrTipoOperacao = "I"
      stbmsg.SimpleText = "Incluindo"
      Exit Sub
    End If
    
    ToolbarCadastroFeriado.Buttons("Excluir").Enabled = True
    ToolbarCadastroFeriado.Buttons("Salvar").Enabled = True
    
    MoveObjetoParaTelaCab rs
    
    mstrTipoOperacao = "A"
    stbmsg.SimpleText = "Alterando"
    
    DoEvents


End Sub

Private Sub Form_Load()

oCreditoFacil.oFeriado.mTIMEOUT = gstrTimeOutGeral
oCreditoFacil.oFeriado.mSTRING_CONEXAO = gstrConexaoCreditoFacil
ToolbarCadastroFeriado.Buttons("Salvar").Enabled = False
ToolbarCadastroFeriado.Buttons("Excluir").Enabled = False
mstrTipoOperacao = ""

End Sub
Private Sub Salvar_Click()

'Validações
If CtlDataFeriado.Text = "__/__" Then Exit Sub

If Not ValidarInsert Then
    Exit Sub
End If

If mstrTipoOperacao = "A" Then
    If MsgBox("Deseja realmente alterar o registro?", vbYesNo, "Pergunta") = vbNo Then Exit Sub
End If

'Move dados para tela para o objeto
MoveTelaParaObjetoCab mstrTipoOperacao

If mstrTipoOperacao = "I" Then
    MsgBox "Registro incluido com sucesso!", vbInformation, "SUCESSO"
ElseIf mstrTipoOperacao = "A" Then
    MsgBox "Registro alterado com sucesso!", vbInformation, "SUCESSO"
End If

Call Limpacampos
CtlDataFeriado.Text = "__/__"
stbmsg.SimpleText = ""
mstrTipoOperacao = ""
ToolbarCadastroFeriado.Buttons("Salvar").Enabled = False

End Sub
Private Sub Novo_Click()

Limpacampos
CtlDataFeriado.SetFocus
ToolbarCadastroFeriado.Buttons("Salvar").Enabled = True
mstrTipoOperacao = "I"
stbmsg.SimpleText = "Incluindo"
DoEvents

End Sub
Private Sub Excluir_Click()

If CtlDataFeriado.Text = "__/__" Then Exit Sub

If MsgBox("Deseja realmente excluir o registro?", vbYesNo, "Pergunta") = vbNo Then Exit Sub

mstrTipoOperacao = "E"

MoveTelaParaObjetoCab mstrTipoOperacao

MsgBox "Registro excluido com sucesso!", vbInformation, "SUCESSO"

Call Limpacampos

CtlDataFeriado.Text = "__/__"
mstrTipoOperacao = ""
stbmsg.SimpleText = ""
ToolbarCadastroFeriado.Buttons("Salvar").Enabled = False
ToolbarCadastroFeriado.Buttons("Excluir").Enabled = False


End Sub
Private Sub Fechar_Click()
  Unload Me
End Sub
'Validar Campos do Formulario
Private Function ValidarInsert() As Boolean

Dim rsFuncao As ADODB.Recordset
    
  If Not IsDate(CtlDataFeriado.Text) Then
    ValidarInsert = False
  End If
  
  If Trim(Len(txtDescricaoEntidade)) = 0 Then
     MsgBox "A descrição do feriado é requerida.", vbInformation, "Mensagem"
     txtDescricaoEntidade.SetFocus
     ValidarInsert = False
     Exit Function
  End If
  
  ValidarInsert = True
  
End Function
Private Sub MoveTelaParaObjetoCab(ByVal strOperacao As String)
    
    'Atributos
    oCreditoFacil.oFeriado.m_01_DATA_FERIADO = CtlDataFeriado.Text
    oCreditoFacil.oFeriado.m_02_FERIADO = txtDescricaoEntidade.Text
    If chkExcepcionalmente.value = 1 Then
      oCreditoFacil.oFeriado.m_03_EXCEPCIONALMENTE = "S"
    Else
      oCreditoFacil.oFeriado.m_03_EXCEPCIONALMENTE = "N"
    End If
        
    If strOperacao = "I" Then
        txtID.Text = oCreditoFacil.oFeriado.crudInsert
    ElseIf strOperacao = "A" Then
        txtID.Text = oCreditoFacil.oFeriado.crudUpdate
    Else
        txtID.Text = oCreditoFacil.oFeriado.crudDelete
    End If
    
End Sub
Private Sub Limpacampos()

'CtlDataFeriado.Text = "__/__"
txtDescricaoEntidade.Text = ""
chkExcepcionalmente.value = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoFacil = Nothing
End Sub

Private Sub ToolbarCadastroFeriado_ButtonClick(ByVal Button As MSComctlLib.Button)

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
ToolbarCadastroFeriado.Buttons("Excluir").Enabled = False
txtID.Text = ""
stbmsg.SimpleText = ""

End Sub

Private Sub MoveObjetoParaTelaCab(ByRef rs As ADODB.Recordset)
    
    CtlDataFeriado.Text = Format(rs("DATA_FERIADO"), "dd/mm")
    txtDescricaoEntidade = rs("FERIADO")
    If rs("EXCEPCIONALMENTE") = "S" Then
      chkExcepcionalmente.value = 1
    Else
      chkExcepcionalmente.value = 0
    End If
    
    
End Sub

