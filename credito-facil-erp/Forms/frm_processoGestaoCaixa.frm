VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF9344E-3A55-11D5-B99A-0060083D6B0C}#1.0#0"; "UCNumero.ocx"
Begin VB.Form frm_processoGestaoCaixa 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Processo Gestão de Caixa"
   ClientHeight    =   3975
   ClientLeft      =   7995
   ClientTop       =   5265
   ClientWidth     =   5085
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_processoGestaoCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   4110
      TabIndex        =   10
      Text            =   "ID"
      Top             =   690
      Width           =   915
   End
   Begin VB.Frame FraDetalhesFechamento 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalhes de fechamento"
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
      Height          =   1035
      Left            =   270
      TabIndex        =   3
      Top             =   2220
      Width           =   4545
      Begin VB.TextBox txtDataFechamento 
         Alignment       =   2  'Center
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
         Left            =   90
         MaxLength       =   10
         TabIndex        =   6
         Top             =   510
         Width           =   1725
      End
      Begin UCNumero.ctlNumero ctlNumValorFechamento 
         Height          =   375
         Left            =   3270
         TabIndex        =   14
         Top             =   510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Calibri"
         FontSize        =   12
         Mascara         =   "999999,99"
         TipoDeDado      =   1
         MaxValue        =   999999.99
         MinValue        =   -999999.99
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2970
         Picture         =   "frm_processoGestaoCaixa.frx":058A
         Top             =   600
         Width           =   240
      End
      Begin VB.Label lblSaldoFinal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldo Final"
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
         Left            =   3630
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDataFechamento 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fechamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame FraDetalhesAbertura 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalhes de abertura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   270
      TabIndex        =   2
      Top             =   780
      Width           =   4545
      Begin UCNumero.ctlNumero ctlNumValorAbertura 
         Height          =   375
         Left            =   3270
         TabIndex        =   13
         Top             =   510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Calibri"
         FontSize        =   12
         Mascara         =   "999999,99"
         TipoDeDado      =   1
         MaxValue        =   999999.99
         MinValue        =   -999999.99
      End
      Begin VB.TextBox txtDataAbertura 
         Alignment       =   2  'Center
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
         Left            =   90
         MaxLength       =   10
         TabIndex        =   4
         Top             =   510
         Width           =   1725
      End
      Begin VB.Image img_saco 
         Height          =   240
         Left            =   2970
         Picture         =   "frm_processoGestaoCaixa.frx":0B14
         Top             =   600
         Width           =   240
      End
      Begin VB.Label lblSaldoInicial 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldo Inicial"
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
         Left            =   3510
         TabIndex        =   8
         Top             =   210
         Width           =   945
      End
      Begin VB.Label lblDataAbertura 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Abertura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":109E
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":1638
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":1BD2
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":216C
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":2706
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":2CA0
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":323A
            Key             =   "abre"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":37D4
            Key             =   "fecha"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_processoGestaoCaixa.frx":3D6E
            Key             =   "Pesquisar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarCadastroFuncao 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   1164
      ButtonWidth     =   1111
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Key             =   "Abrir"
            Description     =   "Abrir"
            Object.ToolTipText     =   "Abertura de Caixa"
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechar"
            Key             =   "Fechar"
            Description     =   "Fechar"
            Object.ToolTipText     =   "Fechar o Caixa"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Description     =   "Fecha a janela atual"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frm_processoGestaoCaixa.frx":4308
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   5085
      _ExtentX        =   8969
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
   Begin VB.Label lblLabel2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   270
      TabIndex        =   12
      Top             =   3210
      Width           =   4575
   End
   Begin VB.Label lblLabel1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   270
      TabIndex        =   11
      Top             =   1770
      Width           =   4545
   End
End
Attribute VB_Name = "frm_processoGestaoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oCaixa As New clsCaixa
Public mstrTipoOperacao As String

Private Sub Form_Activate()

Dim rsCaixa As ADODB.Recordset
    
    Set rsCaixa = oCaixa.consulta(oCaixa.IdUltimoCaixa)
    
    If Not rsCaixa.EOF Then
        MoveObjTela rsCaixa
    End If


End Sub

Private Sub Form_Load()

    oCaixa.mTIMEOUT = gstrTimeOutGeral
    oCaixa.mSTRING_CONEXAO = gstrConexaoCreditoFacil
    ToolbarCadastroFuncao.Buttons("Abrir").Enabled = False
    ToolbarCadastroFuncao.Buttons("Fechar").Enabled = False
    
    If oCaixa.IdUltimoCaixa = 0 Then
        ToolbarCadastroFuncao.Buttons("Abrir").Enabled = True
        'ctlNumValorAbertura.Enabled = True
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrTipoOperacao = ""
    Set oCaixa = Nothing
End Sub

Private Sub ToolbarCadastroFuncao_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "Abrir"
        Click_Abrir
    Case "Fechar"
        Click_Fechar
    Case "Sair"
        Click_Sair

End Select

End Sub

Private Sub MoveObjTela(ByVal rsCaixa As ADODB.Recordset)

    'Se não exisir um caixa aberto
    If rsCaixa.EOF Then
    
        ToolbarCadastroFuncao.Buttons("Abrir").Enabled = True
        ToolbarCadastroFuncao.Buttons("Fechar").Enabled = False
        stbmsg.SimpleText = "Faça a abertura de caixa"
    
    End If
    
    'Sobre o movimento de caixa
    txtID = rsCaixa("ID_CAIXA")
    'Dados Sobre o Caixa: Aberto
    txtDataAbertura = Format(rsCaixa("DATA_ABERTURA"), "dd/mm/yyyy")
    ctlNumValorAbertura.Texto = Format(rsCaixa("SALDO_ABERTURA"), "0.00")
    lblLabel1 = "Aberto por: " & rsCaixa("USUARIO_ABERTURA")
    
    'Caixa está aberto
    If IsNull(rsCaixa("DATA_FECHAMENTO")) Then
        
        'Se o caixa estiver aberto, então mostrar no campo saldo de fechamento o saldo atual do caixa
        ctlNumValorFechamento.Texto = Format(oCaixa.getSaldo(txtID), "0.00")
        txtDataFechamento = ""
        
        ToolbarCadastroFuncao.Buttons("Fechar").Enabled = True
        ToolbarCadastroFuncao.Buttons("Abrir").Enabled = False
        stbmsg.SimpleText = "O caixa está aberto"
        
    'Caixa está fechado
    Else
        'Dados Sobre o Caixa: Fechado
        
        'Se o caixa estiver fechado, então mostrar o valor com o qual foi fechado
        ctlNumValorFechamento.Texto = Format(rsCaixa("SALDO_FECHAMENTO"), "0.00")
        txtDataFechamento = Format(rsCaixa("DATA_FECHAMENTO"), "dd/mm/yyyy")
        lblLabel2 = "Fechado por: " & rsCaixa("USUARIO_FECHAMENTO")
        
        ToolbarCadastroFuncao.Buttons("Abrir").Enabled = True
        ToolbarCadastroFuncao.Buttons("Fechar").Enabled = False
        stbmsg.SimpleText = "O caixa está fechado"
        
        
    End If

End Sub
Private Sub MoveTelaObj()

    Select Case mstrTipoOperacao
    
        Case "A" '- Abrir
            oCaixa.m_04_SALDO_ABERTURA = ctlNumValorAbertura.Texto
            oCaixa.m_06_USUARIO_ABERTURA = LogInUserID
            txtID = oCaixa.abrirCaixa
        Case "F" '- fechar
            oCaixa.m_01_ID_CAIXA = txtID
            oCaixa.m_05_SALDO_FECHAMENTO = ctlNumValorFechamento.Texto
            oCaixa.m_07_USUARIO_FECHAMENTO = LogInUserID
            txtID = oCaixa.fecharCaixa
    End Select

End Sub
Private Sub Click_Abrir()
     
    
    If mstrTipoOperacao = "" Then
        txtDataAbertura = Format(Now(), "dd/mm/yyyy")
        txtDataFechamento = ""
        ctlNumValorAbertura.Texto = oCaixa.getSaldo(oCaixa.IdUltimoCaixa)
        ctlNumValorFechamento.Texto = ctlNumValorAbertura.Texto
        stbmsg.SimpleText = "Abrindo o caixa"
        mstrTipoOperacao = "A"
        Exit Sub
    ElseIf mstrTipoOperacao = "A" Then
        If Not ValidaAbertura Then Exit Sub
        If MsgBox("Confirma a abertura do caixa?", vbYesNo) = vbNo Then
             Exit Sub
        End If
    Else
        Exit Sub
    End If

   mstrTipoOperacao = "A"
   MoveTelaObj
   mstrTipoOperacao = ""
   MsgBox "O Caixa foi aberto com sucesso!"
   Form_Activate

End Sub
Private Sub Click_Fechar()
    
    If Not ValidaFechamento Then Exit Sub
    
    If MsgBox("Confirma o fechamento do caixa?", vbYesNo) = vbNo Then
        Exit Sub
    End If
   
   mstrTipoOperacao = "F"
   MoveTelaObj
   MsgBox "O Caixa foi fechado com sucesso!"
   mstrTipoOperacao = ""
   Form_Activate
   

End Sub
Private Function ValidaFechamento() As Boolean

    'Valida
    If ctlNumValorFechamento.Texto = "" Then
      MsgBox "O valor do saldo final não foi informado"
      ctlNumValorFechamento.SetFocus
      Exit Function
    End If
    
    If ctlNumValorFechamento.Texto = 0 Then
      If MsgBox("O valor do saldo final informado foi zero, confirma?", vbYesNo) = vbNo Then
        ctlNumValorFechamento.SetFocus
        Exit Function
      End If
    End If
    
    If ctlNumValorFechamento.Texto < 0 Then
      If MsgBox("O valor do saldo final informado foi negativo, confirma?", vbYesNo) = vbNo Then
        ctlNumValorFechamento.SetFocus
        Exit Function
      End If
    End If

    ValidaFechamento = True
 
End Function
Private Function ValidaAbertura() As Boolean

    'Valida
    If ctlNumValorAbertura.Texto = "" Then
        MsgBox "O valor do saldo inicial não foi informado"
        ctlNumValorAbertura.SetFocus
        Exit Function
    End If
    
    If ctlNumValorAbertura.Texto = 0 Then
      If MsgBox("O valor do saldo inicial informado foi zero, confirma?", vbYesNo) = vbNo Then
        ctlNumValorAbertura.SetFocus
        Exit Function
      End If
    End If
    
    If ctlNumValorAbertura.Texto < 0 Then
      If MsgBox("O valor do saldo inicial informado foi negativo, confirma?", vbYesNo) = vbNo Then
        ctlNumValorAbertura.SetFocus
        Exit Function
      End If
    End If

    ValidaAbertura = True
 
End Function

Private Sub Click_Sair()
    Unload Me
End Sub
