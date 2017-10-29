VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_relatorioCobranca 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Demonstrativo de Cobrança"
   ClientHeight    =   1905
   ClientLeft      =   6345
   ClientTop       =   6255
   ClientWidth     =   7545
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSCommand cmdGerar 
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      Top             =   1140
      Width           =   1875
      _Version        =   65536
      _ExtentX        =   3307
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "Gerar Demonstrativo"
      ForeColor       =   -2147483630
   End
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
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker DTPickerData 
      Height          =   375
      Left            =   510
      TabIndex        =   0
      Top             =   600
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      Format          =   95354881
      CurrentDate     =   40772
   End
   Begin MSComctlLib.StatusBar stbmsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1650
      Width           =   7545
      _ExtentX        =   13309
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
   Begin VB.Label lblRota 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rota"
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
      Left            =   1950
      TabIndex        =   5
      Top             =   330
      Width           =   885
   End
   Begin VB.Label lblData 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
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
      Left            =   510
      TabIndex        =   4
      Top             =   330
      Width           =   885
   End
End
Attribute VB_Name = "frm_relatorioCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oControle As New ControladorCreditoFacil

Private Sub cmdGerar_Click()

Dim relCobranca As rel_cobranca
Dim rsCobranca As ADODB.Recordset
Dim rsEmpresaMestre As ADODB.Recordset
Dim rsFuncionarioRota As ADODB.Recordset
Dim rsRota As ADODB.Recordset
Dim x As StdPicture

'Valida
If cboRotas.ListIndex = -1 Then
    MsgBox "Selecione a rota do agente responsavel pela cobrança"
    cboRotas.SetFocus
    Exit Sub
End If

stbmsg.SimpleText = "aguarde, gerando a cobrança..."

oControle.oParcelas.m_stringConexao = gstrConexaoCreditoFacil
oControle.oParcelas.m_timeOut = gstrTimeOutGeral
Set rsCobranca = oControle.oParcelas.consultaCobranca(DTPickerData, cboRotas.ItemData(cboRotas.ListIndex))

oControle.oRota.m_stringConexao = gstrConexaoCreditoFacil
oControle.oRota.m_timeOut = gstrTimeOutGeral
Set rsRota = oControle.oRota.consulta(cboRotas.ItemData(cboRotas.ListIndex))
oControle.oFuncionario.m_stringConexao = gstrConexaoCreditoFacil
oControle.oFuncionario.m_timeOut = gstrTimeOutGeral
Set rsFuncionarioRota = oControle.oFuncionario.consulta(rsRota("ID_FUNCIONARIO"))

If rsCobranca.EOF Then
    stbmsg.SimpleText = "Não há cobranças nesta data."
Else
    Set relCobranca = New rel_cobranca
    oControle.oEmpresa.m_timeOut = gstrTimeOutGeral
    oControle.oEmpresa.m_stringConexao = gstrConexaoCreditoFacil
    'Set rsEmpresaMestre = oControle.oEmpresa.recuperarEmpresas
    'Set x = LoadPicture(oControle.oEmpresa.carregarImagem(rsEmpresaMestre))
    'relCobranca.LogoEmpresa.Suppress = False
    'Set relCobranca.LogoEmpresa.FormattedPicture = x
    relCobranca.ParameterFields(2).AddCurrentValue CStr(DTPickerData)
    relCobranca.ParameterFields(1).AddCurrentValue Mid(cboRotas.Text, 1, InStr(1, cboRotas.Text, " ") - 1) '& " (" & rsFuncionario("TELEFONE1") & ")"
    relCobranca.Database.SetDataSource rsCobranca
    frmRelGenerico.crGenerico.ReportSource = relCobranca
    frmRelGenerico.Caption = "DEMONTRATIVO DE COBRANÇA"
    stbmsg.SimpleText = ""
    frmRelGenerico.crGenerico.ViewReport
    frmRelGenerico.Show vbModal
End If

End Sub

Private Sub Form_Load()
PopulaRotas
DTPickerData = Now()
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

