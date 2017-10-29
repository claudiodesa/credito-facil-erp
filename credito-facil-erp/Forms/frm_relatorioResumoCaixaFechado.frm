VERSION 5.00
Begin VB.Form frm_relatorioResumoCaixaFechado 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Resumo de Caixa Fechado"
   ClientHeight    =   2115
   ClientLeft      =   6540
   ClientTop       =   4665
   ClientWidth     =   8220
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   8220
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGerarResumo 
      Caption         =   "Gerar Resumo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6390
      TabIndex        =   2
      Top             =   720
      Width           =   1605
   End
   Begin VB.ComboBox cboCaixasFechados 
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
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   5595
   End
   Begin VB.Label lblSelecao 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione um movimento de caixa fechado"
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
      Left            =   750
      TabIndex        =   1
      Top             =   450
      Width           =   3975
   End
End
Attribute VB_Name = "frm_relatorioResumoCaixaFechado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oControle As New clsCaixa
Private Sub cmdGerarResumo_Click()

Dim relResumo As rel_resumoCaixaAtual
Dim rs As ADODB.Recordset

If cboCaixasFechados.ListIndex = -1 Then
    MsgBox "Um movimento de caixa deve ser selecionado"
    Exit Sub
End If

Set rs = oControle.consultaResumoCaixaFechado(cboCaixasFechados.ItemData(cboCaixasFechados.ListIndex))

If rs.EOF Then
    MsgBox "Não há dados a apresentar, não existem lançamentos no caixa"
End If

Set relResumo = New rel_resumoCaixaAtual
relResumo.Database.SetDataSource rs
relResumo.ParameterFields(1).AddCurrentValue "RESUMO DO CAIXA FECHADO "
frmRelGenerico.crGenerico.ReportSource = relResumo
frmRelGenerico.Caption = "RESUMO DE CAIXA FECHADO / POR DATA"
frmRelGenerico.crGenerico.ViewReport
frmRelGenerico.Show vbModal


End Sub

Private Sub Form_Load()

oControle.mTIMEOUT = gstrTimeOutGeral
oControle.mSTRING_CONEXAO = gstrConexaoCreditoFacil
PopulaMovimentos

End Sub
Private Sub PopulaMovimentos()

Dim rs As ADODB.Recordset

Set rs = oControle.listarCaixasFechados
cboCaixasFechados.Clear
Do While Not rs.EOF
  cboCaixasFechados.AddItem rs(1)
  cboCaixasFechados.ItemData(cboCaixasFechados.NewIndex) = rs(0)
  rs.MoveNext
Loop

cboCaixasFechados.ListIndex = -1

End Sub

