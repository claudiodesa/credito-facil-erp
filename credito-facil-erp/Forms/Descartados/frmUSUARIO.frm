VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USUARIO"
   ClientHeight    =   7065
   ClientLeft      =   5970
   ClientTop       =   5985
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10470
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   5220
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
            Picture         =   "frmUSUARIO.frx":0000
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSUARIO.frx":059A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSUARIO.frx":0B34
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSUARIO.frx":10CE
            Key             =   "Recarregar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSUARIO.frx":1668
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSUARIO.frx":1C02
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   6810
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10470
      TabIndex        =   30
      Top             =   6210
      Width           =   10470
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   6135
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Height          =   300
         Left            =   4980
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   9600
         TabIndex        =   35
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Atualizar"
         Height          =   300
         Left            =   8445
         TabIndex        =   34
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Deletar"
         Height          =   300
         Left            =   7290
         TabIndex        =   33
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   1213
         TabIndex        =   32
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Novo"
         Height          =   300
         Left            =   59
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10470
      TabIndex        =   24
      Top             =   6510
      Width           =   10470
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmUSUARIO.frx":219C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmUSUARIO.frx":24DE
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmUSUARIO.frx":2820
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmUSUARIO.frx":2B62
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   29
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CT_LOCK"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   4275
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATA_ALTERACAO"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   3945
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "USUARIO_ALTERACAO"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   3630
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATA_INCLUSAO"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   3315
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "USUARIO_INCLUSAO"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2985
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SENHA_USUARIO"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2340
      Width           =   2325
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LOGIN_USUARIO"
      Height          =   285
      Index           =   5
      Left            =   2070
      TabIndex        =   11
      Top             =   2355
      Width           =   1545
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STATUS_USUARIO"
      Height          =   285
      Index           =   4
      Left            =   7410
      TabIndex        =   9
      Top             =   1725
      Width           =   525
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NOME_USUARIO"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1710
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID_FUNCAO"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   1095
      Width           =   945
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID_EMPRESA"
      Height          =   285
      Index           =   1
      Left            =   4980
      TabIndex        =   3
      Top             =   765
      Width           =   945
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID_USUARIO"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   750
      Width           =   945
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
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
            Description     =   "Exclui o registro atual"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Description     =   "Fecha a janela atual"
            ImageIndex      =   6
         EndProperty
      EndProperty
      MouseIcon       =   "frmUSUARIO.frx":2EA4
   End
   Begin VB.Label lblLabels 
      Caption         =   "CT_LOCK:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   4275
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATA_ALTERACAO:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3945
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO_ALTERACAO:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   3630
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATA_INCLUSAO:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   3315
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO_INCLUSAO:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2985
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SENHA_USUARIO:"
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   12
      Top             =   2370
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LOGIN_USUARIO:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2355
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STATUS_USUARIO:"
      Height          =   255
      Index           =   4
      Left            =   5490
      TabIndex        =   8
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NOME_USUARIO:"
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   1710
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_FUNCAO:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_EMPRESA:"
      Height          =   255
      Index           =   1
      Left            =   3060
      TabIndex        =   2
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_USUARIO:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   750
      Width           =   1815
   End
End
Attribute VB_Name = "frmUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents PrimaryCLS As clsUSUARIO
Attribute PrimaryCLS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Set PrimaryCLS = New clsUSUARIO

  Dim oText As TextBox
  
  txtFields(7).Text = LogInUserID
  
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.DataMember = "Primary"
    Set oText.DataSource = PrimaryCLS
  Next
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub PrimaryCLS_MoveComplete()
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(PrimaryCLS.AbsolutePosition)
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  PrimaryCLS.AddNew
  lblStatus.Caption = "Add record"
  StatusBar1.SimpleText = "Novo registro"
  mbAddNewFlag = True
  SetButtons False

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  PrimaryCLS.Delete
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  PrimaryCLS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  PrimaryCLS.Cancel
  SetButtons True
End Sub

Private Sub cmdUpdate_Click()
Dim MD5 As New clsMD5
  
  On Error GoTo UpdateErr

 ' get the hash of the passwords
 Me.txtFields(6).Text = UCase(MD5.DigestStrToHexStr(Me.txtFields(6).Text))

  PrimaryCLS.Update
  SetButtons True
  MsgBox "Registro foi salvo com sucesso!", vbInformation
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  PrimaryCLS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  PrimaryCLS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  PrimaryCLS.MoveNext
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  PrimaryCLS.MovePrevious
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "&Novo"
        Call cmdAdd_Click
    Case "&Salvar"
        Call cmdUpdate_Click
    Case "&Recarregar"
        Call cmdRefresh_Click
        
End Select

End Sub

