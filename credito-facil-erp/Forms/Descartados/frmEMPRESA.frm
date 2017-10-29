VERSION 5.00
Begin VB.Form frmEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro: Empresa"
   ClientHeight    =   6465
   ClientLeft      =   2910
   ClientTop       =   2175
   ClientWidth     =   5835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5835
      TabIndex        =   42
      Top             =   5865
      Width           =   5835
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   47
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   46
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   45
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   44
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   43
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
      ScaleWidth      =   5835
      TabIndex        =   36
      Top             =   6165
      Width           =   5835
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmEMPRESA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmEMPRESA.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmEMPRESA.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmEMPRESA.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   41
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.PictureBox picFields 
      DataField       =   "BLOB_LOGO_REL_EMPRESA"
      Height          =   285
      Index           =   17
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   3315
      TabIndex        =   35
      Top             =   5500
      Width           =   3375
   End
   Begin VB.PictureBox picFields 
      DataField       =   "BLOB_LOGO_EMPRESA"
      Height          =   285
      Index           =   16
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   3315
      TabIndex        =   33
      Top             =   5180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "UF"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CIDADE"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Top             =   4540
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CEP"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   4220
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BAIRRO"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Top             =   3900
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "COMPLEMENTO"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   3580
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CT_LOCK"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   3260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NUMERO"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "USUARIO_ALTERACAO"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATA_ALTERACAO"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ENDERECO"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CGC_EMPRESA"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "USUARIO_INCLUSAO"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NOME_FANTASIA"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATA_INCLUSAO"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RAZAO_SOCIAL"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID_EMPRESA"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "BLOB_LOGO_REL_EMPRESA:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   5500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BLOB_LOGO_EMPRESA:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "UF:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CIDADE:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CEP:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BAIRRO:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "COMPLEMENTO:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CT_LOCK:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NUMERO:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO_ALTERACAO:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATA_ALTERACAO:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ENDERECO:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CGC_EMPRESA:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO_INCLUSAO:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NOME_FANTASIA:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATA_INCLUSAO:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RAZAO_SOCIAL:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_EMPRESA:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmEMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents PrimaryCLS As clsEMPRESA
Attribute PrimaryCLS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Set PrimaryCLS = New clsEMPRESA

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.DataMember = "Primary"
    Set oText.DataSource = PrimaryCLS
  Next
  Dim oPic As PictureBox
  'Bind the ole controls to the data provider
  For Each oPic In Me.picFields
    oPic.DataMember = "Primary"
    Set oPic.DataSource = PrimaryCLS
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
  On Error GoTo UpdateErr

  PrimaryCLS.Update
  SetButtons True
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

