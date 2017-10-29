VERSION 5.00
Begin VB.Form frmFuncao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FUNCAO"
   ClientHeight    =   5580
   ClientLeft      =   4350
   ClientTop       =   2085
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7665
      TabIndex        =   24
      Top             =   4980
      Width           =   7665
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   25
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
      ScaleWidth      =   7665
      TabIndex        =   18
      Top             =   5280
      Width           =   7665
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmFUNCAO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmFUNCAO.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmFUNCAO.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmFUNCAO.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   23
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CT_LOCK"
      Height          =   285
      Index           =   8
      Left            =   2820
      TabIndex        =   17
      Top             =   3945
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATA_ALTERACAO"
      Height          =   285
      Index           =   7
      Left            =   2820
      TabIndex        =   15
      Top             =   3615
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "USUARIO_ALTERACAO"
      Height          =   285
      Index           =   6
      Left            =   2820
      TabIndex        =   13
      Top             =   3300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATA_INCLUSAO"
      Height          =   285
      Index           =   5
      Left            =   2820
      TabIndex        =   11
      Top             =   2985
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "USUARIO_INCLUSAO"
      Height          =   285
      Index           =   4
      Left            =   2820
      TabIndex        =   9
      Top             =   2655
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SIGLA_FUNCAO"
      Height          =   285
      Index           =   3
      Left            =   2820
      TabIndex        =   7
      Top             =   2340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STATUS_FUNCAO"
      Height          =   285
      Index           =   2
      Left            =   2820
      TabIndex        =   5
      Top             =   2025
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DESCRICAO_FUNCAO"
      Height          =   285
      Index           =   1
      Left            =   2820
      TabIndex        =   3
      Top             =   1695
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID_FUNCAO"
      Height          =   285
      Index           =   0
      Left            =   2820
      TabIndex        =   1
      Top             =   1380
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "CT_LOCK:"
      Height          =   255
      Index           =   8
      Left            =   900
      TabIndex        =   16
      Top             =   3945
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATA_ALTERACAO:"
      Height          =   255
      Index           =   7
      Left            =   900
      TabIndex        =   14
      Top             =   3615
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO_ALTERACAO:"
      Height          =   255
      Index           =   6
      Left            =   900
      TabIndex        =   12
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATA_INCLUSAO:"
      Height          =   255
      Index           =   5
      Left            =   900
      TabIndex        =   10
      Top             =   2985
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO_INCLUSAO:"
      Height          =   255
      Index           =   4
      Left            =   900
      TabIndex        =   8
      Top             =   2655
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SIGLA_FUNCAO:"
      Height          =   255
      Index           =   3
      Left            =   900
      TabIndex        =   6
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STATUS_FUNCAO:"
      Height          =   255
      Index           =   2
      Left            =   900
      TabIndex        =   4
      Top             =   2025
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DESCRICAO_FUNCAO:"
      Height          =   255
      Index           =   1
      Left            =   900
      TabIndex        =   2
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID_FUNCAO:"
      Height          =   255
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   1380
      Width           =   1815
   End
End
Attribute VB_Name = "frmFUNCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents PrimaryCLS As clsFUNCAO
Attribute PrimaryCLS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Set PrimaryCLS = New clsFUNCAO

  Dim oText As TextBox
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

