VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPesquisa 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisar"
   ClientHeight    =   3765
   ClientLeft      =   4320
   ClientTop       =   2580
   ClientWidth     =   5640
   Icon            =   "frmPesquisa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btmOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
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
      Left            =   4470
      TabIndex        =   1
      Top             =   3450
      Width           =   1035
   End
   Begin MSDataListLib.DataCombo DataPesquisa 
      Bindings        =   "frmPesquisa.frx":058A
      DataSource      =   "AdoPesquisa"
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5054
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPesquisa 
      Height          =   330
      Left            =   120
      Top             =   3030
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483646
      ForeColor       =   -2147483643
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsResultset      As ADODB.Recordset
Public FieldsList       As String
Public FieldsKey        As String
Public FieldsReturn     As String

Private Sub btmOK_Click()
  'Atualiza variável de retorno com resultado da pesquisa
  FieldsReturn = DataPesquisa.BoundText
   If FieldsReturn = "" Then
    MsgBox "Tipo de consulta invalido!"
    Exit Sub
   End If
  Unload Me
End Sub

Private Sub DataPesquisa_Change()
  'Atualiza total de registros da pesquisa
  If rsResultset.BOF Then
      Exit Sub
  End If
  AdoPesquisa.Caption = "Total de registros : " & rsResultset.Bookmark & "/" & rsResultset.RecordCount
End Sub

Private Sub DataPesquisa_DblClick(Area As Integer)
  btmOK_Click
End Sub

'Atualiza propriedades do DataCombo
Private Sub Form_Load()
    'frmPesquisa.Show
  'CentralizaFormulario Me
  
  Set AdoPesquisa.Recordset = rsResultset
  Set DataPesquisa.DataSource = AdoPesquisa
  
  DataPesquisa.BoundColumn = FieldsKey
  DataPesquisa.ListField = FieldsList
  AdoPesquisa.Caption = "Total de registros : " & rsResultset.RecordCount
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rsResultset = Nothing
End Sub
