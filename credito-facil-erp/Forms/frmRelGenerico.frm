VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form frmRelGenerico 
   ClientHeight    =   8235
   ClientLeft      =   8730
   ClientTop       =   2445
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer crGenerico 
      Height          =   7815
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   6225
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "frmRelGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Centraliza MDICreditoFacil, frmRelGenerico
End Sub

Private Sub Form_Resize()
    crGenerico.Top = 0
    crGenerico.Left = 0
    crGenerico.Height = ScaleHeight
    crGenerico.Width = ScaleWidth
End Sub
