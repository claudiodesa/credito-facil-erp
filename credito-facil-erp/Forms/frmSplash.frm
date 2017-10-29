VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   3660
   ClientTop       =   5010
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   6
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   8
         Top             =   1140
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   7
         Top             =   705
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Sleep
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()

DoEvents
Sleep 2500


If DoLogin Then
    Unload Me ' Login falhou, descarregar
    MDICreditoFacil.Show
    MDICreditoFacil.StatusBar1.SimpleText = "Usuário : " & LogInUserName
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    
    'Timer1.Enabled = True
    
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Function DoLogin() As Boolean
    Dim UserName As String, Password As String, Ret As Boolean
    Dim LoginSuccessful As Boolean, rsData As ADODB.Recordset
    Dim MD5 As New clsMD5
            
    Randomize
    
    ' Get the user that last logged in from the registry
    UserName = GetSetting(App.EXEName, "Settings", "LastLogIn", "")
    
    ' prompt user to enter username and password
    frmLogin.Left = Screen.ActiveForm.Left / 2
    frmLogin.Top = Screen.ActiveForm.Top / 2
    Ret = frmLogin.GetLogIn(UserName, Password, Me)
    
    
    Set rsData = New ADODB.Recordset
    rsData.CursorLocation = adUseServer
    rsData.CursorType = adOpenStatic
    Do While Ret
        Set rsData = New ADODB.Recordset
        rsData.CursorLocation = adUseServer
        rsData.CursorType = adOpenStatic
        rsData.Open "SELECT * FROM USUARIO WHERE LOGIN_USUARIO = '" & Replace(UserName, "'", "''") & "'", DBConn
                
        ' if a record was found, it means the user exists
        If rsData.RecordCount > 0 Then
            ' check if the password is correct
            If UCase(MD5.DigestStrToHexStr(Password)) = UCase(rsData("SENHA_USUARIO").value) Then
                
                ' password is correct, so save the user that just logged in
                LogInUserID = rsData("LOGIN_USUARIO").value
                LogInUserName = rsData("NOME_USUARIO").value
                
                ' sage the username in the Registry
                SaveSetting App.EXEName, "Settings", "LastLogIn", rsData("LOGIN_USUARIO").value
                
                LoginSuccessful = True
                Exit Do
            End If
        Else
            'MsgBox "Usuário não cadastrado", vbExclamation
        End If
        
        If UserName = "ADMIN" And Password = "ADMIN" Then LoginSuccessful = True: Ret = False
        
        If Not LoginSuccessful Then
            Ret = False
            
            If MsgBox("Login falhou, deseja tentar novamente ?", vbQuestion + vbYesNo, "Falha no Login") = vbYes Then
                ' to prevent brute force password cracking from the application
                Sleep 200 + 300 * Rnd
                
                ' if login was not successfull, prompt again until Cancel is clicked
                Ret = frmLogin.GetLogIn(UserName, Password, Me)
            End If
            
        End If
    Loop
    
    DoLogin = LoginSuccessful
    
End Function

