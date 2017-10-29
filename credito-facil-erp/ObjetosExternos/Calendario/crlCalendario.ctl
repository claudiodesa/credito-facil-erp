VERSION 5.00
Begin VB.UserControl ctlCalendario 
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   LockControls    =   -1  'True
   ScaleHeight     =   2625
   ScaleMode       =   0  'User
   ScaleWidth      =   3525
   ToolboxBitmap   =   "crlCalendario.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   3855
      Top             =   660
   End
   Begin VB.PictureBox picMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   0
      ScaleHeight     =   2625
      ScaleMode       =   0  'User
      ScaleWidth      =   3495
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3495
      Begin VB.PictureBox picPainel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   3450
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   3450
         Begin VB.PictureBox lblHoje 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   1770
            Picture         =   "crlCalendario.ctx":0312
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   150
            Width           =   480
         End
         Begin VB.PictureBox imgPrev 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1620
            Picture         =   "crlCalendario.ctx":0BDC
            ScaleHeight     =   180
            ScaleWidth      =   105
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   120
            Width           =   105
         End
         Begin VB.PictureBox imgNext 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2310
            Picture         =   "crlCalendario.ctx":10B6
            ScaleHeight     =   180
            ScaleWidth      =   105
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   120
            Width           =   105
         End
         Begin VB.ComboBox cboAno 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "crlCalendario.ctx":1590
            Left            =   2520
            List            =   "crlCalendario.ctx":1592
            TabIndex        =   3
            Top             =   60
            Width           =   825
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "crlCalendario.ctx":1594
            Left            =   90
            List            =   "crlCalendario.ctx":15BF
            TabIndex        =   2
            Top             =   60
            Width           =   1470
         End
      End
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   53
         Top             =   2310
         Width           =   3345
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   36
         Left            =   570
         TabIndex        =   52
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   35
         Left            =   90
         TabIndex        =   51
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   34
         Left            =   2970
         TabIndex        =   50
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   33
         Left            =   2490
         TabIndex        =   49
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   32
         Left            =   2010
         TabIndex        =   48
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   31
         Left            =   1530
         TabIndex        =   47
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   30
         Left            =   1050
         TabIndex        =   46
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   29
         Left            =   570
         TabIndex        =   45
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   28
         Left            =   90
         TabIndex        =   44
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   27
         Left            =   2970
         TabIndex        =   43
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   26
         Left            =   2490
         TabIndex        =   42
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   25
         Left            =   2010
         TabIndex        =   41
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   1530
         TabIndex        =   40
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   23
         Left            =   1050
         TabIndex        =   39
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   22
         Left            =   570
         TabIndex        =   38
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   21
         Left            =   90
         TabIndex        =   37
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   2970
         TabIndex        =   36
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   2490
         TabIndex        =   35
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   2010
         TabIndex        =   34
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   1530
         TabIndex        =   33
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   1050
         TabIndex        =   32
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   570
         TabIndex        =   31
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   90
         TabIndex        =   30
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   2970
         TabIndex        =   29
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   2490
         TabIndex        =   28
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   2010
         TabIndex        =   27
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   1530
         TabIndex        =   26
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   1050
         TabIndex        =   25
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   570
         TabIndex        =   24
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   90
         TabIndex        =   23
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   2970
         TabIndex        =   22
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   2490
         TabIndex        =   21
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   2010
         TabIndex        =   20
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1530
         TabIndex        =   19
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1050
         TabIndex        =   18
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   570
         TabIndex        =   17
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Seg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   570
         TabIndex        =   14
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1050
         TabIndex        =   13
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Qua"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   1530
         TabIndex        =   12
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Qui"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   2010
         TabIndex        =   11
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   2490
         TabIndex        =   10
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sab"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   2970
         TabIndex        =   9
         Top             =   510
         Width           =   450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   90
         X2              =   3345
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   37
         Left            =   1050
         TabIndex        =   8
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   38
         Left            =   1530
         TabIndex        =   7
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   39
         Left            =   2010
         TabIndex        =   6
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   40
         Left            =   2490
         TabIndex        =   5
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lbaDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   41
         Left            =   2970
         TabIndex        =   4
         Top             =   2040
         Width           =   435
      End
      Begin VB.Line lnMens 
         X1              =   90
         X2              =   3420
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.Image lblHoje_ 
      Height          =   480
      Index           =   1
      Left            =   4290
      Picture         =   "crlCalendario.ctx":1628
      Top             =   390
      Width           =   480
   End
   Begin VB.Image lblHoje_ 
      Height          =   480
      Index           =   0
      Left            =   3735
      Picture         =   "crlCalendario.ctx":1EF2
      Top             =   420
      Width           =   480
   End
   Begin VB.Image imgNext_ 
      Height          =   180
      Index           =   1
      Left            =   4215
      Picture         =   "crlCalendario.ctx":27BC
      Top             =   195
      Width           =   105
   End
   Begin VB.Image imgPrev_ 
      Height          =   180
      Index           =   1
      Left            =   3900
      Picture         =   "crlCalendario.ctx":2C96
      Top             =   195
      Width           =   105
   End
   Begin VB.Image imgNext_ 
      Height          =   180
      Index           =   0
      Left            =   4215
      Picture         =   "crlCalendario.ctx":3170
      Top             =   0
      Width           =   105
   End
   Begin VB.Image imgPrev_ 
      Height          =   180
      Index           =   0
      Left            =   3900
      Picture         =   "crlCalendario.ctx":364A
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "ctlCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint& Lib "user32" (ByVal lpPointX As Long, ByVal lpPointY As Long)

'Default Property Values:
Const m_def_DiasForaMesBackColor = 0
Const m_def_DiasForaMesBackStyle = 0
Const m_def_MensagemCaption = "Mensagem"
Const m_def_MensagemBackStyle = 1
Const m_def_MensagemBackColor = &H808080
Const m_def_MensagemForeColor = &H80C0FF
Const m_def_HojeBackColor = vbWhite
Const m_def_HojeForeColor = vbRed
Const m_def_HojeBackStyle = 1

'Const m_def_AutoRedraw = 0
Const m_def_DataSelecionadaForeColor = vbWhite
Const m_def_DataSelecionadaBackColor = vbBlue
Const m_def_DataSelecionadaBackStyle = 1
'Const m_def_DataSelecionada = #12/20/2004#
Const m_def_DiasBackStyle = 0
Const m_def_DiasBackColor = 0
Const m_def_DiasForeColor = vbBlack
Const m_def_GradienteBackGround = 3
Const m_def_DiasForaMesForeColor = &HC0C0C0
Const m_def_BorderStyle = 0
Const m_def_ShowDateSelect = True
Const m_def_MostraMens = True
Const m_def_PainelBackColor = &HE0E0E0
 

'Property Variables:
Dim m_HojeBackColor As OLE_COLOR
Dim m_HojeForeColor As OLE_COLOR
Dim m_HojeBackStyle As MonthBackStyle
Dim m_DiasForaMesBackColor As OLE_COLOR
Dim m_DiasForaMesBackStyle As MonthBackStyle
Dim m_MensagemCaption As String
Dim m_MensagemBackStyle As MonthBackStyle
Dim m_MensagemForeColor As OLE_COLOR
Dim m_MensagemBackColor As OLE_COLOR
Dim m_DataSelecionadaForeColor As OLE_COLOR
Dim m_DataSelecionadaBackColor As OLE_COLOR
Dim m_PainelBackColor As OLE_COLOR
Dim m_DataSelecionadaBackStyle As MonthBackStyle
Dim m_DataSelecionada As Date
Dim m_DiasBackStyle As MonthBackStyle
Dim m_DiasForeColor As OLE_COLOR
Dim m_DiasForaMesForeColor As OLE_COLOR
Dim m_DiasBackColor As OLE_COLOR
Dim m_BorderStyle   As TBorderStyle
Dim m_ShowDateSelect As Boolean
Dim m_MostraMens As Boolean

Private mCursorPos As POINTAPI

Private Type POINTAPI
        x As Long
        y As Long
End Type

'Event Declarations:
Event Click()

Public Enum MonthBackStyle
  [Transparent] = 0
  [Opaque] = 1
End Enum

Public Enum TBorderStyle
  [None] = 0
  [Fixed Single] = 1
End Enum

'-------- Declaração de Variáveis --------
Dim MesCorrente As Date
Dim HojeData As Date
Dim Ligado As Integer

Private Sub cboAno_Click()
  MesCorrente = DateSerial(cboAno.List(cboAno.ListIndex), Month(MesCorrente), Day(MesCorrente))
  AtualizaCalendario
End Sub

Private Sub cboMes_Click()
  MesCorrente = DateSerial(Year(MesCorrente), cboMes.ListIndex + 1, Day(MesCorrente))
  AtualizaCalendario
End Sub

Private Sub imgNext_Click()
  Ligado = 1
  MesCorrente = DateAdd("m", 1, MesCorrente)
  AtualizaCalendario
End Sub

Private Sub imgPrev_Click()
  Ligado = 2
  MesCorrente = DateAdd("m", -1, MesCorrente)
  AtualizaCalendario
End Sub

Private Sub imgPrev_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Ligado > 0 And Ligado <> 2 Then ResetPic
  If imgPrev.Picture = imgPrev_(1).Picture Then Exit Sub
  imgPrev.Picture = imgPrev_(1).Picture
  Ligado = 2
  Timer1.Enabled = True
End Sub

Private Sub imgNext_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Ligado > 0 And Ligado <> 1 Then ResetPic
  If imgNext.Picture = imgNext_(1).Picture Then Exit Sub
  imgNext.Picture = imgNext_(1).Picture
  Ligado = 1
  Timer1.Enabled = True
End Sub

Private Sub lblHoje_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Ligado > 0 And Ligado <> 3 Then ResetPic
  If lblHoje.Picture = lblHoje_(1).Picture Then Exit Sub
  lblHoje.Picture = lblHoje_(1).Picture
  Ligado = 3
  Timer1.Enabled = True
End Sub

Private Sub lbaDay_Click(Index As Integer)
Dim DataClick As Date

  If (Month(lbaDay(Index).Tag) <> Month(MesCorrente)) Or _
     Year(lbaDay(Index).Tag) <> Year(MesCorrente) Then
     MesCorrente = DateSerial(Year(lbaDay(Index).Tag), Month(lbaDay(Index).Tag), Day(lbaDay(Index).Tag))
     m_DataSelecionada = MesCorrente
     AtualizaCalendario
  Else
     DataClick = DateAdd("d", Index, UltimoDomingo(MesCorrente))
     m_DataSelecionada = DataClick
     AtualizaDias
  End If
  RaiseEvent Click
End Sub

Private Sub lblHoje_Click()
  MesCorrente = CDate(HojeData)
  m_DataSelecionada = CDate(HojeData)
  AtualizaCalendario
  Ligado = 3
  RaiseEvent Click
End Sub

Private Sub picMonth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not lblHoje.FontBold Then Exit Sub
  lblHoje.FontBold = False
End Sub

Private Sub picPainel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not lblHoje.FontBold Then Exit Sub
  lblHoje.FontBold = False
End Sub

Private Sub Timer1_Timer()
Dim Status      As Long
Dim CurrenthWnd As Long
  
  If Ligado = 0 Then Exit Sub
  Status = GetCursorPos&(mCursorPos)
  CurrenthWnd = WindowFromPoint(mCursorPos.x, mCursorPos.y)
  
  If Ligado = 1 Then
     If CurrenthWnd <> imgNext.hwnd Then ResetPic: Timer1.Enabled = False
  ElseIf Ligado = 2 Then
     If CurrenthWnd <> imgPrev.hwnd Then ResetPic: Timer1.Enabled = False
  ElseIf Ligado = 3 Then
     If CurrenthWnd <> lblHoje.hwnd Then ResetPic: Timer1.Enabled = False
  End If
  
End Sub

Sub ResetPic()
    If imgPrev.Picture <> imgPrev_(0).Picture Then
        imgPrev.Picture = imgPrev_(0).Picture
    End If
    If imgNext.Picture <> imgNext_(0).Picture Then
        imgNext.Picture = imgNext_(0).Picture
    End If
    If lblHoje.Picture <> lblHoje_(0).Picture Then
        lblHoje.Picture = lblHoje_(0).Picture
    End If
    Ligado = 0
End Sub

Private Sub UserControl_Initialize()
Dim I As Integer

  For I = 1900 To 2100
     cboAno.AddItem I
  Next
  
  picMonth.Top = 0
  picMonth.Left = 0

  UserControl.ScaleWidth = 255
  UserControl.ScaleHeight = 255

  picMonth.ScaleHeight = UserControl.ScaleHeight
  picMonth.ScaleWidth = UserControl.ScaleWidth
  
  UserControl.Width = 3495
  UserControl.Height = 2655

End Sub

Public Sub UserControl_Resize()
  If m_BorderStyle = None Then
     UserControl.Width = 3495
  Else
     UserControl.Width = 3495 + 80
  End If

  If m_MostraMens = True Then
     UserControl.Height = 2655
  Else
     UserControl.Height = 2355
  End If
  
  lblMensagem.Visible = m_MostraMens
  lnMens.Visible = m_MostraMens
  
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_DiasForeColor = m_def_DiasForeColor
  m_DiasForaMesForeColor = m_def_DiasForaMesForeColor
  m_DiasBackStyle = m_def_DiasBackStyle
  m_DiasBackColor = m_def_DiasBackColor
  m_HojeBackColor = m_def_HojeBackColor
  m_HojeForeColor = m_def_HojeForeColor
  m_HojeBackStyle = m_def_HojeBackStyle
  m_BorderStyle = m_def_BorderStyle
  m_ShowDateSelect = m_def_ShowDateSelect
  m_DataSelecionadaForeColor = m_def_DataSelecionadaForeColor
  m_DataSelecionadaBackColor = m_def_DataSelecionadaBackColor
  m_DataSelecionadaBackStyle = m_def_DataSelecionadaBackStyle
  m_DataSelecionada = Format(Now, "dd/mm/yyyy")
  m_MensagemBackColor = m_def_MensagemBackColor
  m_MensagemForeColor = m_def_MensagemForeColor
  m_MensagemCaption = m_def_MensagemCaption
  m_MensagemBackStyle = m_def_MensagemBackStyle
  m_DiasForaMesBackColor = m_def_DiasForaMesBackColor
  m_DiasForaMesBackStyle = m_def_DiasForaMesBackStyle
  m_MostraMens = m_def_MostraMens
  UserControl.BackColor = vbWhite
  m_PainelBackColor = m_def_PainelBackColor
  MesCorrente = m_DataSelecionada
  AtualizaCalendario
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
  picMonth.BackColor() = PropBag.ReadProperty("BackColor", UserControl.BackColor)
  
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  m_DiasForeColor = PropBag.ReadProperty("DiasForeColor", m_def_DiasForeColor)
  m_DiasForaMesForeColor = PropBag.ReadProperty("DiasForaMesForeColor", m_def_DiasForaMesForeColor)

  m_DiasBackStyle = PropBag.ReadProperty("DiasBackStyle", m_def_DiasBackStyle)
  m_DiasBackColor = PropBag.ReadProperty("DiasBackColor", m_def_DiasBackColor)
  m_DataSelecionadaForeColor = PropBag.ReadProperty("DataSelecionadaForeColor", m_def_DataSelecionadaForeColor)
  m_DataSelecionadaBackColor = PropBag.ReadProperty("DataSelecionadaBackColor", m_def_DataSelecionadaBackColor)
  m_DataSelecionadaBackStyle = PropBag.ReadProperty("DataSelecionadaBackStyle", m_def_DataSelecionadaBackStyle)
  m_DataSelecionada = PropBag.ReadProperty("DataSelecionada", Format(Now, "dd/mm/yyyy"))
  m_MensagemBackColor = PropBag.ReadProperty("MensagemBackColor", m_def_MensagemBackColor)
  m_MensagemForeColor = PropBag.ReadProperty("MensagemForeColor", m_def_MensagemForeColor)
  m_MensagemCaption = PropBag.ReadProperty("MensagemCaption", m_def_MensagemCaption)
  m_MensagemBackStyle = PropBag.ReadProperty("MensagemBackStyle", m_def_MensagemBackStyle)
  m_HojeBackColor = PropBag.ReadProperty("HojeBackColor", m_def_HojeBackColor)
  m_HojeForeColor = PropBag.ReadProperty("HojeForeColor", m_def_HojeForeColor)
  m_HojeBackStyle = PropBag.ReadProperty("HojeBackStyle", m_def_HojeBackStyle)
  m_DiasForaMesBackColor = PropBag.ReadProperty("DiasForaMesBackColor", m_def_DiasForaMesBackColor)
  m_DiasForaMesBackStyle = PropBag.ReadProperty("DiasForaMesBackStyle", m_def_DiasForaMesBackStyle)
  m_ShowDateSelect = PropBag.ReadProperty("ShowDateSelect", m_def_ShowDateSelect)
  m_MostraMens = PropBag.ReadProperty("MostraMens", m_def_MostraMens)
  m_PainelBackColor = PropBag.ReadProperty("PainelBackColor", m_def_PainelBackColor)
  
  MesCorrente = m_DataSelecionada
  HojeData = Format(Now, "dd/mm/yyyy")
  AtualizaCalendario
End Sub

Private Sub UserControl_Terminate()
  Timer1.Enabled = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbWhite)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("DiasForeColor", m_DiasForeColor, m_def_DiasForeColor)
  Call PropBag.WriteProperty("DiasForaMesForeColor", m_DiasForaMesForeColor, m_def_DiasForaMesForeColor)
  Call PropBag.WriteProperty("DiasBackStyle", m_DiasBackStyle, m_def_DiasBackStyle)
  Call PropBag.WriteProperty("DiasBackColor", m_DiasBackColor, m_def_DiasBackColor)
  Call PropBag.WriteProperty("DataSelecionadaForeColor", m_DataSelecionadaForeColor, m_def_DataSelecionadaForeColor)
  Call PropBag.WriteProperty("DataSelecionadaBackColor", m_DataSelecionadaBackColor, m_def_DataSelecionadaBackColor)
  Call PropBag.WriteProperty("DataSelecionadaBackStyle", m_DataSelecionadaBackStyle, m_def_DataSelecionadaBackStyle)
  Call PropBag.WriteProperty("DataSelecionada", m_DataSelecionada, Format(Now, "dd/mm/yyyy"))
  Call PropBag.WriteProperty("MensagemBackColor", m_MensagemBackColor, m_def_MensagemBackColor)
  Call PropBag.WriteProperty("MensagemForeColor", m_MensagemForeColor, m_def_MensagemForeColor)
  Call PropBag.WriteProperty("MensagemCaption", m_MensagemCaption, m_def_MensagemCaption)
  Call PropBag.WriteProperty("MensagemBackStyle", m_MensagemBackStyle, m_def_MensagemBackStyle)
  Call PropBag.WriteProperty("DiasForaMesBackColor", m_DiasForaMesBackColor, m_def_DiasForaMesBackColor)
  Call PropBag.WriteProperty("DiasForaMesBackStyle", m_DiasForaMesBackStyle, m_def_DiasForaMesBackStyle)
  Call PropBag.WriteProperty("HojeBackColor", m_HojeBackColor, m_def_HojeBackColor)
  Call PropBag.WriteProperty("HojeForeColor", m_HojeForeColor, m_def_HojeForeColor)
  Call PropBag.WriteProperty("HojeBackStyle", m_HojeBackStyle, m_def_HojeBackStyle)
  Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
  Call PropBag.WriteProperty("ShowDateSelect", m_ShowDateSelect, m_def_ShowDateSelect)
  Call PropBag.WriteProperty("PainelBackColor", m_PainelBackColor, m_def_PainelBackColor)
  Call PropBag.WriteProperty("MostraMens", m_MostraMens, m_def_MostraMens)
End Sub

Private Function UltimoDomingo(ByVal dDia As Date) As String
Dim sPrimeiroDia As Date, I As Integer

    sPrimeiroDia = DateAdd("d", -Day(dDia) + 1, dDia)
     
    For I = 1 To 7
      UltimoDomingo = DateAdd("d", -I, sPrimeiroDia)
      If Weekday(UltimoDomingo) = vbSunday Then Exit Function
    Next

End Function

Private Sub AtualizaCalendario()
  AtualizaTitulo
  AtualizaMensagem
  AtualizaPainel
  AtualizaDias
End Sub

Private Sub AtualizaDias()
Dim iDay As Date, iWeek As Long
Dim MesSelecionado As Byte

MesSelecionado = Month(MesCorrente)

iDay = UltimoDomingo(MesCorrente)
picMonth.Visible = False
For iWeek = 0 To 41
    If MesSelecionado = Month(iDay) Then
       lbaDay(iWeek).BackStyle = m_DiasBackStyle
       lbaDay(iWeek).ForeColor = m_DiasForeColor
       lbaDay(iWeek).BackColor = m_DiasBackColor
    Else
       lbaDay(iWeek).BackStyle = m_DiasForaMesBackStyle
       lbaDay(iWeek).BackColor = m_DiasForaMesBackColor
       lbaDay(iWeek).ForeColor = m_DiasForaMesForeColor
    End If
    lbaDay(iWeek).Caption = Day(iDay)
    lbaDay(iWeek).Tag = iDay

'------ Verifica se Hoje faz parte das Datas -----
    If iDay = HojeData Then MarcaHoje iWeek

'----- Verifica se é data selecionada --------
    If iDay = m_DataSelecionada And m_ShowDateSelect = True Then MarcaData iWeek
    
    iDay = DateAdd("d", 1, iDay)
Next
picMonth.Visible = True
End Sub

Private Sub AtualizaTitulo()
Dim bMes As Byte, I As Integer

  On Error GoTo erro
  bMes = Month(MesCorrente)

  cboMes.Enabled = False: cboAno.Enabled = False
  cboMes.ListIndex = bMes - 1
  cboAno.ListIndex = Year(MesCorrente) - 1900
  cboMes.Enabled = True: cboAno.Enabled = True
erro:
End Sub

Private Sub AtualizaMensagem()
  lblMensagem.BackStyle = m_MensagemBackStyle
  lblMensagem.BackColor = m_MensagemBackColor
  lblMensagem.ForeColor = m_MensagemForeColor
  lblMensagem.Caption = m_MensagemCaption
End Sub

Private Sub AtualizaPainel()
  picPainel.BackColor = m_PainelBackColor
  lblHoje.BackColor = m_PainelBackColor
  imgNext.BackColor = m_PainelBackColor
  imgPrev.BackColor = m_PainelBackColor
End Sub
Private Sub MarcaHoje(ByVal Index As Long)
'------- Marcar a data de Hoje no Componente ------
      lbaDay(Index).BackStyle = m_HojeBackStyle
      lbaDay(Index).ForeColor = m_HojeForeColor
      lbaDay(Index).BackColor = m_HojeBackColor
End Sub

Private Sub MarcaData(ByVal Index As Long)
'------- Marcar a data de Hoje no Componente ------
      lbaDay(Index).BackStyle = m_DataSelecionadaBackStyle
      lbaDay(Index).ForeColor = m_DataSelecionadaForeColor
      lbaDay(Index).BackColor = m_DataSelecionadaBackColor
End Sub

Public Sub RetMarDia()
  m_ShowDateSelect = False
  AtualizaDias
  m_ShowDateSelect = True
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Public Property Get Value() As String
  Value = m_DataSelecionada
End Property

Public Property Let ValueSemClick(ByVal New_DataSelecionada As String)
  If Not IsDate(New_DataSelecionada) Then Exit Property
  
  m_DataSelecionada = CDate(New_DataSelecionada)
  MesCorrente = CDate(New_DataSelecionada)
  PropertyChanged "DataSelecionada"
  
  AtualizaCalendario
End Property

Public Property Let Value(ByVal New_DataSelecionada As String)
  If Not IsDate(New_DataSelecionada) Then Exit Property
  
  m_DataSelecionada = CDate(New_DataSelecionada)
  MesCorrente = CDate(New_DataSelecionada)
  PropertyChanged "DataSelecionada"
  
  AtualizaCalendario
   RaiseEvent Click
End Property

Public Property Get MensagemCaption() As String
  MensagemCaption = m_MensagemCaption
End Property

Public Property Let MensagemCaption(ByVal New_MensagemCaption As String)
  m_MensagemCaption = New_MensagemCaption
  PropertyChanged "MensagemCaption"
  AtualizaMensagem
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  picMonth.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
  If New_Enabled = True Then
     AtualizaPainel
  Else
     picPainel.BackColor = &H80000011
     lblHoje.BackColor = picPainel.BackColor
  End If
End Property

Public Property Get DiasForeColor() As OLE_COLOR
  DiasForeColor = m_DiasForeColor
End Property

Public Property Let DiasForeColor(ByVal New_DiasForeColor As OLE_COLOR)
  m_DiasForeColor = New_DiasForeColor
  PropertyChanged "DiasForeColor"
  AtualizaDias
End Property

Public Property Get DiasForaMesForeColor() As OLE_COLOR
  DiasForaMesForeColor = m_DiasForaMesForeColor
End Property

Public Property Let DiasForaMesForeColor(ByVal New_DiasForaMesForeColor As OLE_COLOR)
  m_DiasForaMesForeColor = New_DiasForaMesForeColor
  PropertyChanged "DiasForaMesForeColor"
  AtualizaDias
End Property

Public Property Get DiasBackStyle() As MonthBackStyle
  DiasBackStyle = m_DiasBackStyle
End Property

Public Property Let DiasBackStyle(ByVal New_DiasBackStyle As MonthBackStyle)
  m_DiasBackStyle = New_DiasBackStyle
  PropertyChanged "DiasBackStyle"
AtualizaDias
End Property

Public Property Get DiasBackColor() As OLE_COLOR
  DiasBackColor = m_DiasBackColor
End Property

Public Property Let DiasBackColor(ByVal New_DiasBackColor As OLE_COLOR)
  m_DiasBackColor = New_DiasBackColor
  PropertyChanged "DiasBackColor"
  AtualizaDias
End Property

Public Property Get DataSelecionadaForeColor() As OLE_COLOR
  DataSelecionadaForeColor = m_DataSelecionadaForeColor
End Property

Public Property Let DataSelecionadaForeColor(ByVal New_DataSelecionadaForeColor As OLE_COLOR)
  m_DataSelecionadaForeColor = New_DataSelecionadaForeColor
  PropertyChanged "DataSelecionadaForeColor"
End Property

Public Property Get DataSelecionadaBackColor() As OLE_COLOR
  DataSelecionadaBackColor = m_DataSelecionadaBackColor
  AtualizaDias
End Property

Public Property Let DataSelecionadaBackColor(ByVal New_DataSelecionadaBackColor As OLE_COLOR)
  m_DataSelecionadaBackColor = New_DataSelecionadaBackColor
  PropertyChanged "DataSelecionadaBackColor"
End Property

Public Property Get DataSelecionadaBackStyle() As MonthBackStyle
  DataSelecionadaBackStyle = m_DataSelecionadaBackStyle
End Property

Public Property Let DataSelecionadaBackStyle(ByVal New_DataSelecionadaBackStyle As MonthBackStyle)
  m_DataSelecionadaBackStyle = New_DataSelecionadaBackStyle
  PropertyChanged "DataSelecionadaBackStyle"
End Property

Public Property Get MensagemBackColor() As OLE_COLOR
  MensagemBackColor = m_MensagemBackColor
End Property

Public Property Let MensagemBackColor(ByVal New_MensagemBackColor As OLE_COLOR)
  m_MensagemBackColor = New_MensagemBackColor
  PropertyChanged "MensagemBackColor"
  AtualizaMensagem
End Property

Public Property Get MensagemForeColor() As OLE_COLOR
  MensagemForeColor = m_MensagemForeColor
End Property

Public Property Let MensagemForeColor(ByVal New_MensagemForeColor As OLE_COLOR)
  m_MensagemForeColor = New_MensagemForeColor
  PropertyChanged "MensagemForeColor"
  AtualizaMensagem
End Property

Public Property Get MensagemBackStyle() As MonthBackStyle
  MensagemBackStyle = m_MensagemBackStyle
End Property

Public Property Let MensagemBackStyle(ByVal New_MensagemBackStyle As MonthBackStyle)
  m_MensagemBackStyle = New_MensagemBackStyle
  PropertyChanged "MensagemBackStyle"
  AtualizaMensagem
End Property

Public Property Get HojeBackColor() As OLE_COLOR
  HojeBackColor = m_HojeBackColor
End Property

Public Property Let HojeBackColor(ByVal New_HojeBackColor As OLE_COLOR)
  m_HojeBackColor = New_HojeBackColor
  PropertyChanged "HojeBackColor"
  AtualizaMensagem
End Property

Public Property Get HojeForeColor() As OLE_COLOR
  HojeForeColor = m_HojeForeColor
End Property

Public Property Let HojeForeColor(ByVal New_HojeForeColor As OLE_COLOR)
  m_HojeForeColor = New_HojeForeColor
  PropertyChanged "HojeForeColor"
  AtualizaMensagem
End Property

Public Property Get HojeBackStyle() As MonthBackStyle
  HojeBackStyle = m_HojeBackStyle
End Property

Public Property Let HojeBackStyle(ByVal New_HojeBackStyle As MonthBackStyle)
  m_HojeBackStyle = New_HojeBackStyle
  PropertyChanged "HojeBackStyle"
  AtualizaMensagem
End Property

Public Property Get DiasForaMesBackColor() As OLE_COLOR
  DiasForaMesBackColor = m_DiasForaMesBackColor
End Property

Public Property Let DiasForaMesBackColor(ByVal New_DiasForaMesBackColor As OLE_COLOR)
  m_DiasForaMesBackColor = New_DiasForaMesBackColor
  PropertyChanged "DiasForaMesBackColor"
  AtualizaDias
End Property

Public Property Get DiasForaMesBackStyle() As MonthBackStyle
  DiasForaMesBackStyle = m_DiasForaMesBackStyle
End Property

Public Property Let DiasForaMesBackStyle(ByVal New_DiasForaMesBackStyle As MonthBackStyle)
  m_DiasForaMesBackStyle = New_DiasForaMesBackStyle
  PropertyChanged "DiasForaMesBackStyle"
  AtualizaDias
End Property

Public Property Get BorderSt() As TBorderStyle
  BorderSt = m_BorderStyle
End Property

Public Property Let BorderSt(ByVal New_BorderStyle As TBorderStyle)
  m_BorderStyle = New_BorderStyle
  PropertyChanged "BorderStyle"
  BorderStyle = New_BorderStyle
  UserControl_Resize
End Property

Public Property Get ShowDateSelect() As Boolean
  ShowDateSelect = m_ShowDateSelect
End Property

Public Property Let ShowDateSelect(ByVal New_ShowDateSelect As Boolean)
  m_ShowDateSelect = New_ShowDateSelect
  PropertyChanged "ShowDateSelect"
  AtualizaDias
End Property

Public Property Get PainelBackColor() As OLE_COLOR
  PainelBackColor = m_PainelBackColor
End Property

Public Property Let PainelBackColor(ByVal New_PainelBackColor As OLE_COLOR)
  m_PainelBackColor = New_PainelBackColor
  PropertyChanged "PainelBackColor"
  AtualizaPainel
End Property

Private Sub cboAno_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Public Property Get MostraMens() As Boolean
  MostraMens = m_MostraMens
End Property


Public Property Get hwnd() As Long
  hwnd = UserControl.hwnd
End Property

Public Property Let MostraMens(ByVal New_MostraMens As Boolean)
  m_MostraMens = New_MostraMens
  PropertyChanged "MostraMens"
  UserControl_Resize
End Property
