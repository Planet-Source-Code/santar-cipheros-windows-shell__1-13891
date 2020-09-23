VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cipher OS"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "CipherOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "CipherOS.frx":4BDA
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frmChooseSkin 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   9480
      TabIndex        =   78
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
      Begin VB.DirListBox DirectorySkin 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   3690
         Left            =   120
         TabIndex        =   81
         Top             =   660
         Width           =   2175
      End
      Begin VB.DriveListBox DriveSkin 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblCancelSkin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   83
         Top             =   4500
         Width           =   975
      End
      Begin VB.Label lblApplySkin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   4500
         Width           =   975
      End
      Begin VB.Image btnCancelSkin 
         Height          =   360
         Left            =   1320
         Picture         =   "CipherOS.frx":2081B
         Top             =   4440
         Width           =   960
      End
      Begin VB.Image btnApplySkin 
         Height          =   360
         Left            =   120
         Picture         =   "CipherOS.frx":21A5F
         Top             =   4440
         Width           =   960
      End
      Begin VB.Label lblSkinCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Choose a Skin Directory..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   79
         Top             =   40
         Width           =   2415
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   4935
         Left            =   0
         Top             =   0
         Width           =   2415
      End
      Begin VB.Image imgChangeSkin 
         Height          =   10500
         Left            =   0
         Picture         =   "CipherOS.frx":22CA3
         Top             =   0
         Width           =   10500
      End
   End
   Begin VB.Frame frmItemValue 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   4920
      TabIndex        =   69
      Top             =   4920
      Visible         =   0   'False
      Width           =   5655
      Begin MSComDlg.CommonDialog dlgBrowseValue 
         Left            =   1320
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "exe"
         DialogTitle     =   "Choose Exe App for Item"
         Filter          =   "Executable Program File|*.exe"
      End
      Begin VB.TextBox txtNewValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   75
         ToolTipText     =   "Double Click to see Common Dialog"
         Top             =   1005
         Width           =   3255
      End
      Begin VB.TextBox txtNewName 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   73
         Top             =   690
         Width           =   3255
      End
      Begin VB.ComboBox cmbAllItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4620
         TabIndex        =   77
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblDone 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3660
         TabIndex        =   76
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image btnCancel 
         Height          =   360
         Left            =   4575
         Picture         =   "CipherOS.frx":2B446
         Top             =   1365
         Width           =   960
      End
      Begin VB.Image btnDone 
         Height          =   360
         Left            =   3600
         Picture         =   "CipherOS.frx":2C68A
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label lblNewValue 
         BackStyle       =   0  'Transparent
         Caption         =   "New Value of the Item:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1005
         Width           =   2055
      End
      Begin VB.Label lblNewName 
         BackStyle       =   0  'Transparent
         Caption         =   "New Name of the Item:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   690
         Width           =   2055
      End
      Begin VB.Label lblItemInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Item You Want to Change the Value from:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   5655
      End
      Begin VB.Image imgItemValue 
         Height          =   10500
         Left            =   0
         Picture         =   "CipherOS.frx":2D8CE
         Top             =   0
         Width           =   10500
      End
   End
   Begin VB.Frame frmShuttingDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3240
      TabIndex        =   67
      Top             =   10440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label lblShuttingDown 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CipherOS is closing. Your computer is now shutting down. Please be patient..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   0
         TabIndex        =   68
         Top             =   210
         Width           =   4335
      End
      Begin VB.Image imgShuttingDown 
         Height          =   10500
         Left            =   0
         Picture         =   "CipherOS.frx":36071
         Top             =   0
         Width           =   10500
      End
   End
   Begin VB.Timer tmrTask 
      Interval        =   1000
      Left            =   10800
      Top             =   9960
   End
   Begin VB.TextBox txtRunFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   9480
      TabIndex        =   64
      Text            =   "Commands are: Exit and Name"
      Top             =   11080
      Width           =   4095
   End
   Begin VB.Frame frameShellSwitcher 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   5760
      TabIndex        =   59
      Top             =   4680
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label lblShell1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Choose the Shell you want your computer to start with. The changes will apply next time you shut down your computer."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Left            =   360
         TabIndex        =   63
         Top             =   120
         Width           =   3105
      End
      Begin VB.Label lblShellCipherOS 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CipherOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   960
         TabIndex        =   62
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label lblShellExplorer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Explorer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   960
         TabIndex        =   61
         Top             =   1635
         Width           =   2295
      End
      Begin VB.Label lblShellCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   960
         TabIndex        =   60
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image imgTop 
         Height          =   360
         Index           =   5
         Left            =   600
         Picture         =   "CipherOS.frx":3E814
         Top             =   1200
         Width           =   2715
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   25
         Left            =   600
         Picture         =   "CipherOS.frx":3EBC5
         Top             =   1560
         Width           =   2715
      End
      Begin VB.Image imgTitle 
         Height          =   360
         Index           =   5
         Left            =   600
         Picture         =   "CipherOS.frx":3EF8A
         Top             =   1920
         Width           =   2490
      End
      Begin VB.Image imgSHellSwitch 
         Height          =   10500
         Left            =   0
         Picture         =   "CipherOS.frx":3F392
         Top             =   0
         Width           =   10500
      End
   End
   Begin VB.Frame frameControlPanel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   4800
      TabIndex        =   39
      Top             =   3720
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   4215
         Left            =   0
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label lblMultimedia 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Multimedia Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   600
         TabIndex        =   58
         Top             =   2235
         Width           =   2175
      End
      Begin VB.Label lblInternetOpt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   630
         TabIndex        =   57
         Top             =   1875
         Width           =   2175
      End
      Begin VB.Label lblKeyBoard 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   675
         TabIndex        =   56
         Top             =   1515
         Width           =   2055
      End
      Begin VB.Label lblJoystick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Joystick Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   720
         TabIndex        =   55
         Top             =   1155
         Width           =   2055
      End
      Begin VB.Label lblDisplay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Display Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   3360
         TabIndex        =   54
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lblAccessability 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accessability Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   180
         Left            =   720
         TabIndex        =   53
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label lblModem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modem Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   3240
         TabIndex        =   52
         Top             =   2220
         Width           =   2175
      End
      Begin VB.Label lblMailFax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mail/Fax Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   3270
         TabIndex        =   51
         Top             =   1860
         Width           =   2175
      End
      Begin VB.Label lblPrinter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printers Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   3315
         TabIndex        =   50
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label lblMouse 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   3360
         TabIndex        =   49
         Top             =   1140
         Width           =   2055
      End
      Begin VB.Label lblRegional 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regional Settings"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   600
         TabIndex        =   48
         Top             =   2955
         Width           =   2175
      End
      Begin VB.Label lblAddRemProm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add/Remove Program"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   180
         Left            =   3360
         TabIndex        =   47
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label lblShSwitch 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shell Switcher"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   3240
         TabIndex        =   46
         Top             =   3300
         Width           =   2175
      End
      Begin VB.Label lblPassword 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Password Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   3240
         TabIndex        =   45
         Top             =   2580
         Width           =   2265
      End
      Begin VB.Label lblAddHardware 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Hardware"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   600
         TabIndex        =   44
         Top             =   795
         Width           =   2295
      End
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time/Date Properties"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   600
         TabIndex        =   43
         Top             =   3315
         Width           =   2175
      End
      Begin VB.Label lblSystem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System Properties"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   3240
         TabIndex        =   42
         Top             =   2925
         Width           =   2295
      End
      Begin VB.Label lblNetwork 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Network Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   600
         TabIndex        =   41
         Top             =   2595
         Width           =   2175
      End
      Begin VB.Label lblCloseCtrlPnel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   180
         Left            =   600
         TabIndex        =   40
         Top             =   3675
         Width           =   1455
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   22
         Left            =   240
         Picture         =   "CipherOS.frx":47B35
         Top             =   2520
         Width           =   2715
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   10
         Left            =   3000
         Picture         =   "CipherOS.frx":47EFA
         Top             =   2520
         Width           =   2625
      End
      Begin VB.Image imgTitle 
         Height          =   360
         Index           =   4
         Left            =   240
         Picture         =   "CipherOS.frx":48261
         Top             =   3600
         Width           =   2490
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   23
         Left            =   240
         Picture         =   "CipherOS.frx":48669
         Top             =   2880
         Width           =   2715
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   11
         Left            =   3000
         Picture         =   "CipherOS.frx":48A2E
         Top             =   2880
         Width           =   2625
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   24
         Left            =   240
         Picture         =   "CipherOS.frx":48D95
         Top             =   3240
         Width           =   2715
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   12
         Left            =   3000
         Picture         =   "CipherOS.frx":4915A
         Top             =   3240
         Width           =   2625
      End
      Begin VB.Image imgTop 
         Height          =   360
         Index           =   4
         Left            =   240
         Picture         =   "CipherOS.frx":494C1
         Top             =   360
         Width           =   2715
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   17
         Left            =   240
         Picture         =   "CipherOS.frx":49872
         Top             =   720
         Width           =   2715
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   18
         Left            =   240
         Picture         =   "CipherOS.frx":49C37
         Top             =   1080
         Width           =   2715
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   19
         Left            =   240
         Picture         =   "CipherOS.frx":49FFC
         Top             =   1440
         Width           =   2715
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   20
         Left            =   240
         Picture         =   "CipherOS.frx":4A3C1
         Top             =   1800
         Width           =   2715
      End
      Begin VB.Image imgLinkMid 
         Height          =   375
         Index           =   21
         Left            =   240
         Picture         =   "CipherOS.frx":4A786
         Top             =   2160
         Width           =   2715
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   4
         Left            =   3000
         Picture         =   "CipherOS.frx":4AB4B
         Top             =   360
         Width           =   2625
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   5
         Left            =   3000
         Picture         =   "CipherOS.frx":4AEB2
         Top             =   720
         Width           =   2625
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   6
         Left            =   3000
         Picture         =   "CipherOS.frx":4B219
         Top             =   1080
         Width           =   2625
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   7
         Left            =   3000
         Picture         =   "CipherOS.frx":4B580
         Top             =   1440
         Width           =   2625
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   8
         Left            =   3000
         Picture         =   "CipherOS.frx":4B8E7
         Top             =   1800
         Width           =   2625
      End
      Begin VB.Image imgRight 
         Height          =   330
         Index           =   9
         Left            =   3000
         Picture         =   "CipherOS.frx":4BC4E
         Top             =   2160
         Width           =   2625
      End
      Begin VB.Image imgControlPanel 
         Height          =   10500
         Left            =   0
         Picture         =   "CipherOS.frx":4BFB5
         Top             =   0
         Width           =   10500
      End
   End
   Begin VB.Timer TimerDate 
      Interval        =   1000
      Left            =   6120
      Top             =   120
   End
   Begin VB.ListBox lstTasks 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   1395
      Left            =   11945
      Sorted          =   -1  'True
      TabIndex        =   66
      ToolTipText     =   "Double Click on the itm in the list you want to see."
      Top             =   9480
      Width           =   2655
   End
   Begin VB.Label lblRunFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Run File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   13680
      TabIndex        =   65
      Top             =   11100
      Width           =   975
   End
   Begin VB.Image btnRunFile 
      Height          =   360
      Left            =   13680
      Picture         =   "CipherOS.frx":54758
      Top             =   11040
      Width           =   960
   End
   Begin VB.Label lblTimeDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time and Date are placed here..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   225
      Left            =   6960
      TabIndex        =   38
      Top             =   100
      Width           =   3135
   End
   Begin VB.Label lblDriveD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drive D:/"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   37
      ToolTipText     =   "Explore CD-ROM"
      Top             =   2830
      Width           =   2175
   End
   Begin VB.Label lblDriveC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drive C:/"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   36
      ToolTipText     =   "Explore Hard Drive"
      Top             =   2465
      Width           =   2175
   End
   Begin VB.Label lblDriveA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drive A:/"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   35
      ToolTipText     =   "Explore Floppy Disk"
      Top             =   2100
      Width           =   2175
   End
   Begin VB.Label lblChangeSkin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Skin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   34
      ToolTipText     =   "Start Explorer In Default Folder"
      Top             =   3190
      Width           =   2175
   End
   Begin VB.Label lblShellSwitcher 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shell Switcher"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   33
      ToolTipText     =   "Change The Default Starting Shell"
      Top             =   4270
      Width           =   2175
   End
   Begin VB.Label lblRestart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   32
      ToolTipText     =   "Reboot The Computer"
      Top             =   4630
      Width           =   2175
   End
   Begin VB.Label lblPrograms 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Programs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   31
      ToolTipText     =   "Pop Up the Programs Menu"
      Top             =   3540
      Width           =   2175
   End
   Begin VB.Label lblControlPanel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   30
      ToolTipText     =   "Show The Control Panel"
      Top             =   3910
      Width           =   2175
   End
   Begin VB.Label Sec1Item6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec1Item6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   29
      Top             =   2235
      Width           =   2175
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   4
      Left            =   240
      Picture         =   "CipherOS.frx":5599C
      Top             =   2160
      Width           =   2715
   End
   Begin VB.Label Sec4Item9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3240
      TabIndex        =   28
      Top             =   9900
      Width           =   2205
   End
   Begin VB.Label Sec4Item8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3240
      TabIndex        =   27
      Top             =   9540
      Width           =   2175
   End
   Begin VB.Label Sec4Item6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3240
      TabIndex        =   26
      Top             =   8820
      Width           =   2145
   End
   Begin VB.Label Sec4Item5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   25
      Top             =   10275
      Width           =   2175
   End
   Begin VB.Label Sec4Item7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3240
      TabIndex        =   24
      Top             =   9180
      Width           =   2175
   End
   Begin VB.Label Sec4Item4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   23
      Top             =   9915
      Width           =   2175
   End
   Begin VB.Label Sec4Item3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   22
      Top             =   9555
      Width           =   2175
   End
   Begin VB.Label Sec4Item2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   21
      Top             =   9195
      Width           =   2175
   End
   Begin VB.Label Sec4Item1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec4Item1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   600
      TabIndex        =   20
      Top             =   8820
      Width           =   2175
   End
   Begin VB.Label Sec3Item5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec3Item5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   19
      Top             =   7515
      Width           =   2175
   End
   Begin VB.Label Sec3Item4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec3Item4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   18
      Top             =   7155
      Width           =   2175
   End
   Begin VB.Label Sec3Item3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec3Item3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   17
      Top             =   6795
      Width           =   2175
   End
   Begin VB.Label Sec3Item2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec3Item2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   16
      Top             =   6435
      Width           =   2175
   End
   Begin VB.Label Sec3Item1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec3Item1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   600
      TabIndex        =   15
      Top             =   6060
      Width           =   2175
   End
   Begin VB.Label Sec2Item4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec2Item4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   14
      Top             =   4635
      Width           =   2175
   End
   Begin VB.Label Sec2Item3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec2Item3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   13
      Top             =   4275
      Width           =   2175
   End
   Begin VB.Label Sec2Item2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec2Item2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   180
      Left            =   600
      TabIndex        =   12
      Top             =   3915
      Width           =   2175
   End
   Begin VB.Label Sec2Item1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec2Item1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   180
      Left            =   600
      TabIndex        =   11
      Top             =   3540
      Width           =   2175
   End
   Begin VB.Label Sec1Item5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec1Item5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   630
      TabIndex        =   10
      Top             =   1875
      Width           =   2085
   End
   Begin VB.Label Sec1Item4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec1Item4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   630
      TabIndex        =   9
      Top             =   1515
      Width           =   2100
   End
   Begin VB.Label Sec1Item3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec1Item3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   1155
      Width           =   2175
   End
   Begin VB.Label Sec1Item2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec1Item2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   210
      Left            =   600
      TabIndex        =   7
      Top             =   795
      Width           =   2175
   End
   Begin VB.Label Sec1Item1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sec1Item1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   420
      Width           =   2175
   End
   Begin VB.Label Section4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Section4Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   675
      TabIndex        =   5
      Top             =   10630
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cipher Operating System 1.0 alpha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9430
      TabIndex        =   4
      Top             =   985
      Width           =   3135
   End
   Begin VB.Label Section3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Section3Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   7875
      Width           =   1305
   End
   Begin VB.Label Section2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Section2Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   4995
      Width           =   1455
   End
   Begin VB.Label Section1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Section1Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   1
      Top             =   2595
      Width           =   1455
   End
   Begin VB.Image imgTop 
      Height          =   360
      Index           =   3
      Left            =   240
      Picture         =   "CipherOS.frx":55D61
      Top             =   8760
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   12
      Left            =   240
      Picture         =   "CipherOS.frx":56112
      Top             =   9120
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   13
      Left            =   240
      Picture         =   "CipherOS.frx":564D7
      Top             =   9480
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   14
      Left            =   240
      Picture         =   "CipherOS.frx":5689C
      Top             =   9840
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   15
      Left            =   240
      Picture         =   "CipherOS.frx":56C61
      Top             =   10200
      Width           =   2715
   End
   Begin VB.Image imgTitle 
      Height          =   360
      Index           =   3
      Left            =   240
      Picture         =   "CipherOS.frx":57026
      Top             =   10560
      Width           =   2490
   End
   Begin VB.Image imgTop 
      Height          =   360
      Index           =   2
      Left            =   240
      Picture         =   "CipherOS.frx":5742E
      Top             =   6000
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   8
      Left            =   240
      Picture         =   "CipherOS.frx":577DF
      Top             =   6360
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   9
      Left            =   240
      Picture         =   "CipherOS.frx":57BA4
      Top             =   6720
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   10
      Left            =   240
      Picture         =   "CipherOS.frx":57F69
      Top             =   7080
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   11
      Left            =   240
      Picture         =   "CipherOS.frx":5832E
      Top             =   7440
      Width           =   2715
   End
   Begin VB.Image imgTitle 
      Height          =   360
      Index           =   2
      Left            =   240
      Picture         =   "CipherOS.frx":586F3
      Top             =   7800
      Width           =   2490
   End
   Begin VB.Image imgTop 
      Height          =   360
      Index           =   1
      Left            =   240
      Picture         =   "CipherOS.frx":58AFB
      Top             =   3480
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   5
      Left            =   240
      Picture         =   "CipherOS.frx":58EAC
      Top             =   3840
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   6
      Left            =   240
      Picture         =   "CipherOS.frx":59271
      Top             =   4200
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   7
      Left            =   240
      Picture         =   "CipherOS.frx":59636
      Top             =   4560
      Width           =   2715
   End
   Begin VB.Image imgTitle 
      Height          =   360
      Index           =   1
      Left            =   240
      Picture         =   "CipherOS.frx":599FB
      Top             =   4920
      Width           =   2490
   End
   Begin VB.Image imgTitle 
      Height          =   360
      Index           =   0
      Left            =   240
      Picture         =   "CipherOS.frx":59E03
      Top             =   2520
      Width           =   2490
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   3
      Left            =   240
      Picture         =   "CipherOS.frx":5A20B
      Top             =   1800
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   2
      Left            =   240
      Picture         =   "CipherOS.frx":5A5D0
      Top             =   1440
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   1
      Left            =   240
      Picture         =   "CipherOS.frx":5A995
      Top             =   1080
      Width           =   2715
   End
   Begin VB.Image imgLinkMid 
      Height          =   375
      Index           =   0
      Left            =   240
      Picture         =   "CipherOS.frx":5AD5A
      Top             =   720
      Width           =   2715
   End
   Begin VB.Image imgTop 
      Height          =   360
      Index           =   0
      Left            =   240
      Picture         =   "CipherOS.frx":5B11F
      Top             =   360
      Width           =   2715
   End
   Begin VB.Label lblShutDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shut Down"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   0
      ToolTipText     =   "Shut Down The Computer"
      Top             =   4990
      Width           =   2175
   End
   Begin VB.Image imgRight 
      Height          =   330
      Index           =   2
      Left            =   2970
      Picture         =   "CipherOS.frx":5B4D0
      Top             =   9480
      Width           =   2625
   End
   Begin VB.Image imgRight 
      Height          =   330
      Index           =   3
      Left            =   2970
      Picture         =   "CipherOS.frx":5B837
      Top             =   9840
      Width           =   2625
   End
   Begin VB.Image imgRight 
      Height          =   330
      Index           =   1
      Left            =   3000
      Picture         =   "CipherOS.frx":5BB9E
      Top             =   9120
      Width           =   2625
   End
   Begin VB.Image imgRight 
      Height          =   330
      Index           =   0
      Left            =   2970
      Picture         =   "CipherOS.frx":5BF05
      Top             =   8760
      Width           =   2625
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi,
'So I would like to say some things. Thank You to the guy from HardShell.
'And thank you to BeOS... I modified their Menu code.


'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'    Start Item Hover Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////

Public Sub ButtonsAllGray()
'have all the item buttons go gray
Sec1Item1.ForeColor = &H8000000F
Sec1Item2.ForeColor = &H8000000F
Sec1Item3.ForeColor = &H8000000F
Sec1Item4.ForeColor = &H8000000F
Sec1Item5.ForeColor = &H8000000F
Sec1Item6.ForeColor = &H8000000F
Sec2Item1.ForeColor = &H8000000F
Sec2Item2.ForeColor = &H8000000F
Sec2Item3.ForeColor = &H8000000F
Sec2Item4.ForeColor = &H8000000F
Sec3Item1.ForeColor = &H8000000F
Sec3Item2.ForeColor = &H8000000F
Sec3Item3.ForeColor = &H8000000F
Sec3Item4.ForeColor = &H8000000F
Sec3Item5.ForeColor = &H8000000F
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
lblRunFile.ForeColor = &H8000000F
End Sub


Private Sub DriveSkin_Change()
On Error Resume Next
DirectorySkin.path = DriveSkin.Drive
End Sub

Private Sub imgTitle_Click(Index As Integer)
Call ButtonsAllGray
End Sub

Private Sub lblApplySkin_Click()
Call WriteToINI("Skins", "SkinFolder", DirectorySkin.path, App.path & "\CipherOS.ini")  'Write to INI new Skin Path
frmChooseSkin.Visible = False  'hide this frame
Call ApplySkin 'Apply the new skin
End Sub

Private Sub lblCancelSkin_Click()
frmChooseSkin.Visible = False 'hide the frame
End Sub

Private Sub lblChangeSkin_Click()
frmChooseSkin.Visible = True
End Sub

Private Sub lblCloseCtrlPnel_Click()
frameControlPanel.Visible = False
End Sub

'Section 1 Items
Private Sub Sec1Item1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec1Item1.ForeColor = vbWhite
Sec1Item2.ForeColor = &H8000000F
Sec1Item3.ForeColor = &H8000000F
Sec1Item4.ForeColor = &H8000000F
Sec1Item5.ForeColor = &H8000000F
Sec1Item6.ForeColor = &H8000000F
End Sub
Private Sub Sec1Item2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec1Item1.ForeColor = &H8000000F
Sec1Item2.ForeColor = vbWhite
Sec1Item3.ForeColor = &H8000000F
Sec1Item4.ForeColor = &H8000000F
Sec1Item5.ForeColor = &H8000000F
Sec1Item6.ForeColor = &H8000000F
End Sub
Private Sub Sec1Item3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec1Item1.ForeColor = &H8000000F
Sec1Item2.ForeColor = &H8000000F
Sec1Item3.ForeColor = vbWhite
Sec1Item4.ForeColor = &H8000000F
Sec1Item5.ForeColor = &H8000000F
Sec1Item6.ForeColor = &H8000000F
End Sub
Private Sub Sec1Item4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec1Item1.ForeColor = &H8000000F
Sec1Item2.ForeColor = &H8000000F
Sec1Item3.ForeColor = &H8000000F
Sec1Item4.ForeColor = vbWhite
Sec1Item5.ForeColor = &H8000000F
Sec1Item6.ForeColor = &H8000000F
End Sub
Private Sub Sec1Item5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec1Item1.ForeColor = &H8000000F
Sec1Item2.ForeColor = &H8000000F
Sec1Item3.ForeColor = &H8000000F
Sec1Item4.ForeColor = &H8000000F
Sec1Item5.ForeColor = vbWhite
Sec1Item6.ForeColor = &H8000000F
End Sub
Private Sub Sec1Item6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec1Item1.ForeColor = &H8000000F
Sec1Item2.ForeColor = &H8000000F
Sec1Item3.ForeColor = &H8000000F
Sec1Item4.ForeColor = &H8000000F
Sec1Item5.ForeColor = &H8000000F
Sec1Item6.ForeColor = vbWhite
End Sub
'Section 2 Items
Private Sub Sec2Item1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec2Item1.ForeColor = vbWhite
Sec2Item2.ForeColor = &H8000000F
Sec2Item3.ForeColor = &H8000000F
Sec2Item4.ForeColor = &H8000000F
End Sub
Private Sub Sec2Item2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec2Item1.ForeColor = &H8000000F
Sec2Item2.ForeColor = vbWhite
Sec2Item3.ForeColor = &H8000000F
Sec2Item4.ForeColor = &H8000000F
End Sub
Private Sub Sec2Item3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec2Item1.ForeColor = &H8000000F
Sec2Item2.ForeColor = &H8000000F
Sec2Item3.ForeColor = vbWhite
Sec2Item4.ForeColor = &H8000000F
End Sub
Private Sub Sec2Item4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec2Item1.ForeColor = &H8000000F
Sec2Item2.ForeColor = &H8000000F
Sec2Item3.ForeColor = &H8000000F
Sec2Item4.ForeColor = vbWhite
End Sub
'Section 3 Items
Private Sub Sec3Item1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec3Item1.ForeColor = vbWhite
Sec3Item2.ForeColor = &H8000000F
Sec3Item3.ForeColor = &H8000000F
Sec3Item4.ForeColor = &H8000000F
Sec3Item5.ForeColor = &H8000000F
End Sub
Private Sub Sec3Item2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec3Item1.ForeColor = &H8000000F
Sec3Item2.ForeColor = vbWhite
Sec3Item3.ForeColor = &H8000000F
Sec3Item4.ForeColor = &H8000000F
Sec3Item5.ForeColor = &H8000000F
End Sub
Private Sub Sec3Item3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec3Item1.ForeColor = &H8000000F
Sec3Item2.ForeColor = &H8000000F
Sec3Item3.ForeColor = vbWhite
Sec3Item4.ForeColor = &H8000000F
Sec3Item5.ForeColor = &H8000000F
End Sub
Private Sub Sec3Item4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec3Item1.ForeColor = &H8000000F
Sec3Item2.ForeColor = &H8000000F
Sec3Item3.ForeColor = &H8000000F
Sec3Item4.ForeColor = vbWhite
Sec3Item5.ForeColor = &H8000000F
End Sub
Private Sub Sec3Item5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec3Item1.ForeColor = &H8000000F
Sec3Item2.ForeColor = &H8000000F
Sec3Item3.ForeColor = &H8000000F
Sec3Item4.ForeColor = &H8000000F
Sec3Item5.ForeColor = vbWhite
End Sub
'Section 4 Items
Private Sub Sec4Item1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = vbWhite
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = vbWhite
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = vbWhite
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = vbWhite
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = vbWhite
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = vbWhite
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = vbWhite
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = vbWhite
Sec4Item9.ForeColor = &H8000000F
End Sub
Private Sub Sec4Item9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Sec4Item1.ForeColor = &H8000000F
Sec4Item2.ForeColor = &H8000000F
Sec4Item3.ForeColor = &H8000000F
Sec4Item4.ForeColor = &H8000000F
Sec4Item5.ForeColor = &H8000000F
Sec4Item6.ForeColor = &H8000000F
Sec4Item7.ForeColor = &H8000000F
Sec4Item8.ForeColor = &H8000000F
Sec4Item9.ForeColor = vbWhite
End Sub
'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Item Hover Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////


'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     Start Item Run App Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////
'On each click we will load the file path from the INI
'And then use it with the shell command

'Items in Section 1
Private Sub Sec1Item1_Click()
Sec1Item1Path = GetFromINI("AppPath", "Sec1Item1", App.path & "\CipherOS.ini")
Shell Sec1Item1Path, vbNormalFocus
End Sub
Private Sub Sec1Item2_Click()
Sec1Item2Path = GetFromINI("AppPath", "Sec1Item2", App.path & "\CipherOS.ini")
Shell Sec1Item2Path, vbNormalFocus
End Sub
Private Sub Sec1Item3_Click()
Sec1Item3Path = GetFromINI("AppPath", "Sec1Item3", App.path & "\CipherOS.ini")
Shell Sec1Item3Path, vbNormalFocus
End Sub
Private Sub Sec1Item4_Click()
Sec1Item4Path = GetFromINI("AppPath", "Sec1Item4", App.path & "\CipherOS.ini")
Shell Sec1Item4Path, vbNormalFocus
End Sub
Private Sub Sec1Item5_Click()
Sec1Item5Path = GetFromINI("AppPath", "Sec1Item5", App.path & "\CipherOS.ini")
Shell Sec1Item5Path, vbNormalFocus
End Sub
Private Sub Sec1Item6_Click()
Sec1Item6Path = GetFromINI("AppPath", "Sec1Item6", App.path & "\CipherOS.ini")
Shell Sec1Item6Path, vbNormalFocus
End Sub
'Items in Section 2
Private Sub Sec2Item1_Click()
Sec2Item1Path = GetFromINI("AppPath", "Sec2Item1", App.path & "\CipherOS.ini")
Shell Sec2Item1Path, vbNormalFocus
End Sub
Private Sub Sec2Item2_Click()
Sec2Item2Path = GetFromINI("AppPath", "Sec2Item2", App.path & "\CipherOS.ini")
Shell Sec2Item2Path, vbNormalFocus
End Sub
Private Sub Sec2Item3_Click()
Sec2Item3Path = GetFromINI("AppPath", "Sec2Item3", App.path & "\CipherOS.ini")
Shell Sec2Item3Path, vbNormalFocus
End Sub
Private Sub Sec2Item4_Click()
Sec2Item4Path = GetFromINI("AppPath", "Sec2Item4", App.path & "\CipherOS.ini")
Shell Sec2Item4Path, vbNormalFocus
End Sub
'Items in Section 3
Private Sub Sec3Item1_Click()
Sec3Item1Path = GetFromINI("AppPath", "Sec3Item1", App.path & "\CipherOS.ini")
Shell Sec3Item1Path, vbNormalFocus
End Sub
Private Sub Sec3Item2_Click()
Sec3Item2Path = GetFromINI("AppPath", "Sec3Item2", App.path & "\CipherOS.ini")
Shell Sec3Item2Path, vbNormalFocus
End Sub
Private Sub Sec3Item3_Click()
Sec3Item3Path = GetFromINI("AppPath", "Sec3Item3", App.path & "\CipherOS.ini")
Shell Sec3Item3Path, vbNormalFocus
End Sub
Private Sub Sec3Item4_Click()
Sec3Item4Path = GetFromINI("AppPath", "Sec3Item4", App.path & "\CipherOS.ini")
Shell Sec3Item4Path, vbNormalFocus
End Sub
Private Sub Sec3Item5_Click()
Sec3Item5Path = GetFromINI("AppPath", "Sec3Item5", App.path & "\CipherOS.ini")
Shell Sec3Item5Path, vbNormalFocus
End Sub
'Item in Section 4
Private Sub Sec4Item1_Click()
Sec4Item1Path = GetFromINI("AppPath", "Sec4Item1", App.path & "\CipherOS.ini")
Shell Sec4Item1Path, vbNormalFocus
End Sub
Private Sub Sec4Item2_Click()
Sec4Item2Path = GetFromINI("AppPath", "Sec4Item2", App.path & "\CipherOS.ini")
Shell Sec4Item2Path, vbNormalFocus
End Sub
Private Sub Sec4Item3_Click()
Sec4Item3Path = GetFromINI("AppPath", "Sec4Item3", App.path & "\CipherOS.ini")
Shell Sec4Item3Path, vbNormalFocus
End Sub
Private Sub Sec4Item4_Click()
Sec4Item4Path = GetFromINI("AppPath", "Sec4Item4", App.path & "\CipherOS.ini")
Shell Sec4Item4Path, vbNormalFocus
End Sub
Private Sub Sec4Item5_Click()
Sec4Item5Path = GetFromINI("AppPath", "Sec4Item5", App.path & "\CipherOS.ini")
Shell Sec4Item5Path, vbNormalFocus
End Sub
Private Sub Sec4Item6_Click()
Sec4Item6Path = GetFromINI("AppPath", "Sec4Item6", App.path & "\CipherOS.ini")
Shell Sec4Item6Path, vbNormalFocus
End Sub
Private Sub Sec4Item7_Click()
Sec4Item7Path = GetFromINI("AppPath", "Sec4Item7", App.path & "\CipherOS.ini")
Shell Sec4Item7Path, vbNormalFocus
End Sub
Private Sub Sec4Item8_Click()
Sec4Item8Path = GetFromINI("AppPath", "Sec4Item8", App.path & "\CipherOS.ini")
Shell Sec4Item8Path, vbNormalFocus
End Sub
Private Sub Sec4Item9_Click()
Sec4Item9Path = GetFromINI("AppPath", "Sec4Item9", App.path & "\CipherOS.ini")
Shell Sec4Item9Path, vbNormalFocus
End Sub
'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Item Run App Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////

'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     Start Get Item Names
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////

Public Sub GetItemNames()
'Get Sections Names from the INI and apply them
    Sec1Name$ = GetFromINI("SecNames", "Section1Name", App.path & "\CipherOS.ini")
    Sec2Name$ = GetFromINI("SecNames", "Section2Name", App.path & "\CipherOS.ini")
    Sec3Name$ = GetFromINI("SecNames", "Section3Name", App.path & "\CipherOS.ini")
    Sec4Name$ = GetFromINI("SecNames", "Section4Name", App.path & "\CipherOS.ini")
    Section1.Caption = Sec1Name$
    Section2.Caption = Sec2Name$
    Section3.Caption = Sec3Name$
    Section4.Caption = Sec4Name$
    'Section Names have been applied
    
    'Get Item Names from the INI and apply them
    Sec1Item1Name$ = GetFromINI("AppNames", "Sec1Item1", App.path & "\CipherOS.ini")
    Sec1Item2Name$ = GetFromINI("AppNames", "Sec1Item2", App.path & "\CipherOS.ini")
    Sec1Item3Name$ = GetFromINI("AppNames", "Sec1Item3", App.path & "\CipherOS.ini")
    Sec1Item4Name$ = GetFromINI("AppNames", "Sec1Item4", App.path & "\CipherOS.ini")
    Sec1Item5Name$ = GetFromINI("AppNames", "Sec1Item5", App.path & "\CipherOS.ini")
    Sec1Item6Name$ = GetFromINI("AppNames", "Sec1Item6", App.path & "\CipherOS.ini")
    Sec2Item1Name$ = GetFromINI("AppNames", "Sec2Item1", App.path & "\CipherOS.ini")
    Sec2Item2Name$ = GetFromINI("AppNames", "Sec2Item2", App.path & "\CipherOS.ini")
    Sec2Item3Name$ = GetFromINI("AppNames", "Sec2Item3", App.path & "\CipherOS.ini")
    Sec2Item4Name$ = GetFromINI("AppNames", "Sec2Item4", App.path & "\CipherOS.ini")
    Sec3Item1Name$ = GetFromINI("AppNames", "Sec3Item1", App.path & "\CipherOS.ini")
    Sec3Item2Name$ = GetFromINI("AppNames", "Sec3Item2", App.path & "\CipherOS.ini")
    Sec3Item3Name$ = GetFromINI("AppNames", "Sec3Item3", App.path & "\CipherOS.ini")
    Sec3Item4Name$ = GetFromINI("AppNames", "Sec3Item4", App.path & "\CipherOS.ini")
    Sec3Item5Name$ = GetFromINI("AppNames", "Sec3Item5", App.path & "\CipherOS.ini")
    Sec4Item1Name$ = GetFromINI("AppNames", "Sec4Item1", App.path & "\CipherOS.ini")
    Sec4Item2Name$ = GetFromINI("AppNames", "Sec4Item2", App.path & "\CipherOS.ini")
    Sec4Item3Name$ = GetFromINI("AppNames", "Sec4Item3", App.path & "\CipherOS.ini")
    Sec4Item4Name$ = GetFromINI("AppNames", "Sec4Item4", App.path & "\CipherOS.ini")
    Sec4Item5Name$ = GetFromINI("AppNames", "Sec4Item5", App.path & "\CipherOS.ini")
    Sec4Item6Name$ = GetFromINI("AppNames", "Sec4Item6", App.path & "\CipherOS.ini")
    Sec4Item7Name$ = GetFromINI("AppNames", "Sec4Item7", App.path & "\CipherOS.ini")
    Sec4Item8Name$ = GetFromINI("AppNames", "Sec4Item8", App.path & "\CipherOS.ini")
    Sec4Item9Name$ = GetFromINI("AppNames", "Sec4Item9", App.path & "\CipherOS.ini")
    Sec1Item1.Caption = Sec1Item1Name$
    Sec1Item2.Caption = Sec1Item2Name$
    Sec1Item3.Caption = Sec1Item3Name$
    Sec1Item4.Caption = Sec1Item4Name$
    Sec1Item5.Caption = Sec1Item5Name$
    Sec1Item6.Caption = Sec1Item6Name$
    Sec2Item1.Caption = Sec2Item1Name$
    Sec2Item2.Caption = Sec2Item2Name$
    Sec2Item3.Caption = Sec2Item3Name$
    Sec2Item4.Caption = Sec2Item4Name$
    Sec3Item1.Caption = Sec3Item1Name$
    Sec3Item2.Caption = Sec3Item2Name$
    Sec3Item3.Caption = Sec3Item3Name$
    Sec3Item4.Caption = Sec3Item4Name$
    Sec3Item5.Caption = Sec3Item5Name$
    Sec4Item1.Caption = Sec4Item1Name$
    Sec4Item2.Caption = Sec4Item2Name$
    Sec4Item3.Caption = Sec4Item3Name$
    Sec4Item4.Caption = Sec4Item4Name$
    Sec4Item5.Caption = Sec4Item5Name$
    Sec4Item6.Caption = Sec4Item6Name$
    Sec4Item7.Caption = Sec4Item7Name$
    Sec4Item8.Caption = Sec4Item8Name$
    Sec4Item9.Caption = Sec4Item9Name$
    'Item Names have be applied
    'And we're all done with the INI Names
End Sub
'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Get Item Names
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////


'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     Start Apply Skin
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////
Public Sub ApplySkin()
SkinPath$ = GetFromINI("Skins", "SkinFolder", App.path & "\CipherOS.ini")
    Me.Picture = LoadPicture(SkinPath$ & "\Background.jpg")           'Apply Main Skin
    For i = 0 To 4
    imgTop(i).Picture = LoadPicture(SkinPath$ & "\link_top.gif")      'Apply Top Buttons
    Next i
    For i = 0 To 25
    On Error Resume Next
    imgLinkMid(i).Picture = LoadPicture(SkinPath$ & "\link_mid.gif")  'Apply Middle Buttons
    Next i
    For i = 0 To 12
    imgRight(i).Picture = LoadPicture(SkinPath$ & "\link_right.gif")  'Apply Right Buttons
    Next i
    For i = 0 To 5
    imgTitle(i).Picture = LoadPicture(SkinPath$ & "\link_title.gif")  'Apply Title Buttons
    Next i
    imgControlPanel.Picture = LoadPicture(SkinPath$ & "\BackgroundDark.jpg")
    imgChangeSkin.Picture = LoadPicture(SkinPath$ & "\BackgroundDark.jpg")
    imgSHellSwitch.Picture = LoadPicture(SkinPath$ & "\BackgroundDark.jpg")
    imgItemValue.Picture = LoadPicture(SkinPath$ & "\BackgroundDark.jpg")
    imgShuttingDown.Picture = LoadPicture(SkinPath$ & "\BackgroundDark.jpg")
    frmProgramsMenu.picItem(0).Picture = LoadPicture(SkinPath$ & "\BackgroundDark.jpg")
    btnRunFile.Picture = LoadPicture(SkinPath$ & "\RunButton.bmp")
    btnDone.Picture = LoadPicture(SkinPath$ & "\RunButton.bmp")
    btnCancel.Picture = LoadPicture(SkinPath$ & "\RunButton.bmp")
    btnApplySkin.Picture = LoadPicture(SkinPath$ & "\RunButton.bmp")
    btnCancelSkin.Picture = LoadPicture(SkinPath$ & "\RunButton.bmp")
End Sub
'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Apply Skin
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////



'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     Start Main Form Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////

Private Sub Form_Click()
Unload frmProgramsMenu
End Sub

Private Sub Form_Load()

    Shell "c:\windows\system\systray.exe" 'make this work as a shell
    'set time
    lblTimeDate.Caption = Time & "                         " & Date
    TimerDate.Enabled = True
    'end set time

    'Set frmShuttingDown in middle
    frmShuttingDown.Left = 5520
    frmShuttingDown.Top = 5280
    'End Set frmShuttinDown in middle

    'taskbar load
    Call WhichWindows(lstTasks)
    On Error Resume Next
    lstTasks.ListIndex = 0
    'taskbar loaded
    
    Call GetItemNames  'Get the names of every item
    Call ApplySkin     'Apply the current skin
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ButtonsAllGray
End Sub


Private Sub lblDriveA_Click()
Shell ("explorer A:\"), vbNormalFocus
End Sub

Private Sub lblDriveC_Click()
Shell ("explorer C:\"), vbNormalFocus
End Sub

Private Sub lblDriveD_Click()
Shell ("explorer D:\"), vbNormalFocus
End Sub


Private Sub lblPrograms_Click()
frmProgramsMenu.GetMenu "C:\Windows\Start Menu\Programs"
frmProgramsMenu.Show
End Sub

Private Sub lblRestart_Click()
lblShuttingDown.Caption = "CipherOS is closing. Your computer is now restarting. Please be patient..."
frmShuttingDown.Visible = True
Restart
End Sub

Private Sub lblRunFile_Click()
On Error Resume Next
If LCase$(txtRunFile.Text) = "exit" Then 'Exit Application
 Unload Me
 End
ElseIf LCase$(txtRunFile.Text) = "name" Then 'Show Item Value Frame
    On Error Resume Next
    frameControlPanel.Visible = False
    frameShellSwitcher.Visible = False
    frmItemValue.Visible = True
Else
On Error Resume Next
ShellExecute Me.hWnd, "open", txtRunFile.Text, "", "", 1
End If
End Sub

Private Sub lblRunFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRunFile.ForeColor = vbWhite
End Sub

Private Sub lstTasks_dblClick()
    If lstTasks.ListCount = 0 Then Exit Sub
    Call pSetForegroundWindow(lstTasks.ItemData(lstTasks.ListIndex))
End Sub

Private Sub tmrTask_Timer()
On Error Resume Next
    Call WhichWindows(lstTasks)
    On Error Resume Next
    lstTasks.ListIndex = 0
End Sub

Private Sub txtNewValue_DblClick()
dlgBrowseValue.ShowOpen
txtNewValue.Text = dlgBrowseValue.FileName
End Sub

Private Sub txtRunFile_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCrLf) Then
lblRunFile_Click
End If
End Sub

Private Sub lblShellSwitcher_Click()
frameControlPanel.Visible = False
frameShellSwitcher.Visible = True
frmItemValue.Visible = False
End Sub

Private Sub lblShSwitch_Click()
frameControlPanel.Visible = False
frameShellSwitcher.Visible = True
End Sub

Private Sub lblShutDown_Click()
lblShuttingDown.Caption = "CipherOS is closing. Your computer is now shutting down. Please be patient..."
frmShuttingDown.Visible = True
ShutDown
End Sub
Private Sub cmbAllItems_Click()
On Error Resume Next
NewSecName = GetFromINI("SecNames", cmbAllItems.Text, App.path & "\CipherOS.ini")
NewItemName = GetFromINI("AppNames", cmbAllItems.Text, App.path & "\CipherOS.ini")
NewItemPath = GetFromINI("AppPath", cmbAllItems.Text, App.path & "\CipherOS.ini")
txtNewName.Text = NewItemName & NewSecName
txtNewValue.Text = NewItemPath
End Sub
Private Sub lblControlPanel_Click()
frameControlPanel.Visible = True
frameShellSwitcher.Visible = False
frmItemValue.Visible = False
End Sub
Private Sub lblCancel_Click()
frmItemValue.Visible = False
End Sub

Private Sub lblDone_Click()
'If it's a Section then write the new name
    If cmbAllItems.Text = "Section1Name" Then
        Call WriteToINI("SecNames", cmbAllItems.Text, txtNewName.Text, App.path & "\CipherOS.ini")
    ElseIf cmbAllItems.Text = "Section2Name" Then
        Call WriteToINI("SecNames", cmbAllItems.Text, txtNewName.Text, App.path & "\CipherOS.ini")
    ElseIf cmbAllItems.Text = "Section3Name" Then
        Call WriteToINI("SecNames", cmbAllItems.Text, txtNewName.Text, App.path & "\CipherOS.ini")
    ElseIf cmbAllItems.Text = "Section4Name" Then
        Call WriteToINI("SecNames", cmbAllItems.Text, txtNewName.Text, App.path & "\CipherOS.ini")
    Else 'If it's an Item (button) write the name and value
        Call WriteToINI("AppNames", cmbAllItems.Text, txtNewName.Text, App.path & "\CipherOS.ini")
        Call WriteToINI("AppPath", cmbAllItems.Text, txtNewValue.Text, App.path & "\CipherOS.ini")
End If
frmItemValue.Visible = False
Call GetItemNames 'refresh the item names
cmbAllItems.ListIndex = 0 'back to the primitive
txtNewName = ""       'back to the primitive
txtNewValue = ""      'back to the primitive
End Sub

Private Sub lblTimeDate_DblClick()
Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
End Sub

Private Sub TimerDate_Timer()
lblTimeDate.Caption = Time & "                         " & Date
End Sub

'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Main Form Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////



'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'    Begin Control Panel Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////

Private Sub lblAccessability_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL access.cpl,,5", vbNormalFocus
End Sub

Private Sub lblAddHardware_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1"
End Sub

Private Sub lblAddRemProm_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1"
End Sub

Private Sub lblDateTime_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
End Sub

Private Sub lblDisplay_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3"
End Sub

Private Sub lblJoystick_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL joy.cpl"
End Sub

Private Sub lblKeyBoard_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL main.cpl @1"
End Sub

Private Sub lblInternetOpt_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0"
End Sub

Private Sub lblMailFax_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL mlcfg32.cpl"
End Sub

Private Sub lblModem_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL modem.cpl"
End Sub

Private Sub lblMouse_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL main.cpl @0"
End Sub

Private Sub lblNetwork_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl"
End Sub

Private Sub lblPassword_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL password.cpl"
End Sub

Private Sub lblPrinter_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL main.cpl @2"
End Sub

Private Sub lblRegional_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0"
End Sub
Private Sub lblSystem_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0"
End Sub

'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Control Panel Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////




'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     Start Shell Switcher Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////
Private Sub lblShellCancel_Click()
frameShellSwitcher.Visible = False
End Sub

Private Sub lblShellCipherOS_Click()
Call WriteToINI("boot", "shell", App.path & "\CipherOS.exe", "C:\windows\system.ini")
MsgBox "CipherOS has been made your default Shell", vbExclamation
frameShellSwitcher.Visible = False
End Sub

Private Sub lblShellExplorer_Click()
Call WriteToINI("boot", "shell", "explorer.exe", "C:\windows\system.ini")
MsgBox "Explorer has been made your default Shell" & vbNewLine & "Please restart your computer now.", vbExclamation
frameShellSwitcher.Visible = False
End Sub
'///////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     End Shell Switcher Codes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////
