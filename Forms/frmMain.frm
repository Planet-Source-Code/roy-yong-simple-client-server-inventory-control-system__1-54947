VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main Console"
   ClientHeight    =   7815
   ClientLeft      =   -120
   ClientTop       =   660
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   5
      ToolTipText     =   "Click here to search for your keywords based on the search criteria."
      Top             =   5400
      Width           =   1335
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   11280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "bar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7E
            Key             =   "guy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2358
            Key             =   "trolley"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C32
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":350C
            Key             =   "app"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DE6
            Key             =   "right"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46C0
            Key             =   "line"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F9A
            Key             =   "exclaimation"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5874
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":614E
            Key             =   "db"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A28
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7302
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BDC
            Key             =   "gng"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":84B6
            Key             =   "key"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D90
            Key             =   "arrows"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":966A
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F44
            Key             =   "magnifier"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A81E
            Key             =   "synon"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0F8
            Key             =   "people"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9D2
            Key             =   "silverlock"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C2AC
            Key             =   "server"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB86
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D460
            Key             =   "plus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   10680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD3A
            Key             =   "app"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E2D4
            Key             =   "server"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E86E
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EE08
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F3A2
            Key             =   "gng"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F93C
            Key             =   "exclaimation"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FED6
            Key             =   "synon"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10470
            Key             =   "arrows"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1300A
            Key             =   "magnifier"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":135A4
            Key             =   "trolley"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search:"
      Height          =   1575
      Left            =   4920
      TabIndex        =   2
      Top             =   4320
      Width           =   6735
      Begin VB.ComboBox cmbCondition 
         Height          =   315
         ItemData        =   "frmMain.frx":13B3E
         Left            =   5280
         List            =   "frmMain.frx":13B4B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search Criteria:"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "To search for something, select the search criteria and enter your keywords. Then, click on 'Search'."
         Height          =   615
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":13B78
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame frmShortcuts 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   4575
      Begin VB.Image Image2 
         Height          =   495
         Index           =   4
         Left            =   120
         MouseIcon       =   "frmMain.frx":14442
         MousePointer    =   99  'Custom
         ToolTipText     =   "Control users allowed to access this system"
         Top             =   2400
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   3
         Left            =   120
         MouseIcon       =   "frmMain.frx":1474C
         MousePointer    =   99  'Custom
         ToolTipText     =   "See who your colleagues are"
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   2
         Left            =   120
         MouseIcon       =   "frmMain.frx":14A56
         MousePointer    =   99  'Custom
         ToolTipText     =   "Check stock level by browsing the Inventory"
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmMain.frx":14D60
         MousePointer    =   99  'Custom
         ToolTipText     =   "Create a new Purchase Order"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmMain.frx":1506A
         MousePointer    =   99  'Custom
         ToolTipText     =   "Create a new Delivery Order"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick A Task"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label receivables 
         BackStyle       =   0  'Transparent
         Caption         =   "Create a new Delivery Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmMain.frx":15374
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Tag             =   "0"
         ToolTipText     =   "press CTRL + d"
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label po 
         BackStyle       =   0  'Transparent
         Caption         =   "Create a new Purchase Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmMain.frx":1567E
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Tag             =   "1"
         ToolTipText     =   "press CTRL + p"
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label inventory 
         BackStyle       =   0  'Transparent
         Caption         =   "Check stock level by browsing the Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmMain.frx":15988
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Tag             =   "2"
         ToolTipText     =   "press CTRL + i"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label human 
         BackStyle       =   0  'Transparent
         Caption         =   "See who your colleagues are"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmMain.frx":15C92
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Tag             =   "3"
         ToolTipText     =   "press CTRL + e"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label users 
         BackStyle       =   0  'Transparent
         Caption         =   "Control users allowed to access this system"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmMain.frx":15F9C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Tag             =   "4"
         ToolTipText     =   "press CTRL + u"
         Top             =   2520
         Width           =   3855
      End
   End
   Begin VB.Timer tmrTicker 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9600
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recent Deliveries:"
      Height          =   3135
      Left            =   4920
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
      Begin MSComctlLib.ListView lvDeliveries 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   10080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmTick 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "News:"
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9495
      Begin VB.Label lblTicker 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9255
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Shape bgWelcome 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4935
      Left            =   0
      Top             =   960
      Width           =   4815
   End
   Begin VB.Menu mnu_Main_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Options_Refresh 
         Caption         =   "&Refresh Recent Deliveries"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_Option_Password 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnu_Bar_12 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Personal_Ticker 
         Caption         =   "&Add Personal Ticker"
      End
      Begin VB.Menu mnu_Manage_Ticker 
         Caption         =   "&Manage My Tickers..."
      End
      Begin VB.Menu mnu_Bar_05 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options_Quit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_Main_Admin 
      Caption         =   "&Admin"
      Begin VB.Menu mnu_Admin_Database 
         Caption         =   "&Database"
         Begin VB.Menu mnu_Database_Backup 
            Caption         =   "&Backup"
         End
         Begin VB.Menu mnu_Database_Restore 
            Caption         =   "&Restore"
         End
      End
      Begin VB.Menu mnu_Bar_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Admin_Users 
         Caption         =   "&Users"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu_Bar_03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Admin_Settings 
         Caption         =   "Se&ttings..."
      End
      Begin VB.Menu mnu_Admin_SQL 
         Caption         =   "&SQL Console"
      End
      Begin VB.Menu mnu_Admin_Logs 
         Caption         =   "&View Logs"
      End
      Begin VB.Menu mnu_Bar_11 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Add_PublicTicker 
         Caption         =   "&Add Public Ticker"
      End
   End
   Begin VB.Menu mnu_Admin_Maintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnu_Maintenance_Countries 
         Caption         =   "&Countries"
      End
      Begin VB.Menu mnu_Maintenance_States 
         Caption         =   "&States"
      End
      Begin VB.Menu mnu_Maintenance_Cities 
         Caption         =   "C&ities"
      End
   End
   Begin VB.Menu mnu_Main_Receivable 
      Caption         =   "&Accounts Receivable"
      Begin VB.Menu mnu_Receivable_Add 
         Caption         =   "&Add New Customer"
      End
      Begin VB.Menu mnu_Receivable_Customers 
         Caption         =   "&Customers..."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_Bar_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Receivable_View 
         Caption         =   "&View Delivery Orders"
      End
   End
   Begin VB.Menu mnu_Main_Payable 
      Caption         =   "Accounts &Payable"
      Begin VB.Menu mnu_Payable_Add 
         Caption         =   "&Add New Supplier"
      End
      Begin VB.Menu mnu_Payable_Suppliers 
         Caption         =   "&Suppliers..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Bar_06 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Payable_View 
         Caption         =   "&View Purchase Orders"
      End
   End
   Begin VB.Menu mnu_Main_Employees 
      Caption         =   "&Human Resource"
      Begin VB.Menu mnu_Human_Employees 
         Caption         =   "&Employees..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_Bar_07 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Human_Payrol 
         Caption         =   "&New Payroll"
      End
      Begin VB.Menu mnu_Bar_08 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Human_Leave 
         Caption         =   "&Apply For Leave"
      End
      Begin VB.Menu mnu_Human_Browse 
         Caption         =   "&Browse Leaves"
      End
   End
   Begin VB.Menu mnu_Main_Inventory 
      Caption         =   "&Inventory"
      Begin VB.Menu mnu_Inventory_Browse 
         Caption         =   "&Browse Inventory"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_Bar_04 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Inventory_Categories 
         Caption         =   "Categories..."
      End
      Begin VB.Menu mnu_Bar_09 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Inventory_Contracts 
         Caption         =   "&Consignment Contracts"
      End
      Begin VB.Menu mnu_Inventory_View 
         Caption         =   "Browse Consignment Inventory..."
         Begin VB.Menu mnu_Contracts 
            Caption         =   "--{NONE}--"
            Index           =   0
         End
      End
      Begin VB.Menu mnu_Bar_13 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Receivable_Delivery 
         Caption         =   "&New Delivery Order"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_Payable_Order 
         Caption         =   "&New Purchase Order"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnu_Main_Invoice 
      Caption         =   "&Invoicing"
      Begin VB.Menu mnu_Invoicing 
         Caption         =   "&DO Received..."
      End
   End
   Begin VB.Menu mnu_Main_Report 
      Caption         =   "&Report"
      Begin VB.Menu mnu_Report_Main 
         Caption         =   "&Main Report"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnu_Main_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_Help 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnu_Bar_14 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnu_Main_Print 
      Caption         =   "P&rint"
      Visible         =   0   'False
      Begin VB.Menu mnu_Print_Print 
         Caption         =   "&Print..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Integer
Dim prevHour As Byte
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub loadTicker()
Dim tickSQL As String
tickSQL = " SELECT msgTitle,msgText FROM Ticker WHERE ((username='" & CurrentUser.strUsername & "' Or username='GENERAL') AND ((dateToBeShown Is Null) Or dateToBeShown='" & Format$(Now(), "dd/mm/yyyy") & "'));"

Dim tickRS As Recordset
RSOpen tickRS, tickSQL, dbOpenSnapshot
'First load an array of labels
Dim i As Integer
'Assign recordset to labels
i = 0
While Not tickRS.EOF
    On Error Resume Next
    lblTicker(i).Container = frmTick
    lblTicker(i).Visible = False
    lblTicker(i).Caption = tickRS("msgTitle") & vbCrLf & tickRS("msgText")
    i = i + 1
    Load lblTicker(i)
    tickRS.MoveNext
Wend
tickRS.Close
Set tickRS = Nothing
tmrTicker.Enabled = True
c = 0
End Sub

Private Sub destroyTicker()
Dim i As Integer
For i = lblTicker.LBound To lblTicker.UBound
    'Suppose to free memory
    If i <> 0 Then
        Unload lblTicker(i)
    End If
Next i
End Sub

Private Sub tickerStart()
'Starts the news ticker
'Set the first ticker visible and position on top of frame
lblTicker(c).Visible = True
lblTicker(c).Top = frmTick.height
lblTicker(c).Left = 120
tmrTicker.Enabled = True
End Sub

Private Sub tickerStop()
'Stops the news ticker
tmrTicker.Enabled = False
destroyTicker
End Sub

Private Sub moveTicker(ByVal amt As Integer)
lblTicker(c).ZOrder vbBringToFront
lblTicker(c).Visible = True
lblTicker(c).Top = lblTicker(c).Top - amt

If lblTicker(c).Top < 0 - lblTicker(c).height Then
    'hide the current lbl
    lblTicker(c).Visible = False
    c = c + 1
    If c > lblTicker.UBound Then
        c = 0
    End If
    tickerStart
End If
End Sub

Private Sub checkAccStatus() 'Checks user for password expiry
Dim expDuration As Long
Debug.Print "User last: " & CurrentUser.lastPassword
expDuration = DateDiff("d", Now(), Format$(CurrentUser.lastPassword, "mm/dd/yyyy"))
Debug.Print "Duration: " & expDuration
If CurrentUser.mustChange = True Then
    InfoMsg "Your account password has expired. You are required to change your password now.", "Password expired"
    Load frmPassword
    frmPassword.setFormMode = force_Change
    frmPassword.Show vbModal
    DoEvents
Else
    If expDuration < -60 Then
        ValidMsg "Your password has expired. You will be required to change your password now.", "Password expired"
        Load frmPassword
        frmPassword.setFormMode = force_Change
        frmPassword.Show vbModal
    ElseIf ((expDuration < -45) And (expDuration > -61)) Then
        If MsgBox("Your password is expiring in " & expDuration & " days. Do you wish to change your password now?", vbYesNo + vbQuestion, "Password expiring") = vbYes Then
            Load frmPassword
            frmPassword.Show vbModal
        End If
    End If
End If
End Sub

Private Sub implementSystemPolicy()
'Hides or show menus based on user's privileges
With CurrentUser
    If .prvlgAdmin = True Then
        mnu_Main_Admin.Visible = True
        users.Visible = True
        Image2(users.Tag).Visible = True
    Else
        mnu_Main_Admin.Visible = False
        users.Visible = False
        Image2(users.Tag).Visible = False
    End If
    If .prvlgAPS = True Then
        mnu_Main_Payable.Visible = True
        po.Visible = True
        Image2(po.Tag).Visible = True
    Else
        mnu_Main_Payable.Visible = False
        po.Visible = False
        Image2(po.Tag).Visible = False

    End If
    If .prvlgARS = True Then
        mnu_Main_Receivable.Visible = True
        mnu_Main_Invoice.Visible = True
    Else
        mnu_Main_Receivable.Visible = False
        mnu_Main_Invoice.Visible = False
    End If
    If .prvlgDOS = True Then
        mnu_Receivable_Delivery.Visible = True
        receivables.Visible = True
        Image2(receivables.Tag).Visible = True
    Else
        mnu_Receivable_Delivery.Visible = False
        receivables.Visible = False
        Image2(receivables.Tag).Visible = False
    End If
    If .prvlgHRS = True Then
        mnu_Main_Employees.Visible = True
        human.Visible = True
        Image2(human.Tag).Visible = True
    Else
        mnu_Main_Employees.Visible = False
        human.Visible = False
        Image2(human.Tag).Visible = False
    End If
    If .prvlgReport = True Then
        mnu_Main_Report.Visible = True
    Else
        mnu_Main_Report.Visible = False
    End If
End With
End Sub

Public Sub loadRecDeliveries()
'loads the 10 most recent deliveries
With lvDeliveries
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "DO No."
    .ColumnHeaders.add , , "Customer"
    .ColumnHeaders(2).width = 2000
    .ColumnHeaders.add , , "Date"
    .ColumnHeaders.add , , "Status"
    
    Dim recentRS As Recordset, recentSQL As String
    recentSQL = "SELECT Delivery.DOnumber, Delivery.Date, Customers.Name, Delivery.Status " & _
                "FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID " & _
                "WHERE Delivery.Status='DELIVERING' ORDER BY Delivery.Date DESC"
    On Error GoTo ErrHandler
    .ListItems.Clear
    RSOpen recentRS, recentSQL, dbOpenSnapshot
    While Not recentRS.EOF
        .ListItems.add , , recentRS("DOnumber")
        .ListItems(.ListItems.Count).SubItems(1) = recentRS("Name")
        .ListItems(.ListItems.Count).SubItems(2) = recentRS("Date")
        .ListItems(.ListItems.Count).SubItems(3) = recentRS("Status")
        recentRS.MoveNext
    Wend
    recentRS.Close
    Set recentRS = Nothing
End With
ErrHandler:
If Err.Number <> 0 Then
    lvDeliveries.ListItems.add , , "ERROR"
    lvDeliveries.ListItems(lvDeliveries.ListItems.Count).SubItems(1) = "AN ERROR HAS OCCURED. UNABLE TO LOAD RECENT DELIVERIES"
    Exit Sub
End If
End Sub

Private Sub performSearch()
Dim strCriteria As String, searchSQL As String
Dim criteriaNum As Integer, NumRecords As Integer
Dim searchRS As Recordset
criteriaNum = 0
strCriteria = txtSearch.Text
Select Case cmbCondition.Text
    Case "Delivery Order"
        criteriaNum = 1
        searchSQL = "SELECT Delivery.DOnumber, Customers.Name, Delivery.PONumber, Delivery.DelDate, Delivery.DelTime, Delivery.Status FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE ((DOnumber LIKE '*" & strCriteria & "*') OR (PONumber LIKE '*" & strCriteria & "*') OR (Attn LIKE '*" & strCriteria & "*') OR (Remark LIKE '*" & strCriteria & "*'));"
    Case "Purchase Order"
        criteriaNum = 2
        searchSQL = "SELECT Purchase.poNumber, Suppliers.Name, Purchase.Date FROM Suppliers INNER JOIN Purchase ON Suppliers.SupplierID=Purchase.SupplierID WHERE (Purchase.poNumber LIKE '" & strCriteria & "') OR (Suppliers.Name LIKE '" & strCriteria & "') OR (Purchase.Date LIKE '" & strCriteria & "');"
    Case "Product"
        criteriaNum = 3
        searchSQL = "SELECT * FROM Products WHERE ((ProductID LIKE '*" & strCriteria & "*') OR (Description LIKE '*" & strCriteria & "*') OR (Brand LIKE '*" & strCriteria & "*') OR (CategoryID LIKE '*" & strCriteria & "*'));"
End Select
If criteriaNum > 0 Then
    'Open recordset to see if query returns result
    RSOpen searchRS, searchSQL, dbOpenSnapshot
    'Proceed if not end of file
    If Not searchRS.EOF Then
        searchRS.MoveLast
        NumRecords = searchRS.RecordCount
        If criteriaNum = 1 Then
            frmDelivery_Main.Show , frmMain
            frmDelivery_Main.getDeliveries searchSQL
        ElseIf criteriaNum = 3 Then
            frmProduct_Browse.Show , frmMain
            frmProduct_Browse.showListofProducts searchSQL
        End If
        InfoMsg "Search completed. " & NumRecords & " matching record(s) found.", "Search complete"
    Else
        InfoMsg "No matching record found based on search string provided. Please try again.", "No record found"
        txtSearch.SetFocus
    End If
    searchRS.Close
    Set searchRS = Nothing
Else
    ValidMsg "Please specify a search criteria.", "Missing criteria"
End If
End Sub

Private Sub Command1_Click()
If Len(txtSearch.Text) > 0 Then
    performSearch
Else
    ValidMsg "Please enter a keyword for search. Example: Bearing A222", "Missing keyword"
    txtSearch.SetFocus
End If
End Sub

Private Sub Form_Load()
prevHour = Hour(Now())
lblWelcome.Caption = "Welcome " & CurrentUser.strUsername
checkAccStatus
implementSystemPolicy
loadRecDeliveries
loadTicker
Dim i As Integer
For i = 0 To 4
    Set Image2(i).Picture = img32.ListImages(7).Picture
Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("All windows will be closed and unsaved delivery orders, etc would be destroyed." & vbCrLf & "Are you sure you want to quit now? ", vbYesNoCancel + vbQuestion, "Quit") <> vbYes Then
    Cancel = True
End If
End Sub

Private Sub Form_Resize()
frmTick.width = Me.width
bgWelcome.height = Me.height
End Sub

Private Sub Form_Unload(Cancel As Integer)
tickerStop
Dim tempSQL As String
'Because users might just click on x button to close form
With MySynonDatabase
    'Begin writting of system logging here.
    insertLog "User logged off"
    
    tempSQL = "UPDATE Users SET Users.status = 'OFFLINE' " & _
            "WHERE (([Username]='" & CurrentUser.strUsername & "'));"
    .Execute tempSQL
End With
closeDB
Set frmMain = Nothing
End 'Ends the process of the application
End Sub

Private Sub lvDeliveries_DblClick()
With lvDeliveries
If .ListItems.Count > 0 Then
    If .SelectedItem.Selected Then
        frmDelivery_Main.Show
        frmDelivery_Main.getDeliveries "SELECT Delivery.DOnumber, Customers.Name, Delivery.PONumber, Delivery.DelDate, Delivery.DelTime, Delivery.Status FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE DOnumber='" & .SelectedItem.Text & "';"
    End If
End If
End With
End Sub

Private Sub mnu_Help_Help_Click()
InfoMsg "This version currently do not have an updated help file. Sorry for the inconvenience caused.", "Not Available"
End Sub

Private Sub mnu_Invoicing_Click()
frmInvoicing.Show vbModal
End Sub

Private Sub mnu_Payable_Order_Click()
Dim p As frmPurchase
Set p = New frmPurchase
Load p
p.Show
End Sub

Private Sub mnu_Payable_Suppliers_Click()
frmSuppliers.Show vbModal
End Sub

Private Sub mnu_Payable_View_Click()
frmPurchases_Main.Show
End Sub

Private Sub receivables_Click()
Call mnu_Receivable_Delivery_Click
End Sub

Private Sub txtSearch_GotFocus()
SelText txtSearch
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call Command1_Click
End If
End Sub

Private Sub users_Click()
Call mnu_Admin_Users_Click
End Sub

Private Sub human_Click()
Call mnu_Human_Employees_Click
End Sub

Private Sub inventory_Click()
Call mnu_Inventory_Browse_Click
End Sub

Private Sub po_Click()
Call mnu_Payable_Order_Click
End Sub

Private Sub Image2_Click(Index As Integer)
Select Case Index
    Case 0
        receivables_Click
    Case 1
        po_Click
    Case 2
        inventory_Click
    Case 3
        human_Click
    Case 4
        users_Click
End Select
End Sub

Private Sub mnu_Add_PublicTicker_Click()
frmTicker_New.Show vbModal
End Sub

Private Sub mnu_Admin_Logs_Click()
frmAdmin_Logging.Show vbModal
End Sub

Private Sub mnu_Admin_Settings_Click()
frmAdmin_Settings.Show vbModal
End Sub

Private Sub mnu_Admin_SQL_Click()
frmAdmin_SQL.Show vbModal
End Sub

Private Sub mnu_Admin_Users_Click()
frmUsers_Main.Show
End Sub

Private Sub mnu_Contracts_Click(Index As Integer)
With mnu_Contracts
    If .Item(Index).Caption <> "--{NONE}--" Then
        Dim newF As frmConsignment_Browse
        Set newF = frmConsignment_Browse
        Load newF
        newF.Caption = newF.Caption & " " & .Item(Index).Caption
        newF.getConsignmentDetails .Item(Index).Tag
        newF.Show
    End If
End With
End Sub

Private Sub mnu_Database_Backup_Click()
Dim strFileName As String
InfoMsg "Before attempting to backup the database, ensure no one else is logged on to the system" & _
"because any updates made would not be recorded or captured by the system.", "Warning"

With cDialog

    .DialogTitle = "Backup database"
    .filter = "Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
    .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    .FilterIndex = 1
    .CancelError = False
    .ShowSave
    On Error GoTo ErrHandler
    If .FileName <> "" Then
        Screen.MousePointer = 11
        Dim strTarget As String
        'On Error GoTo ErrHandler
        'mysynondatabase
        MkDir App.Path & Format$(Now(), "dd") & Format$(Now(), "MMMM") & Format$(Now(), "YYYY")
        insertLog "Database backup."
        InfoMsg "The database has been successfully backup.", "Database backup"
    End If
End With
Screen.MousePointer = 0

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub mnu_Database_Restore_Click()
With cDialog
    ' Sets the Dialog Title
    .DialogTitle = "Restore Database"
    
    ' Sets the filter to Access database only
    .filter = "Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
    
    ' Set the default files type to databases
    .FilterIndex = 1
    
    ' Sets the flags - File must exist and Hide Read only
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    
    ' Set dialog box so an error occurs if the dialogbox is cancelled
    .CancelError = False
    
    ' Enables error handling to catch cancel error
    On Error GoTo ErrHandler
    ' display the dialog box
    .ShowOpen
    If .FileName <> "" Then
        Dim sTarget As String 'Open button is clicked
        sTarget = .FileName
        If MsgBox("Are you sure you want to restore the database with this backup copy?" & vbCrLf & sTarget, vbQuestion + vbYesNo, "Restore") = vbYes Then
            'Copy the backup to the default database location.
            Screen.MousePointer = 11
            On Error GoTo ErrHandler
            FileCopy sTarget, DBLocation
            insertLog "Database restored."
            InfoMsg "The database has been successfully restored.", "Database restored"
        End If
    End If
Screen.MousePointer = 0
End With

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub mnu_Help_About_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnu_Human_Browse_Click()
frmLeave_Browse.Show vbModal
End Sub

Private Sub mnu_Human_Employees_Click()
frmEmployees.Show
End Sub

Private Sub mnu_Human_Leave_Click()
frmLeave_Apply.Show vbModal
End Sub

Private Sub mnu_Human_Payrol_Click()
frmPayroll_New.Show vbModal
End Sub

Private Sub mnu_Inventory_Browse_Click()
Load frmProduct_Browse
frmProduct_Browse.WindowState = vbMaximized
frmProduct_Browse.Show
End Sub

Private Sub mnu_Inventory_Categories_Click()
frmCategory_Browse.Show vbModal
End Sub

Private Sub mnu_Inventory_Contracts_Click()
frmConsignment_Main.Show
End Sub

Private Sub mnu_Main_Inventory_Click()
Dim invRS As Recordset
Dim i As Integer
For i = mnu_Contracts.LBound To mnu_Contracts.UBound
    If i <> 0 Then
        Unload mnu_Contracts(i)
    End If
Next i
On Error GoTo ErrHandler
RSOpen invRS, "SELECT Contracts.ContractNo, Customers.Name FROM Customers INNER JOIN Contracts ON Customers.CustomerID=Contracts.CustomerID", dbOpenSnapshot
i = 0
While Not invRS.EOF
    mnu_Contracts(i).Caption = invRS("Name")
    mnu_Contracts(i).Tag = invRS("ContractNo")
    invRS.MoveNext
    If Not invRS.EOF Then
        i = i + 1
        Load mnu_Contracts(i)
    End If
Wend
invRS.Close
Set invRS = Nothing
ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub mnu_Maintenance_Cities_Click()
frmCities.Show vbModal
End Sub

Private Sub mnu_Maintenance_Countries_Click()
frmCountries.Show vbModal
End Sub

Private Sub mnu_Maintenance_States_Click()
frmStates.Show vbModal
End Sub

Private Sub mnu_Manage_Ticker_Click()
frmTickers.Show vbModal
End Sub

Private Sub mnu_Option_Password_Click()
Load frmPassword
frmPassword.setFormMode = mild_Change
frmPassword.Show vbModal
End Sub

Private Sub mnu_Options_Quit_Click()
Unload Me
End Sub

Private Sub mnu_Options_Refresh_Click()
loadRecDeliveries
End Sub

Private Sub mnu_Payable_Add_Click()
Load frmSupplier_New
frmSupplier_New.Show vbModal
End Sub

Private Sub mnu_Personal_Ticker_Click()
frmTicker_Personal.Show vbModal
End Sub

Private Sub mnu_Receivable_Add_Click()
Load frmCustomer_New
frmCustomer_New.Show vbModal
End Sub

Private Sub mnu_Receivable_Customers_Click()
frmCustomers.Show
End Sub

Private Sub mnu_Receivable_Delivery_Click()
Dim f As frmDelivery
Set f = New frmDelivery
Load f
f.Show
End Sub

Private Sub mnu_Receivable_View_Click()
frmDelivery_Main.Show , frmMain
End Sub

Private Sub mnu_Report_Main_Click()
frmReport_Main.Show , frmMain
End Sub

Private Sub tmrTicker_Timer()
Dim currHour As Byte
currHour = Hour(Now())
If currHour > prevHour Then 'Refreshes every hour
    prevHour = currHour
    destroyTicker
    loadTicker
End If
moveTicker 15
End Sub
