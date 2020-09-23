VERSION 5.00
Begin VB.Form frmReport_Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reports"
   ClientHeight    =   4920
   ClientLeft      =   5250
   ClientTop       =   2325
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   5415
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "View Listing of Payroll"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":0BD4
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   480
         Picture         =   "frmReport_Main.frx":0EDE
         Top             =   480
         Width           =   480
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   975
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame rptInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5415
      Begin VB.Label lblInventory 
         BackStyle       =   0  'Transparent
         Caption         =   "View Consignment Transactions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":17A8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   720
         Width           =   2655
      End
      Begin VB.Image imgInventory 
         Height          =   480
         Left            =   480
         Picture         =   "frmReport_Main.frx":1AB2
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblInventory 
         BackStyle       =   0  'Transparent
         Caption         =   "View Inventory Transactions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":237C
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblInventory 
         BackStyle       =   0  'Transparent
         Caption         =   "View Products Below Required Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":2686
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.Shape shpInventory 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1215
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5415
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Picture         =   "frmReport_Main.frx":2990
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "View Sales Drill-Down Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":325A
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmReport_Main.frx":3564
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   975
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmReport_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Report divided into
'Deliveries/Sales/Invoices
'Purchases
'Employees
'Transactions
'Inventory - low level

Private Sub Form_Load()
Me.Move 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmReport_Main = Nothing
End Sub

Private Sub Label3_Click()
frmReport_Sales.Show vbModal
End Sub

Private Sub Label4_Click()
frmReport_Payroll.Show vbModal
End Sub

Private Sub lblInventory_Click(Index As Integer)
Load frmReport_Inventory
With frmReport_Inventory
    Select Case Index
    Case 0
        .getTransactions
        .Show vbModal
    Case 1
        .getLevel
        .Show vbModal
    Case 2
        frmReport_Consignment.Show vbModal
    End Select
End With
End Sub
