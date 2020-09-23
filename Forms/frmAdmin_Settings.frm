VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmin_Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administration Settings"
   ClientHeight    =   6840
   ClientLeft      =   4755
   ClientTop       =   1185
   ClientWidth     =   4560
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdateUser 
      Caption         =   "&Update"
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
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Click here to update the settings."
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "Click here to close this window."
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Environment-Related:"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      Begin VB.TextBox worker 
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "9"
         ToolTipText     =   "The default rate is 9 percent."
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox boss 
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "11"
         ToolTipText     =   "The default rate is 11 percent."
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox compName 
         Height          =   285
         Left            =   1560
         MaxLength       =   120
         TabIndex        =   12
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox numDays 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "26"
         ToolTipText     =   "The default value is 26 days."
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtprod 
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "0"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtpo 
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "0"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtdo 
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "0"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtemp 
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtSalary 
         Height          =   285
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1320
         Width           =   975
      End
      Begin MSComCtl2.UpDown udCart 
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCartMax"
         BuddyDispid     =   196617
         OrigLeft        =   3600
         OrigTop         =   840
         OrigRight       =   3855
         OrigBottom      =   1095
         Max             =   25
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCartMax 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         Top             =   840
         Width           =   360
      End
      Begin VB.CheckBox chkPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow &Price to be shown in Inventory:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "%"
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "%"
         Height          =   255
         Left            =   3840
         TabIndex        =   29
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "days"
         Height          =   255
         Left            =   3840
         TabIndex        =   28
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Company Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Employer EPF contribution rate:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Employee EPF contribution rate:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Human Resource related:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Default number of working days:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Product:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Purchase Order:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Delivery Order:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Employee:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Next unique key for:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Max salary allowed per employee:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Max item allowed in delivery order cart:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   22
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAdmin_Settings.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmAdmin_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub boss_Change()

End Sub

Private Sub boss_GotFocus()
SelText boss
End Sub

Private Sub boss_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub boss_LostFocus()
If boss.Text = "" Then
    'By default - 11%
    boss.Text = "11"
Else
    If Val(boss.Text) > 50 Then
        ValidMsg "The government must be crazy. Please try again.", "Invalid EPF contribution"
        boss.SetFocus
    End If
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdUpdateUser_Click()
Dim tempSQL As String
If compName.Text = "" Then
    ValidMsg "Please enter a company name.", "Missing company name"
    compName.SetFocus
Else
    If chkPrice.Value = vbChecked Then
        tempSQL = "TRUE"
    Else
        tempSQL = "FALSE"
    End If
    On Error GoTo ErrHandler
    BeginTrans
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & tempSQL & "' WHERE Pub_settings.Subject='allowPrice';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & txtCartMax.Text & "' WHERE Pub_settings.Subject='cartSize';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & txtSalary.Text & "' WHERE Pub_settings.Subject='maxSalary';"
    MySynonDatabase.Execute tempSQL
    'Unique keys related
    tempSQL = "UPDATE Misc SET Misc.DataValue='" & txtpo.Text & "' WHERE Misc.DataType='PO';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Misc SET Misc.DataValue='" & txtdo.Text & "' WHERE Misc.DataType='DLVR';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Misc SET Misc.DataValue='" & txtemp.Text & "' WHERE Misc.DataType='EMP';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Misc SET Misc.DataValue='" & txtprod.Text & "' WHERE Misc.DataType='PRODUCT';"
    MySynonDatabase.Execute tempSQL
    'Human resource related
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & numDays.Text & "' WHERE Pub_settings.Subject='numDays';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & worker.Text & "' WHERE Pub_settings.Subject='EPFWorkRate';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & boss.Text & "' WHERE Pub_settings.Subject='EPFEmpRate';"
    MySynonDatabase.Execute tempSQL
    tempSQL = "UPDATE Pub_settings SET Pub_settings.Value='" & compName.Text & "' WHERE Pub_settings.Subject='compName';"
    MySynonDatabase.Execute tempSQL
    
    insertLog "Administration settings changed."
    CommitTrans
    
    InfoMsg "Settings updated. Effect of the changes would begin immediately.", "Settings updated"
End If
ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, "Settings could not be updated. Please try again." & vbCrLf & "If you see this message again, please contact system administrator immediately."
End If
End Sub

Private Sub compName_Change()

End Sub

Private Sub compName_GotFocus()
SelText compName
End Sub

Private Sub compName_KeyPress(KeyAscii As Integer)
tickerKeys KeyAscii
End Sub

Private Sub Form_Load()
lblNotes.Caption = "These administrative settings should be carefully configured as it affects the entire system as a whole." & vbCrLf & _
"Please ensure no other users are logged on before attempting to modify these settings."
If getSettings("allowPrice") = "TRUE" Then
    chkPrice.Value = vbChecked
Else
    chkPrice.Value = vbUnchecked
End If
txtCartMax.Text = getSettings("cartSize")
txtSalary.Text = getSettings("maxSalary")
txtpo.Text = getNextKeys("PO")
txtdo.Text = getNextKeys("DLVR")
txtemp.Text = getNextKeys("EMP")
txtprod.Text = getNextKeys("PRODUCT")
numDays.Text = getSettings("numDays")
worker.Text = getSettings("EPFWorkRate")
boss.Text = getSettings("EPFEmpRate")
compName.Text = getSettings("compName")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAdmin_Settings = Nothing
End Sub

Private Sub numDays_Change()

End Sub

Private Sub numDays_GotFocus()
SelText numDays
End Sub

Private Sub numDays_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub numDays_LostFocus()
If numDays.Text = "" Then
    'By default - 26 days
    numDays.Text = "26"
Else
    If Val(numDays.Text) > 31 Then
        ValidMsg "The number of working days in a month cannot exceed 31 days.", "Invalid days"
        numDays.SetFocus
    End If
End If
End Sub

Private Sub txtCartMax_GotFocus()
SelText txtCartMax
End Sub

Private Sub txtCartMax_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtdo_GotFocus()
SelText txtdo
End Sub

Private Sub txtdo_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtemp_GotFocus()
SelText txtemp
End Sub

Private Sub txtemp_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtPO_GotFocus()
SelText txtpo
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtprod_GotFocus()
SelText txtprod
End Sub

Private Sub txtprod_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtSalary_GotFocus()
SelText txtSalary
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtSalary_LostFocus()
If txtSalary.Text <> "" Then
    txtSalary.Text = Format$(txtSalary.Text, "#,##0.00")
End If
End Sub

Private Sub worker_Change()

End Sub

Private Sub worker_GotFocus()
SelText worker
End Sub

Private Sub worker_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub worker_LostFocus()
If worker.Text = "" Then
    'By default - 9%
    worker.Text = "9"
Else
    If Val(worker.Text) > 50 Then
        ValidMsg "The government must be crazy. Please try again.", "Invalid EPF contribution"
        worker.SetFocus
    End If
End If

End Sub
