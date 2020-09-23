VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLeave_Apply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application for Leave"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeave_Apply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker mvLast 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51052544
      CurrentDate     =   38148
   End
   Begin MSComCtl2.DTPicker mvFirst 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51052544
      CurrentDate     =   38148
      MinDate         =   31048
   End
   Begin VB.TextBox txtNotes 
      Height          =   1005
      Left            =   120
      MaxLength       =   125
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmLeave_Apply.frx":058A
      Left            =   1200
      List            =   "frmLeave_Apply.frx":059A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cmbEmployeeID 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Click here to close this window."
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Click here to save the leave now."
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLeave_Apply.frx":05BA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "Additional Notes:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Last Day:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "First Day:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Leave Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Employee ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmLeave_Apply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim firstDay As String, secondDay As String
'Assign to variables
firstDay = Format$(mvFirst.Day, "00") & "/" & Format$(mvFirst.Month, "00") & "/" & mvFirst.Year
secondDay = Format$(mvLast.Day, "00") & "/" & Format$(mvLast.Month, "00") & "/" & mvLast.Year
'MsgBox firstDay & vbCrLf & secondDay & vbCrLf & DateDiff("d", Format$(firstDay, "dd/mm/yyyy"), Format$(secondDay, "dd/mm/yyyy"))
If DateDiff("d", Format$(firstDay, "dd/mm/yyyy"), Format$(secondDay, "dd/mm/yyyy")) < 0 Then
    ValidMsg "Invalid choice of date. Please try again.", "Invalid value"
    mvLast.SetFocus
Else
    'proceed to save changes
    Dim tempSQL As String
    Dim leaveRS As Recordset
    tempSQL = "SELECT * FROM Emp_Leaves"
    Set leaveRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset)
    leaveRS.AddNew
    leaveRS("EmployeeID") = cmbEmployeeID.Text
    leaveRS("date") = Format$(Now(), "dd/mm/yyyy")
    leaveRS("beginDate") = firstDay
    leaveRS("endDate") = secondDay
    leaveRS("type") = cmbType.Text
    leaveRS("notes") = txtNotes.Text
    leaveRS.Update
    leaveRS.Close
    Set leaveRS = Nothing
    InfoMsg "Employee leave record has been created.", "Record saved"
    Unload Me
End If
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Error has occured while adding the record. No changes have been made. Please try again.", "Error"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
lblInstructions.Caption = "Select an employee. Then select the type of leave." & vbCrLf & _
                            "Choose the first day and last day for the duration of the leave." & vbCrLf & _
                            "Enter any notes in addition to the leave. Click on 'Save' when you are done."
FillCombo cmbEmployeeID, "SELECT Employees.EmployeeID FROM Employees;", "EmployeeID"
cmbEmployeeID.ListIndex = 0
cmbType.ListIndex = 0
mvFirst.Value = Format$(Now(), "dd/mm/yyyy")
mvLast.Value = Format$(Now(), "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLeave_Apply = Nothing
End Sub
