VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployees 
   Caption         =   "Human Resource System"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmployees.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Picture         =   "frmEmployees.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click here to view the properties of the selected employee."
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Picture         =   "frmEmployees.frx":1254
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click here to remove the selected employee."
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Picture         =   "frmEmployees.frx":1B1E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click here to add a new employee."
      Top             =   0
      Width           =   1095
   End
   Begin MSComctlLib.ListView list_Employees 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4683
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmEmployees.frx":23E8
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEmployees.frx":2CB2
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Menu mnu_Pop 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnu_Refresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub formatListView(ByRef strListView As ListView)
With strListView
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Employee ID", 1100
    .ColumnHeaders.add , , "Name", 4000
    .ColumnHeaders.add , , "Gender", 700
    .ColumnHeaders.add , , "Race", 900
End With
End Sub

Public Sub GetAllEmployees()
Dim tempSQL As String
tempSQL = "SELECT Employees.EmployeeID, Employees.Name, Employees.Gender, Employees.Race" & _
        " FROM Employees ORDER BY Employees.EmployeeID ASC;"
Dim employeeRS As Recordset
On Error GoTo ErrHandler
RSOpen employeeRS, tempSQL, dbOpenSnapshot
list_Employees.ListItems.Clear
While Not employeeRS.EOF
    With list_Employees
        .ListItems.add , , employeeRS("EmployeeID")
        .ListItems(.ListItems.Count).SubItems(1) = employeeRS("Name")
        .ListItems(.ListItems.Count).SubItems(2) = IIf(employeeRS("Gender"), "Female", "Male")
        .ListItems(.ListItems.Count).SubItems(3) = IIf(IsNull(employeeRS("Race")), "", employeeRS("Race"))
        employeeRS.MoveNext
    End With
Wend
employeeRS.Close
Set employeeRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub cmdAdd_Click()
frmEmployees_New.Show vbModal
End Sub

Private Sub cmdProperties_Click()
If list_Employees.SelectedItem.Selected Then
    loadEmployeeProperties
End If
End Sub

Private Sub cmdRemove_Click()
Dim tempString As String
With list_Employees
    If .ListItems.Count > 0 Then
        If .SelectedItem.Selected = True Then
            If MsgBox("Are you sure you want to remove this employee from the database?" & vbCrLf & "All his/her related records would be deleted. Eg: Leaves, Payroll, etc.", vbYesNo + vbQuestion, "Remove") = vbYes Then
                tempString = .SelectedItem.Text
                Dim tempSQL As String
                Dim tempRS As Recordset
                tempSQL = "SELECT Users.Status, Users.EmployeeID FROM Users WHERE Users.EmployeeID='" & tempString & "';"
                RSOpen tempRS, tempSQL, dbOpenSnapshot
                If Not tempRS.EOF Then
                    If tempRS("Status") = "ONLINE" Then 'Checks if user is currently online
                        CriticalMsg "Unable to remove selected employee. The employee is currently logged on to the system." & vbCrLf & _
                        "Ensure employee has logged off and retry again.", "Employee online"
                    Else
                        MySynonDatabase.Execute ("DELETE Employees.* FROM Employees WHERE Employees.EmployeeID='" & tempString & "'")
                        If Err.Number <> 0 Then
                            CriticalMsg "Unable to remove employee from database. Please try again. If you see this message twice, please contact system administrator immediately.", "Critical Error"
                        Else
                            InfoMsg "Employee with the ID: " & tempString & " has been successfully removed from the database.", "Remove successful"
                        End If
                    End If
                Else
                    'Means the user account to this system does not exist. Safe to delete.
                    MySynonDatabase.Execute ("DELETE Employees.* FROM Employees WHERE Employees.EmployeeID='" & tempString & "'")
                    If Err.Number <> 0 Then
                        CriticalMsg "Unable to remove employee from database. Please try again. If you see this message twice, please contact system administrator immediately.", "Critical Error"
                    Else
                        InfoMsg "Employee with the ID: " & tempString & " has been successfully removed from the database.", "Remove successful"
                    End If
                End If
                tempRS.Close
                Set tempRS = Nothing
            End If
        End If
    End If
End With

GetAllEmployees
End Sub

Private Sub Form_Load()
formatListView list_Employees
GetAllEmployees
If CurrentUser.prvlgAdmin = True Then
    cmdRemove.Visible = True
Else
    cmdRemove.Visible = False
End If
cmdRemove.Enabled = False
cmdProperties.Enabled = False
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
list_Employees.Move list_Employees.Left, list_Employees.Top, Me.ScaleWidth - 225, Me.ScaleHeight - (list_Employees.Top + 150)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmEmployees = Nothing
End Sub

Private Sub list_Employees_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With list_Employees '// change to the name of the list view
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    .SortKey = iCur
    iLast = iCur
End With

End Sub

Private Sub list_Employees_DblClick()
With list_Employees
If .ListItems.Count > 0 Then
    If .SelectedItem.Selected = True Then
        loadEmployeeProperties
    End If
End If
End With
End Sub

Private Sub list_Employees_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Selected Then
    cmdRemove.Enabled = True
    cmdProperties.Enabled = True
Else
    cmdRemove.Enabled = False
    cmdProperties.Enabled = False
End If
End Sub

Private Sub list_Employees_KeyDown(KeyCode As Integer, Shift As Integer)
With list_Employees
    If .SelectedItem.Selected = True Then
        If KeyCode = vbKeyReturn Then
            loadEmployeeProperties
        End If
    End If
End With
End Sub

Private Sub loadEmployeeProperties()
If CurrentUser.prvlgAdmin = True Then
    Load frmEmployees_Properties
    frmEmployees_Properties.getEmployeeInfo list_Employees.SelectedItem.Text
    frmEmployees_Properties.Show vbModal
Else
    InfoMsg "You do not have the permission to view further details.", "Unauthorised access"
End If
End Sub

Private Sub list_Employees_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnu_Pop, vbPopupMenuLeftAlign
End If
End Sub
