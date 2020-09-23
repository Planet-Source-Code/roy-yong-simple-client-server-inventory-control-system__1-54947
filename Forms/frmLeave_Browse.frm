VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLeave_Browse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leaves"
   ClientHeight    =   5865
   ClientLeft      =   -135
   ClientTop       =   435
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeave_Browse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8760
   Begin VB.ComboBox cmbFilter_plus 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      ItemData        =   "frmLeave_Browse.frx":058A
      Left            =   1080
      List            =   "frmLeave_Browse.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtApproved 
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ComboBox cmbEmployeeID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmLeave_Browse.frx":05B1
      Left            =   1320
      List            =   "frmLeave_Browse.frx":05C1
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtFirst 
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtLast 
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtNotes 
      Height          =   1095
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4080
      Width           =   3855
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
      Left            =   7440
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvLeaves 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
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
      Left            =   6120
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
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
      Left            =   7440
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Filter by:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLeave_Browse.frx":05E1
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "(dd/mm/yyyy)"
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Date Approved:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   19
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "(dd/mm/yyyy)"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(dd/mm/yyyy)"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label lblHidden 
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Last Day:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "First Day:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Employee ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "frmLeave_Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private type use to remember values of the current instance
Private Type tempLeave
    id As String
    employeeID As String
    type As String
    date As String
    beginDate As String
    endDate As String
    notes As String
End Type
Dim tmpVar As tempLeave, currVar As tempLeave
Private Sub FormMode(ByVal strFormCondition As ModeStatus)
    Select Case strFormCondition
    Case Editing
        cmbEmployeeID.Enabled = True
        cmbType.Enabled = True
        txtFirst.Enabled = True
        txtLast.Enabled = True
        txtNotes.Enabled = True
        cmbEmployeeID.Enabled = True
        cmdEdit.Visible = False
        cmdClose.Visible = False
        lvLeaves.Enabled = False
    Case Viewing
        cmbEmployeeID.Enabled = False
        cmbType.Enabled = False
        txtFirst.Enabled = False
        txtLast.Enabled = False
        txtNotes.Enabled = False
        cmbEmployeeID.Enabled = False
        cmdEdit.Visible = True
        cmdClose.Visible = True
        lvLeaves.Enabled = True
    End Select
End Sub

Private Sub cmbFilter_Click()
Dim i As Integer
cmbFilter_plus.Clear
If cmbFilter.Text = "Employee ID" Then
    FillCombo cmbFilter_plus, "SELECT Employees.EmployeeID FROM Employees;", "EmployeeID"
    cmbFilter_plus.ListIndex = 0
Else
    For i = 0 To cmbType.ListCount
        cmbFilter_plus.addItem cmbType.List(i)
    Next i
End If
cmbFilter_plus.ListIndex = 0
End Sub

Private Sub cmbFilter_plus_Click()
Dim msgFilter As String
msgFilter = "SELECT * FROM Emp_Leaves"
If cmbFilter.Text = "Employee ID" Then
    msgFilter = msgFilter & " WHERE EmployeeID='" & cmbFilter_plus.List(cmbFilter_plus.ListIndex) & "';"
Else
    msgFilter = msgFilter & " WHERE Type='" & cmbFilter_plus.List(cmbFilter_plus.ListIndex) & "';"
End If
getLeavesDB msgFilter
End Sub

Private Sub cmdCancel_Click()
FormMode Viewing
assignValues tmpVar
If isSame(currVar, tmpVar) Then
    showLeave currVar
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If lblhidden.Caption <> "" Then
    FormMode Editing
    assignValues currVar
Else
    ValidMsg "Please select a record to be edited from the list.", "No record selected"
    lvLeaves.SetFocus
End If
End Sub

Private Sub cmdSave_Click()
If cmbEmployeeID.Text = "" Then
    ValidMsg "Please select an employee.", "Missing selection"
    cmbEmployeeID.SetFocus
ElseIf cmbType.Text = "" Then
    ValidMsg "Please select a type of leave.", "Missing selection"
    cmbType.SetFocus
ElseIf txtApproved.Text = "" Then
    ValidMsg "Please enter the approved date.", "Missing date"
    txtApproved.SetFocus
ElseIf txtFirst.Text = "" Then
    ValidMsg "Please enter the first date of leave.", "Missing date"
    txtFirst.SetFocus
ElseIf txtLast.Text = "" Then
    ValidMsg "Please enter the last date of leave.", "Missing date"
    txtLast.SetFocus
ElseIf isDateValid(Mid$(txtFirst.Text, 0, 2), Mid$(txtFirst.Text, 4, 2), Mid$(txtFirst.Text, 7, 4)) = False Then
    ValidMsg "Please enter a valid first date for this leave.", "Invalid date"
    txtFirst.SetFocus
ElseIf isDateValid(Mid$(txtLast.Text, 0, 2), Mid$(txtLast.Text, 4, 2), Mid$(txtLast.Text, 7, 4)) = False Then
    ValidMsg "Please enter a valid last date for this leave.", "Invalid date"
    txtLast.SetFocus
ElseIf isDateValid(Mid$(txtApproved.Text, 0, 2), Mid$(txtApproved.Text, 4, 2), Mid$(txtApproved.Text, 7, 4)) = False Then
    ValidMsg "Please enter a valid approved date for this leave.", "Invalid date"
    txtApproved.SetFocus
Else
    'Done with validation
    assignValues tmpVar
    Dim tempSQL As String
    With tmpVar
        tempSQL = "UPDATE Emp_Leaves " & _
        "SET EmployeeID = '" & .employeeID & "', " & _
        "date = '" & .date & "', type='" & .type & "', beginDate='" & .beginDate & "', " & _
        "endDate = '" & .endDate & "', notes = '" & .notes & "' " & _
        "WHERE leaveID = " & tmpVar.id & ";"
        On Error GoTo ErrHandler
        BeginTrans
        MySynonDatabase.Execute tempSQL
        CommitTrans
        InfoMsg "The leave record has been successfully updated.", "Record saved"
        FormMode Viewing
    End With
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
FillCombo cmbEmployeeID, "SELECT Employees.EmployeeID FROM Employees;", "EmployeeID"
Me.Move 0, 0
formatList
getLeavesDB "SELECT * FROM Emp_Leaves;"
FormMode Viewing
DisableClose Me, True
lblNotes.Caption = "Filter records by selecting the type of filter and followed by the sub-filter options." & vbCrLf & _
"Select a record from the list to edit its detail. Data would not be updated if the Save button is not clicked."
End Sub

Private Sub assignValues(ByRef strVariable As tempLeave)
'Assign control values to variable
With strVariable
    .id = lblhidden.Caption
    .employeeID = cmbEmployeeID.Text
    .date = txtApproved.Text
    .beginDate = txtFirst.Text
    .endDate = txtLast.Text
    .notes = txtNotes.Text
    .type = cmbType.Text
End With
End Sub
Private Sub showLeave(ByRef strVariable As tempLeave)
'display variable values in controls
With strVariable
    lblhidden.Caption = .id
    cmbEmployeeID.Text = .employeeID
    cmbType.Text = .type
    txtApproved.Text = .date
    txtFirst.Text = .beginDate
    txtLast.Text = .endDate
    txtNotes.Text = .notes
End With
End Sub

Private Sub getLeavesDB(ByVal strFilter As String)
'Clear current items and get the leave records from the DB
lvLeaves.ListItems.Clear
Dim leaveRS As Recordset
RSOpen leaveRS, strFilter, dbOpenSnapshot
While Not leaveRS.EOF
    'Assigns the values from the record to the variable
    tmpVar.id = leaveRS("leaveID")
    tmpVar.beginDate = leaveRS("leaveID")
    tmpVar.employeeID = leaveRS("EmployeeID")
    tmpVar.type = leaveRS("type")
    tmpVar.date = leaveRS("date")
    tmpVar.beginDate = leaveRS("beginDate")
    tmpVar.endDate = leaveRS("endDate")
    tmpVar.notes = leaveRS("notes")
    addToList
    leaveRS.MoveNext
Wend
leaveRS.Close
Set leaveRS = Nothing
End Sub

Private Sub addToList()
'Adds the custom type variable into the list
With lvLeaves
    .ListItems.add , , tmpVar.id
    .ListItems(.ListItems.Count).SubItems(1) = tmpVar.employeeID
    .ListItems(.ListItems.Count).SubItems(2) = tmpVar.type
    .ListItems(.ListItems.Count).SubItems(3) = tmpVar.date
    .ListItems(.ListItems.Count).SubItems(4) = tmpVar.beginDate
    .ListItems(.ListItems.Count).SubItems(5) = tmpVar.endDate
    .ListItems(.ListItems.Count).SubItems(6) = tmpVar.notes
End With
End Sub

Private Sub formatList()
With lvLeaves
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "ID"
    .ColumnHeaders(1).width = 0
    .ColumnHeaders.add , , "Employee ID"
    .ColumnHeaders.add , , "Type"
    .ColumnHeaders(3).width = 850
    .ColumnHeaders.add , , "Approved Date"
    .ColumnHeaders.add , , "Beginning of Leave"
    .ColumnHeaders(5).width = 1200
    .ColumnHeaders.add , , "End of Leave"
    .ColumnHeaders.add , , "Notes"
End With
End Sub

Private Function isSame(firstVar As tempLeave, secondvar As tempLeave) As Boolean
With firstVar
    If (.beginDate = secondvar.beginDate) And (.date = secondvar.date) And (.employeeID = secondvar.employeeID) And (.endDate = secondvar.endDate) And _
    (.id = secondvar.id) And (.notes = secondvar.notes) And (.type = secondvar.type) Then
        isSame = True
    Else
        isSame = False
    End If
End With
End Function

Private Sub resetForm()
cmbEmployeeID.ListIndex = 0
cmbType.ListIndex = 0
txtApproved.Text = ""
txtFirst.Text = ""
txtLast.Text = ""
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLeave_Browse = Nothing
End Sub

Private Sub lvLeaves_Click()
With lvLeaves
    If .ListItems.Count > 0 Then 'When there is more than 0 items in list
        If .SelectedItem.Selected = True Then
            tmpVar.id = .SelectedItem.Text
            tmpVar.employeeID = .SelectedItem.SubItems(1)
            tmpVar.type = .SelectedItem.SubItems(2)
            tmpVar.date = .SelectedItem.SubItems(3)
            tmpVar.beginDate = .SelectedItem.SubItems(4)
            tmpVar.endDate = .SelectedItem.SubItems(5)
            tmpVar.notes = .SelectedItem.SubItems(6)
            showLeave tmpVar
        End If
    End If
End With
End Sub

Private Sub lvLeaves_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lvLeaves '// change to the name of the list view
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    .SortKey = iCur
    iLast = iCur
End With

End Sub

Private Sub lvLeaves_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then
    If CurrentUser.prvlgAdmin = False Then
        InfoMsg "You do not have the administrator privilege to perform the deletion task.", "Access denied"
    Else
        If lvLeaves.SelectedItem.Selected Then
            If MsgBox("Are you sure you want to remove this record permanently?" & vbCrLf & _
            "The record cannot be retrieved.", vbQuestion + vbYesNo, "Delete record") = vbYes Then
                Dim deleteSQL As String
                deleteSQL = "DELETE * FROM Emp_Leaves WHERE Emp_Leaves.leaveID='" & lvLeaves.SelectedItem.Text & "';"
                On Error GoTo ErrHandler
                BeginTrans
                MySynonDatabase.Execute deleteSQL
                CommitTrans
                InfoMsg "The selected record has been successfully deleted.", "Record deleted"
                getLeavesDB "SELECT * FROM Emp_Leaves;"
                resetForm
            End If
        End If
    End If
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub txtApproved_GotFocus()
SelText txtApproved
End Sub

Private Sub txtApproved_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc("/") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtFirst_GotFocus()
SelText txtFirst
End Sub

Private Sub txtFirst_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc("/") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtLast_GotFocus()
SelText txtLast
End Sub

Private Sub txtLast_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc("/") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtNotes_GotFocus()
SelText txtNotes
End Sub
