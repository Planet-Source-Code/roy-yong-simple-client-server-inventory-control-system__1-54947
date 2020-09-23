VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmployees_Properties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Properties"
   ClientHeight    =   7200
   ClientLeft      =   -135
   ClientTop       =   255
   ClientWidth     =   9000
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
   ScaleHeight     =   7200
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      ScaleHeight     =   255
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.ComboBox mmComm 
      Height          =   315
      ItemData        =   "frmEmployees_Properties.frx":0000
      Left            =   7080
      List            =   "frmEmployees_Properties.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   24
      ToolTipText     =   "Month"
      Top             =   1920
      Width           =   615
   End
   Begin VB.ComboBox ddComm 
      Height          =   315
      ItemData        =   "frmEmployees_Properties.frx":0061
      Left            =   6360
      List            =   "frmEmployees_Properties.frx":00C5
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.ComboBox yyyyComm 
      Height          =   315
      ItemData        =   "frmEmployees_Properties.frx":0147
      Left            =   7800
      List            =   "frmEmployees_Properties.frx":0149
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox mmBirth 
      Height          =   315
      ItemData        =   "frmEmployees_Properties.frx":014B
      Left            =   2160
      List            =   "frmEmployees_Properties.frx":0173
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Month"
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox ddBirth 
      Height          =   315
      ItemData        =   "frmEmployees_Properties.frx":01A7
      Left            =   1440
      List            =   "frmEmployees_Properties.frx":0208
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox yyyyBirth 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin MSComCtl2.UpDown udChildren 
      Height          =   285
      Left            =   1800
      TabIndex        =   61
      Top             =   2640
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtChildren"
      BuddyDispid     =   196647
      OrigLeft        =   1800
      OrigTop         =   2640
      OrigRight       =   2040
      OrigBottom      =   2895
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resignation:"
      Height          =   1455
      Left            =   5160
      TabIndex        =   26
      Top             =   2280
      Width           =   3735
      Begin VB.ComboBox ddResign 
         Height          =   315
         ItemData        =   "frmEmployees_Properties.frx":0288
         Left            =   1320
         List            =   "frmEmployees_Properties.frx":02EC
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox mmResign 
         Height          =   315
         ItemData        =   "frmEmployees_Properties.frx":036E
         Left            =   2040
         List            =   "frmEmployees_Properties.frx":0399
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Month"
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox yyyyResign 
         Height          =   315
         ItemData        =   "frmEmployees_Properties.frx":03CF
         Left            =   2760
         List            =   "frmEmployees_Properties.frx":03D1
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cmbReason 
         Height          =   315
         ItemData        =   "frmEmployees_Properties.frx":03D3
         Left            =   1320
         List            =   "frmEmployees_Properties.frx":03E6
         TabIndex        =   32
         Text            =   "[Please select one]"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton optResignYes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optResignNo 
         Caption         =   "No"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Reason:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date of Resign:"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Resigned:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.OptionButton optMarriedNo 
      Caption         =   "No"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton optMarriedYes 
      Caption         =   "Yes"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
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
      Left            =   7800
      TabIndex        =   34
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   7800
      TabIndex        =   33
      ToolTipText     =   "Click here to save any changes."
      Top             =   4080
      Width           =   1095
   End
   Begin MSComctlLib.ListView list_Progress 
      Height          =   975
      Left            =   120
      TabIndex        =   36
      Top             =   6120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1720
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtNotes 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   5040
      Width           =   8775
   End
   Begin VB.ComboBox cmbPosition 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtSalary 
      Height          =   285
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   22
      ToolTipText     =   "This is the monthly salary."
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtTFN 
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   20
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtSocso 
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   19
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtEPF 
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   18
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cmbRace 
      Height          =   315
      ItemData        =   "frmEmployees_Properties.frx":0437
      Left            =   1440
      List            =   "frmEmployees_Properties.frx":0447
      TabIndex        =   9
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtIC 
      Height          =   285
      Left            =   1440
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   17
      Text            =   "00000"
      Top             =   4440
      Width           =   735
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ComboBox cmbCity 
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   13
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox txtEmployeeID 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtChildren 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "0"
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblHidden 
      Height          =   255
      Left            =   2520
      TabIndex        =   62
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Career Progress:"
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Married:"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Comm:"
      Height          =   255
      Index           =   17
      Left            =   5160
      TabIndex        =   54
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Salary:"
      Height          =   255
      Index           =   16
      Left            =   5160
      TabIndex        =   53
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Position:"
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   52
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "TFN:"
      Height          =   255
      Index           =   14
      Left            =   5160
      TabIndex        =   51
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Socso:"
      Height          =   255
      Index           =   13
      Left            =   5160
      TabIndex        =   50
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "EPF:"
      Height          =   255
      Index           =   12
      Left            =   5160
      TabIndex        =   49
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   48
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Country:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "State:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   46
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   45
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   44
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "No. of Children:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   43
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Race:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   42
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "IC Number:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   41
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Birth:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   40
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Gender:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   39
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Employee ID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEmployees_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim previousPos As String 'Used to check if the employee has been promoted
Private maxSalary As Single
Private Sub cmbCity_Change()
CheckEmpFields
End Sub

Private Sub cmbCity_GotFocus()
If cmbCity.Text = "[PLEASE SELECT ONE]" Then
    cmbCity.Text = ""
End If
SelText cmbCity
End Sub

Private Sub cmbCity_LostFocus()
CapCon cmbCity
End Sub

Private Sub cmbCountry_Change()
CheckEmpFields
End Sub

Private Sub cmbCountry_Click()
FillComboState cmbState, cmbCountry.Text
End Sub

Private Sub cmbCountry_GotFocus()
If cmbCountry.Text = "[PLEASE SELECT ONE]" Then
    cmbCountry.Text = ""
End If
SelText cmbCountry
End Sub

Private Sub cmbCountry_LostFocus()
CapCon cmbCountry
End Sub

Private Sub cmbPosition_Change()
CheckEmpFields
End Sub

Private Sub cmbPosition_GotFocus()
SelText cmbPosition
End Sub

Private Sub cmbPosition_LostFocus()
CapCon cmbPosition
End Sub

Private Sub cmbRace_Change()
CheckEmpFields
End Sub

Private Sub cmbRace_GotFocus()
If cmbRace.Text = "[PLEASE SELECT ONE]" Then
    cmbRace.Text = ""
End If
SelText cmbRace
End Sub

Private Sub cmbReason_Change()
CheckEmpFields
End Sub

Private Sub cmbReason_GotFocus()
If cmbReason.Text = "[PLEASE SELECT ONE]" Then
    cmbReason.Text = ""
End If
End Sub

Private Sub cmbState_Change()
CheckEmpFields
End Sub

Private Sub cmbState_Click()
FillComboCity cmbCity, cmbState.Text
End Sub

Private Sub cmbState_GotFocus()
If cmbState.Text = "[PLEASE SELECT ONE]" Then
    cmbState.Text = ""
End If
SelText cmbState
End Sub

Private Sub cmbState_LostFocus()
CapCon cmbState
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
'Run validation
If isDateValid(CByte(ddBirth.Text), CByte(mmBirth.Text), CInt(yyyyBirth.Text)) = False Then
    Err.Clear
    ValidMsg "Please enter a valid date of birth.", "Invalid date"
    ddBirth.SetFocus
ElseIf isDateValid(CByte(ddComm.Text), CByte(mmComm.Text), CInt(yyyyComm.Text)) = False Then
    Err.Clear
    ValidMsg "Please enter a valid commencement date.", "Invalid date"
    ddComm.SetFocus
ElseIf (CCur(txtSalary.Text) < 0) Or (CCur(txtSalary.Text) > maxSalary) Then
    Err.Clear
    ValidMsg "Please enter a value between 0 and " & maxSalary & " for salary.", "Invalid salary"
    txtSalary.SetFocus
Else
    If optResignYes.Value = True Then
        If isDateValid(CByte(ddResign.Text), CByte(mmResign.Text), CInt(yyyyResign.Text)) = False Then
            Err.Clear
            ValidMsg "Please enter a valid resignation date.", "Invalid date"
            ddResign.SetFocus
            Exit Sub
        End If
    End If
    
    Dim tempSQL As String
    Dim empRS As Recordset
    On Error GoTo ErrHandler
    tempSQL = "SELECT * FROM Employees WHERE Employees.EmployeeID='" & lblhidden.Caption & "';"
    'On Error GoTo ErrHandler
    Set empRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset, dbDenyWrite + dbDenyRead)
    If Not empRS.EOF Then
        empRS.Edit
        empRS("EmployeeID") = txtEmployeeID.Text
        empRS("Name") = txtName.Text
        If optMale.Value = True Then
            empRS("Gender") = False
        Else
            empRS("Gender") = True
        End If
        empRS("DOB") = ddBirth.Text & "/" & mmBirth.Text & "/" & yyyyBirth.Text
        empRS("IC") = txtIC.Text
        empRS("Race") = cmbRace.Text
        If optMarriedYes.Value = True Then
            empRS("Maritial") = True
        Else
            empRS("Maritial") = False
        End If
            empRS("Children") = txtChildren.Text
        empRS("Address") = txtAddress.Text
        empRS("CountryID") = cmbCountry.Text
        empRS("StateID") = cmbState.Text
        empRS("City") = cmbCity.Text
        empRS("Zip") = txtZip.Text
        empRS("SSN") = txtSocso.Text
        empRS("EPF") = txtEPF.Text
        empRS("TFN") = txtTFN.Text
        empRS("PositionID") = cmbPosition.Text
        empRS("Salary") = txtSalary.Text
        empRS("Commence") = ddComm.Text & "/" & mmComm.Text & "/" & yyyyComm.Text
        If optResignYes.Value = True Then
            empRS("Resign") = True
            empRS("Resignation") = ddResign.Text & "/" & mmResign.Text & "/" & yyyyResign.Text
            empRS("Reason") = cmbReason.Text
        Else
            empRS("Resign") = False
        End If
        BeginTrans
        empRS.Update
        CommitTrans
        empRS.Close
        If cmbPosition.Text <> previousPos Then
            tempSQL = "INSERT INTO Progress VALUES ('" & txtEmployeeID.Text & "','" & Format$(Now(), "dd/mm/yyyy") & "','Promoted from " & _
            previousPos & " to the position of a " & cmbPosition.Text & ".');"
            MySynonDatabase.Execute tempSQL
        End If
        InfoMsg "Employee details have been successfully updated.", "Record saved"
        Set empRS = Nothing
        Unload Me
    End If
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub ddBirth_Click()
CheckEmpFields
End Sub

Private Sub ddBirth_GotFocus()
SelText ddBirth
End Sub

Private Sub ddBirth_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub ddComm_Click()
CheckEmpFields
End Sub

Private Sub ddComm_GotFocus()
SelText ddComm
End Sub

Private Sub ddComm_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub ddComm_LostFocus()
If Len(ddComm.Text) > 0 Then
    ddComm.Text = Format(ddComm.Text, "00")
End If
End Sub

Private Sub ddResign_Click()
CheckEmpFields
End Sub

Private Sub ddResign_GotFocus()
SelText ddResign
End Sub

Private Sub ddResign_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub Form_Load()
Move 0, 0
Dim i As Integer
i = 65
While Not i < 17
    yyyyBirth.addItem Format$(Year(Now()) - i)
    i = i - 1
Wend

For i = 0 To 20
    yyyyComm.addItem Format$(Year(Now()) - 10 + i)
    yyyyResign.addItem Format$(Year(Now()) - 10 + i)
Next i
FillComboCountry cmbCountry
Dim tempSQL As String
maxSalary = getSettings("maxSalary")
tempSQL = "SELECT Positions.PositionID FROM Positions ORDER BY Positions.PositionID ASC;"
FillCombo cmbPosition, tempSQL, "PositionID"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmEmployees_Properties = Nothing
End Sub

Private Sub mmBirth_Click()
CheckEmpFields
End Sub

Private Sub mmBirth_GotFocus()
SelText mmBirth
End Sub

Private Sub mmBirth_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub mmComm_Click()
CheckEmpFields
End Sub

Private Sub mmComm_GotFocus()
SelText mmComm
End Sub

Private Sub mmComm_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub


Private Sub mmResign_Click()
CheckEmpFields
End Sub

Private Sub mmResign_GotFocus()
SelText mmResign
End Sub

Private Sub mmResign_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub mmResign_LostFocus()
If Len(mmResign.Text) > 0 Then
    mmResign.Text = Format(mmResign.Text, "00")
End If
End Sub

Private Sub optMarriedNo_Click()
If optMarriedYes.Value = True Then
    udChildren.Enabled = True
Else
    udChildren.Enabled = False
End If
CheckEmpFields
End Sub

Private Sub optMarriedYes_Click()
If optMarriedYes.Value = True Then
    udChildren.Enabled = True
Else
    udChildren.Enabled = False
End If
CheckEmpFields
End Sub

Private Sub optResignNo_Click()
If optResignYes.Value = True Then
    ddResign.Enabled = True
    mmResign.Enabled = True
    yyyyResign.Enabled = True
    cmbReason.Enabled = True
Else
    ddResign.Enabled = False
    mmResign.Enabled = False
    yyyyResign.Enabled = False
    cmbReason.Enabled = False
End If
CheckEmpFields

End Sub

Private Sub optResignYes_Click()
If optResignYes.Value = True Then
    ddResign.Enabled = True
    mmResign.Enabled = True
    yyyyResign.Enabled = True
    cmbReason.Enabled = True
Else
    ddResign.Enabled = False
    mmResign.Enabled = False
    yyyyResign.Enabled = False
    cmbReason.Enabled = False
End If
CheckEmpFields

End Sub

Private Sub txtAddress_Change()
CheckEmpFields
End Sub

Private Sub txtAddress_GotFocus()
SelText txtAddress
End Sub

Private Sub txtAddress_LostFocus()
CapCon txtAddress
End Sub

Private Sub txtChildren_Change()
CheckEmpFields
End Sub

Private Sub txtChildren_GotFocus()
SelText txtChildren
End Sub

Private Sub txtChildren_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtChildren_LostFocus()
If Len(txtChildren.Text) = 0 Then
    txtChildren.Text = "0"
End If
End Sub

Private Sub txtEmployeeID_Change()
CheckEmpFields
End Sub

Private Sub txtEmployeeID_GotFocus()
SelText txtEmployeeID
End Sub

Private Sub txtEPF_Change()
CheckEmpFields
End Sub

Private Sub txtEPF_GotFocus()
SelText txtEPF
End Sub

Private Sub txtIC_Change()
CheckEmpFields
End Sub

Private Sub txtIC_GotFocus()
SelText txtIC
End Sub

Private Sub txtIC_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtName_Change()
CheckEmpFields
End Sub

Private Sub txtName_GotFocus()
SelText txtName
End Sub

Private Sub txtName_LostFocus()
CapCon txtName
End Sub

Private Sub txtNotes_Change()
CheckEmpFields
End Sub

Private Sub txtNotes_GotFocus()
SelText txtNotes
End Sub

Private Sub txtSalary_Change()
CheckEmpFields
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
If Len(txtSalary.Text) > 0 Then
    txtSalary.Text = Format$(txtSalary.Text, "#,##0.00")
End If
End Sub

Private Sub txtSocso_Change()
CheckEmpFields
End Sub

Private Sub txtSocso_GotFocus()
SelText txtSocso
End Sub

Private Sub txtTFN_Change()
CheckEmpFields
End Sub

Private Sub txtTFN_GotFocus()
SelText txtTFN
End Sub

Private Sub txtZip_Change()
CheckEmpFields
End Sub

Private Sub txtZip_GotFocus()
SelText txtZip
End Sub

Private Sub txtZip_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub yyyyBirth_Click()
CheckEmpFields
End Sub

Private Sub yyyyBirth_GotFocus()
SelText yyyyBirth
End Sub

Private Sub yyyyBirth_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub yyyyBirth_LostFocus()
If Len(yyyyBirth.Text) > 0 Then
    yyyyBirth.Text = Format(yyyyBirth.Text, "00")
End If
End Sub

Private Sub yyyyComm_Click()
CheckEmpFields
End Sub

Private Sub yyyyComm_GotFocus()
SelText yyyyComm
End Sub

Private Sub yyyyComm_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub


Private Sub yyyyResign_Click()
CheckEmpFields
End Sub

Private Sub yyyyResign_GotFocus()
SelText yyyyResign
End Sub

Private Sub yyyyResign_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub CheckEmpFields()
If (Len(txtEmployeeID.Text) = 0) Or (Len(txtName.Text) = 0) Or ((optMale.Value = False) And (optFemale.Value = False)) Or _
(Len(ddBirth.Text) = 0) Or (Len(mmBirth.Text) = 0) Or (Len(yyyyBirth.Text) = 0) Or (Len(txtIC.Text) = 0) Or _
(Len(cmbRace.Text) = 0) Or (Len(txtAddress.Text) = 0) Or (Len(cmbCountry.Text) = 0) Or (Len(cmbState.Text) = 0) Or _
(Len(cmbCity.Text) = 0) Or (Len(txtZip.Text) = 0) Or (Len(cmbPosition.Text) = 0) Or (Len(txtSalary.Text) = 0) Or _
(Len(ddComm.Text) = 0) Or (Len(mmComm.Text) = 0) Or (Len(yyyyComm.Text) = 0) Then
    cmdSave.Enabled = False
ElseIf (optMarriedYes.Value = True) And (Len(txtChildren.Text) = 0) Then
    cmdSave.Enabled = False
ElseIf ((optResignYes.Value = True) And (optResignNo.Value = False)) And ((Len(ddResign.Text) = 0) Or (Len(mmResign.Text) = 0) Or (Len(yyyyResign.Text) = 0)) And (Len(cmbReason.Text) = 0) Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If
End Sub

Public Sub getEmployeeInfo(ByVal strEmployeeID As String)
Dim tempSQL As String
Dim empRS As Recordset
Screen.MousePointer = 11
tempSQL = "SELECT * FROM Employees WHERE Employees.EmployeeID='" & strEmployeeID & "';"

'On Error GoTo ErrHandler
RSOpen empRS, tempSQL, dbOpenSnapshot
Dim tempDate() As String

txtEmployeeID.Text = empRS("EmployeeID")
lblhidden.Caption = empRS("EmployeeID")
txtName.Text = empRS("Name")
optFemale.Value = IIf((empRS("Gender") = True), True, False)
optMale.Value = IIf((empRS("Gender") = True), False, True)
tempDate = Split(empRS("DOB"), "/")
ddBirth.Text = Format$(tempDate(0), "00")
mmBirth.Text = Format$(tempDate(1), "00")
yyyyBirth.Text = Format$(tempDate(2), "0000")
txtIC.Text = empRS("IC")
cmbRace.Text = IIf(IsNull(empRS("Race")), "", empRS("Race"))
optMarriedYes.Value = IIf((empRS("Maritial") = True), True, False)
If optMarriedYes.Value = True Then
    udChildren.Enabled = True
Else
    udChildren.Enabled = False
End If
optMarriedNo.Value = IIf((empRS("Maritial") = True), False, True)
txtChildren.Text = empRS("Children")
txtAddress.Text = empRS("Address")
cmbCountry.Text = empRS("CountryID")
cmbState.Text = empRS("StateID")
cmbCity.Text = empRS("City")
txtZip.Text = empRS("Zip")
txtEPF.Text = IIf(IsNull(empRS("EPF")), "", empRS("EPF"))
txtSocso.Text = IIf(IsNull(empRS("SSN")), "", empRS("SSN"))
txtTFN.Text = IIf(IsNull(empRS("TFN")), "", empRS("TFN"))
cmbPosition.Text = empRS("PositionID")
previousPos = empRS("PositionID")
txtSalary.Text = Format$(empRS("Salary"), "#,##0.00")
tempDate = Split(empRS("Commence"), "/")
ddComm.Text = Format$(tempDate(0), "00")
mmComm.Text = Format$(tempDate(1), "00")
yyyyComm.Text = Format$(tempDate(2), "0000")
optResignYes.Value = IIf((empRS("Resign") = True), True, False)
optResignNo.Value = IIf((empRS("Resign") = True), False, True)
If Not IsNull(empRS("Resignation")) Then
    tempDate = Split(empRS("Resignation"), "/") 'date
    ddResign.Text = Format$(tempDate(0), "00")
    mmResign.Text = Format$(tempDate(1), "00")
    yyyyResign.Text = Format$(tempDate(2), "0000")
End If
cmbReason.Text = IIf(IsNull(empRS("Reason")), "", empRS("Reason"))
txtNotes.Text = IIf(IsNull(empRS("Notes")), "", empRS("Notes"))
empRS.Close
Set empRS = Nothing
Call cmbCountry_Click
Call cmbState_Click
getHistory txtEmployeeID.Text
Screen.MousePointer = 0


ErrHandler:
If Err.Number <> 0 Then
    Screen.MousePointer = 0
    ErrorNotifier Err.Number, Err.description
    Unload Me
End If
End Sub

Private Sub getHistory(ByVal strEmployeeID As String)
Dim pastRS As Recordset, pastSQL As String
pastSQL = "SELECT Progress.Date, Progress.Note FROM Progress WHERE Progress.EmployeeID='" & strEmployeeID & "';"
On Error GoTo ErrHandler
RSOpen pastRS, pastSQL, dbOpenSnapshot

With list_Progress
    'Clear any existing header
    .View = lvwReport
    .ColumnHeaders.Clear
    'Specify the columns and properties
    .ColumnHeaders.add , , "Date"
    .ColumnHeaders.add , , "Memo"
    .ColumnHeaders(2).width = .width - .ColumnHeaders(1).width
    'Clear any existing content
    .ListItems.Clear
    While Not pastRS.EOF
        .ListItems.add , , pastRS("Date")
        .ListItems(.ListItems.Count).SubItems(1) = pastRS("Note")
        pastRS.MoveNext
    Wend
End With
pastRS.Close
Set pastRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

