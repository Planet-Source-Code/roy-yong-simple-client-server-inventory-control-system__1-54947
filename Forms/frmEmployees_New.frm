VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmployees_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Employee"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmEmployees_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox yyyyComm 
      Height          =   315
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2880
      Width           =   855
   End
   Begin VB.ComboBox ddComm 
      Height          =   315
      ItemData        =   "frmEmployees_New.frx":000C
      Left            =   6360
      List            =   "frmEmployees_New.frx":006D
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox mmComm 
      Height          =   315
      ItemData        =   "frmEmployees_New.frx":00ED
      Left            =   7080
      List            =   "frmEmployees_New.frx":0115
      Style           =   2  'Dropdown List
      TabIndex        =   23
      ToolTipText     =   "Month"
      Top             =   2880
      Width           =   615
   End
   Begin VB.ComboBox yyyyBirth 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox ddBirth 
      Height          =   315
      ItemData        =   "frmEmployees_New.frx":0149
      Left            =   1440
      List            =   "frmEmployees_New.frx":01AA
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.ComboBox mmBirth 
      Height          =   315
      ItemData        =   "frmEmployees_New.frx":022A
      Left            =   2160
      List            =   "frmEmployees_New.frx":0252
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Month"
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notes"
      Height          =   1455
      Left            =   5160
      TabIndex        =   25
      Top             =   3240
      Width           =   3735
      Begin VB.TextBox txtNotes 
         Height          =   1095
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         ToolTipText     =   "Additional notes in regards to the new employee"
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox txtChildren 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "0"
      ToolTipText     =   "Number of children"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   12
      ToolTipText     =   "Current address"
      Top             =   3600
      Width           =   3615
   End
   Begin VB.ComboBox cmbCity 
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      Text            =   "[PLEASE SELECT ONE]"
      ToolTipText     =   "City of address"
      Top             =   4680
      Width           =   3615
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Text            =   "[PLEASE SELECT ONE]"
      ToolTipText     =   "State of address"
      Top             =   4320
      Width           =   2655
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Text            =   "[PLEASE SELECT ONE]"
      ToolTipText     =   "Country of address"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   16
      Text            =   "00000"
      ToolTipText     =   "Zip/postal code of address"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   0
      ToolTipText     =   "Enter the name of the new employee"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtIC 
      Height          =   285
      Left            =   1440
      MaxLength       =   12
      TabIndex        =   6
      ToolTipText     =   "Identity Card number. (000000000000)"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ComboBox cmbRace 
      Height          =   315
      ItemData        =   "frmEmployees_New.frx":0286
      Left            =   1440
      List            =   "frmEmployees_New.frx":0296
      TabIndex        =   7
      Text            =   "[PLEASE SELECT ONE]"
      ToolTipText     =   "Select the racial status of the new employee"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtEPF 
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   17
      ToolTipText     =   "EPF account number"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtSocso 
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   18
      ToolTipText     =   "Social Security Number"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtTFN 
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   19
      ToolTipText     =   "Income Tax File Number"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtSalary 
      Height          =   285
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   21
      Text            =   "0.00"
      ToolTipText     =   "Current salary."
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox cmbPosition 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2160
      Width           =   2535
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
      Left            =   6600
      TabIndex        =   27
      ToolTipText     =   "Click here to save the new employee details."
      Top             =   4920
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
      TabIndex        =   28
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   4920
      Width           =   1095
   End
   Begin VB.OptionButton optMarriedYes 
      Caption         =   "Yes"
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Click here if employee is married"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.OptionButton optMarriedNo 
      Caption         =   "No"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Click here if employee is not married"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      Top             =   1440
      Width           =   3615
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Click here if employee is a female"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Click here if employee is a male"
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComCtl2.UpDown udChildren 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Click to adjust the number of children"
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtChildren"
      BuddyDispid     =   196617
      OrigLeft        =   1800
      OrigTop         =   3240
      OrigRight       =   2040
      OrigBottom      =   3525
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmEmployees_New.frx":02BA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "* - Red labels indicate required fields"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   48
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   47
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Gender:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Birth:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   45
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "IC Number:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   44
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Race:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   43
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "No. of Children:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   42
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   41
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   40
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "State:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   39
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Country:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   38
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   37
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "EPF:"
      Height          =   255
      Index           =   12
      Left            =   5160
      TabIndex        =   36
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Socso:"
      Height          =   255
      Index           =   13
      Left            =   5160
      TabIndex        =   35
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "TFN:"
      Height          =   255
      Index           =   14
      Left            =   5160
      TabIndex        =   34
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Position:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   33
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Salary:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   16
      Left            =   5160
      TabIndex        =   32
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Comm:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   5160
      TabIndex        =   31
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Married:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmEmployees_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private maxIncome As Single
Private Sub cmdSave_Click()
If isDateValid(CByte(ddBirth.Text), CByte(mmBirth.Text), CInt(yyyyBirth.Text)) = False Then
    ValidMsg "Please enter a valid date of birth.", "Invalid date"
    ddBirth.SetFocus
ElseIf isDateValid(CByte(ddComm.Text), CByte(mmComm.Text), CInt(yyyyComm.Text)) = False Then
    ValidMsg "Please enter a valid commencement date.", "Invalid date"
    ddComm.SetFocus
ElseIf ((CCur(txtSalary.Text) <= 0) Or (CCur(txtSalary.Text) > maxIncome)) Then
    ValidMsg "Please enter a salary between 0 and " & maxIncome & ".", "Invalid salary"
    txtSalary.SetFocus
Else
    Dim tempSQL As String
    'Obtain new ID
    tempSQL = "SELECT * FROM Misc WHERE Misc.DataType='EMP';"
    
    Screen.MousePointer = 11
    Dim newRS As Recordset, empRS As Recordset
    Dim newEmpID As Integer
    Set newRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset)
    newEmpID = CInt(newRS("DataValue"))
    'Insert new record
    tempSQL = "SELECT * FROM Employees;"
    Set empRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset, dbAppendOnly)
    
   'On Error GoTo ErrHandler
    empRS.AddNew
    empRS("EmployeeID") = "EMP" & Format$(newEmpID, "0000")
    empRS("Name") = txtName.Text
    If optMale.Value = True Then
        empRS("Gender") = False
    Else
        empRS("Gender") = True
    End If
    empRS("DOB") = ddBirth.Text & "/" & mmBirth.Text & "/" & yyyyBirth.Text
    empRS("IC") = txtIC.Text
    If optMarriedYes.Value = True Then
        empRS("Maritial") = True
        empRS("Children") = txtChildren.Text
    Else
        empRS("Maritial") = False
    End If
    empRS("Race") = cmbRace.Text
    empRS("Address") = txtAddress.Text
    empRS("CountryID") = cmbCountry.Text
    empRS("StateID") = cmbState.Text
    empRS("City") = cmbCity.Text
    empRS("Zip") = txtZip.Text
    empRS("EPF") = txtEPF.Text
    empRS("SSN") = txtSocso.Text
    empRS("TFN") = txtTFN.Text
    empRS("PositionID") = cmbPosition.Text
    empRS("Salary") = txtSalary.Text
    empRS("Commence") = ddComm.Text & "/" & mmComm.Text & "/" & yyyyComm.Text
    empRS("Notes") = IIf(IsNull(txtNotes.Text), "", txtNotes.Text)
    empRS.Update 'Save record
    
    empRS.Close
    'Increment the existing key ID
    newRS.Edit
    newRS("DataValue") = newEmpID + 1
    newRS.Update
    
    newRS.Close
    Set newRS = Nothing
    Set empRS = Nothing
    'Update the progress
    tempSQL = "INSERT INTO Progress VALUES ('EMP" & Format$(newEmpID, "0000") & "','" & Format$(Now(), "dd/mm/yyyy") & "','Joined the company.');"
    MySynonDatabase.Execute tempSQL
    'Inform user
    Screen.MousePointer = 0
    InfoMsg "Employee ID: EMP" & Format$(newEmpID, "0000") & vbCrLf & "New employee record has been successfully created.", "Record saved"
    frmEmployees.GetAllEmployees
    Unload Me
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
    Exit Sub
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

Private Sub ddComm_GotFocus()
SelText ddComm
End Sub

Private Sub ddComm_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub mmBirth_Click()
CheckEmpFields
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

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
tickerKeys KeyAscii
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

Private Sub yyyyComm_Click()
CheckEmpFields
End Sub

Private Sub yyyyComm_GotFocus()
SelText yyyyComm
End Sub

Private Sub yyyyComm_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
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
Private Sub Form_Load()
lblNotes.Caption = "* - Red labels indicate required fields." & vbCrLf & _
"Please enter all the required details carefully."
FillComboCountry cmbCountry
Dim tempSQL As String
Dim tempRS As Recordset
Dim i As Integer
i = 65
While Not i < 17
    yyyyBirth.addItem Format$(Year(Now()) - i)
    i = i - 1
Wend

For i = 0 To 20
    yyyyComm.addItem Format$(Year(Now()) - 10 + i)
'    yyyyResign.addItem Format$(Year(Now()) - 10 + i)
Next i
tempSQL = "SELECT Positions.PositionID FROM Positions ORDER BY Positions.PositionID ASC;"
FillCombo cmbPosition, tempSQL, "PositionID"
maxIncome = getSettings("maxSalary")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmEmployees_New = Nothing
End Sub

Private Sub mmBirth_Change()
CheckEmpFields
If Len(mmBirth.Text) = 2 Then
    yyyyBirth.SetFocus
End If
End Sub

Private Sub mmBirth_GotFocus()
SelText mmBirth
End Sub

Private Sub mmBirth_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub mmBirth_LostFocus()
If Len(mmBirth.Text) > 0 Then
    mmBirth.Text = Format(mmBirth.Text, "00")
End If
End Sub
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

Private Sub CheckEmpFields()
If (Len(txtName.Text) = 0) Or ((optMale.Value = False) And (optFemale.Value = False)) Or _
(Len(ddBirth.Text) = 0) Or (Len(mmBirth.Text) = 0) Or (Len(yyyyBirth.Text) = 0) Or (Len(txtIC.Text) = 0) Or _
(Len(cmbRace.Text) = 0) Or (Len(txtAddress.Text) = 0) Or (Len(cmbCountry.Text) = 0) Or (Len(cmbState.Text) = 0) Or _
(Len(cmbCity.Text) = 0) Or (Len(txtZip.Text) = 0) Or (Len(cmbPosition.Text) = 0) Or (Len(txtSalary.Text) = 0) Or _
(Len(ddComm.Text) = 0) Or (Len(mmComm.Text) = 0) Or (Len(yyyyComm.Text) = 0) Then
    cmdSave.Enabled = False
ElseIf (optMarriedYes.Value = True) And (Len(txtChildren.Text) = 0) Then
    cmdSave.Enabled = False
'ElseIf ((optResignYes.Value = True) And (optResignNo.Value = False)) And ((Len(ddResign.Text) = 0) Or (Len(mmResign.Text) = 0) Or (Len(yyyyResign.Text) = 0)) And (Len(cmbReason.Text) = 0) Then
'    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If
End Sub


