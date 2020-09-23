VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "States"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
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
   ScaleHeight     =   6210
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvCountries 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5953
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Click here to add a new state to the current selected country."
      Top             =   5760
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Click here to edit the selected state."
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "Click here to delete the selected state."
      Top             =   5760
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
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Click here to close this window."
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   4
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   4800
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvStates 
      Height          =   3375
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5953
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
      Left            =   3720
      TabIndex        =   10
      Top             =   5760
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Country:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmStates.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "State ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblHidden 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpID As String, tmpDes As String
Dim isAdding As Boolean
Private Sub getListOfCountries()
Dim tempSQL As String
Dim tempRS As Recordset
tempSQL = "SELECT CountryID FROM Countries ORDER BY CountryID ASC"
On Error GoTo ErrHandler
RSOpen tempRS, tempSQL, dbOpenSnapshot
While Not tempRS.EOF
    With lvCountries
        .ListItems.add , , tempRS("CountryID")
        tempRS.MoveNext
    End With
Wend
tempRS.Close
Set tempRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
    Exit Sub
End If
End Sub

Private Sub formatListViews()
With lvCountries
    .View = lvwReport
    .ColumnHeaders.add , , "Country"
    .ColumnHeaders(1).width = .width
End With
With lvStates
    .View = lvwReport
    .ColumnHeaders.add , , "States"
    .ColumnHeaders.add , , "ID"
    .ColumnHeaders(1).width = 0.95 * .width
    .ColumnHeaders(2).width = 0
End With
End Sub

Private Sub getStatesForCountry(ByVal strCountryID As String)
Dim tempSQL As String
Dim tempRS As Recordset
lvStates.ListItems.Clear
tempSQL = "SELECT StateID, StateName FROM States WHERE CountryID='" & strCountryID & "';"
'On Error GoTo ErrHandler
RSOpen tempRS, tempSQL, dbOpenSnapshot
While Not tempRS.EOF
    With lvStates
        .ListItems.add , , tempRS("StateName")
        .ListItems(.ListItems.Count).SubItems(1) = tempRS("StateID")
    End With
    tempRS.MoveNext
Wend

tempRS.Close
Set tempRS = Nothing
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Unable to load states for " & strCountryID & ". Please close this window and try again.", "Unable to find record"
    Exit Sub
End If
End Sub

Private Sub changeMode(ByVal currMode As ModeStatus)
Select Case currMode
    Case Editing
        tmpID = txtID.Text
        tmpDes = txtDescription.Text
        txtID.Enabled = True
        txtDescription.Enabled = True
        cmdAdd.Visible = False
        cmdDelete.Visible = False
        cmdEdit.Visible = False
        cmdClose.Visible = False
        lvCountries.Enabled = False
        lvStates.Enabled = False
    Case Viewing
        txtID.Enabled = False
        txtDescription.Enabled = False
        cmdAdd.Visible = True
        cmdDelete.Visible = True
        cmdEdit.Visible = True
        cmdClose.Visible = True
        lvCountries.Enabled = True
        lvStates.Enabled = True
End Select
End Sub

Private Sub cmdAdd_Click()
isAdding = True
changeMode Editing
txtID.Text = ""
txtDescription.Text = ""
End Sub

Private Sub cmdCancel_Click()
changeMode Viewing
txtID.Text = tmpID
txtDescription.Text = tmpDes
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtID.Text = "" Then
    ValidMsg "Please select a state.", "No item selected"
Else
    If MsgBox("Are you sure you want to delete this state from the country of " & lblHidden.Caption & "?", vbYesNo + vbQuestion, "Delete state") = vbYes Then
        Dim tempSQL As String
        tempSQL = "DELETE * FROM States WHERE CountryID='" & lblHidden.Caption & "' AND StateID='" & txtID.Text & "';"
        BeginTrans
        MySynonDatabase.Execute tempSQL
        CommitTrans
        InfoMsg "The state has been successfully removed from the country of " & lblHidden.Caption & ".", "Record deleted"
        txtID.Text = ""
        txtDescription.Text = ""
        getStatesForCountry lblHidden.Caption
    End If
End If
End Sub

Private Sub cmdEdit_Click()
isAdding = False
changeMode Editing
End Sub

Private Sub cmdSave_Click()
If txtID.Text = "" Then
    ValidMsg "Please enter a state ID.", "Missing value"
    txtID.SetFocus
ElseIf txtDescription.Text = "" Then
    ValidMsg "Please enter a description for the state.", "Missing value"
    txtDescription.SetFocus
Else
    Dim saveSQL As String
    If isAdding = True Then
        saveSQL = "INSERT INTO States VALUES ('" & lblHidden.Caption & "','" & txtID.Text & "','" & txtDescription.Text & "');"
    Else
        saveSQL = "UPDATE States SET CountryID='" & lblHidden.Caption & "', StateID='" & txtID.Text & "', StateName='" & txtDescription.Text & "' WHERE CountryID='" & lblHidden.Caption & "' AND StateID='" & tmpID & "';"
    End If
    On Error GoTo ErrHandler
    BeginTrans
    MySynonDatabase.Execute saveSQL
    CommitTrans
    If isAdding = True Then
        InfoMsg "New state record has been successfully created.", "Record saved"
    Else
        InfoMsg "State information has been successfully updated.", "Record saved"
    End If
    changeMode Viewing
    getStatesForCountry lblHidden.Caption
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
DisableClose Me, True
formatListViews
getListOfCountries
lblNotes.Caption = "The list of states here are for data entry purposes." & vbCrLf & "Add/edit/delete the states here to make it available for selection." & _
vbCrLf & "The ID will be automatically converted to upper case letters. Each state ID must be unique."
changeMode Viewing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmStates = Nothing
End Sub

Private Sub lvCountries_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    txtID.Text = ""
    txtDescription.Text = ""
    If .Selected Then
        lblHidden.Caption = .Text
        Me.Caption = "States - " & lblHidden.Caption
        getStatesForCountry .Text
    Else
        Me.Caption = "States - "
    End If
End With
End Sub

Private Sub lvStates_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    'If .ListItems.Count > 0 Then
        If .Selected Then
            txtID.Text = .SubItems(1)
            txtDescription.Text = .Text
        End If
    'End If
End With
End Sub

Private Sub txtDescription_GotFocus()
SelText txtDescription
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

Private Sub txtID_GotFocus()
SelText txtID
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

Private Sub txtID_LostFocus()
CapCon txtID
End Sub
