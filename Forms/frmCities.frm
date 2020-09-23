VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cities"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
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
   Icon            =   "frmCities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   4560
      Width           =   5535
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
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Click here to close this window."
      Top             =   5160
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
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Click here to delete the selected city."
      Top             =   5160
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
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Click here to edit the selected city."
      Top             =   5160
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
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Click here to add a new city for the selected state."
      Top             =   5160
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvCities 
      Height          =   3375
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5953
      View            =   3
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
   Begin MSComctlLib.TreeView tvState 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5953
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
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
      Left            =   5880
      TabIndex        =   7
      Top             =   5160
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
      Left            =   4680
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmCities.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "City Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpID As String ', tmpDes As String
Dim isAdding As Boolean
Private Sub getListOfCountries()
Dim tempSQL As String
Dim tempRS As Recordset
With tvState.Nodes
    .Clear
    .add , , "C_root", "Countries"
    
    tempSQL = "SELECT * FROM Countries ORDER BY CountryID ASC"
    'On Error GoTo ErrHandler
    RSOpen tempRS, tempSQL, dbOpenSnapshot
    While Not tempRS.EOF
        .add "C_root", tvwChild, tempRS("CountryID"), tempRS("CountryName")
        tempRS.MoveNext
    Wend
    
    tempSQL = "SELECT * FROM States;"
    RSOpen tempRS, tempSQL, dbOpenSnapshot
    While Not tempRS.EOF
        .add CStr(tempRS("CountryID")), tvwChild, tempRS("StateID"), tempRS("StateName")
        .Item(.Count).Tag = "State"
        tempRS.MoveNext
    Wend
    tempRS.Close
    Set tempRS = Nothing
End With

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
    Exit Sub
End If
End Sub

Private Sub formatListViews()
With lvCities
    .View = lvwReport
    .ColumnHeaders.add , , "City"
    .ColumnHeaders(1).width = 0.95 * .width
End With
End Sub

Private Sub getCitiesForState(ByVal strStateID As String)
Dim tempSQL As String
Dim tempRS As Recordset
lvCities.ListItems.Clear
tempSQL = "SELECT City FROM Cities WHERE StateID='" & strStateID & "';"
On Error GoTo ErrHandler
RSOpen tempRS, tempSQL, dbOpenSnapshot
While Not tempRS.EOF
    With lvCities
        .ListItems.add , , tempRS("City")
    End With
    tempRS.MoveNext
Wend
tempRS.Close
Set tempRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Unable to load cities for " & strStateID & ". Please close this window and try again.", "Unable to find record"
    Exit Sub
End If
End Sub

Private Sub changeMode(ByVal currMode As ModeStatus)
Select Case currMode
    Case Editing
        tmpID = txtID.Text
        'tmpDes = txtDes.Text
        txtID.Enabled = True
        'txtDes.Enabled = True
        cmdAdd.Visible = False
        cmdDelete.Visible = False
        cmdEdit.Visible = False
        cmdClose.Visible = False
        tvState.Enabled = False
        lvCities.Enabled = False
    Case Viewing
        txtID.Enabled = False
        'txtDes.Enabled = False
        cmdAdd.Visible = True
        cmdDelete.Visible = True
        cmdEdit.Visible = True
        cmdClose.Visible = True
        tvState.Enabled = True
        lvCities.Enabled = True
End Select
End Sub

Private Sub cmdAdd_Click()
If lblhidden.Caption <> "" Then
    changeMode Editing
    txtID.Text = ""
    isAdding = True
Else
    InfoMsg "Please select a state where the city will be added.", "Missing selection"
End If
End Sub

Private Sub cmdCancel_Click()
changeMode Viewing
txtID.Text = tmpID

'txtDes.Text = tmpDes
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim strCity As String, strState As String
strCity = txtID.Text
strState = tvState.SelectedItem.Text
If txtID.Text = "" Then
    ValidMsg "Please select a city.", "No item selected"
Else
    If MsgBox("Are you sure you want to delete this city from the state of " & strState & "?", vbYesNo + vbQuestion, "Delete city") = vbYes Then
        Dim tempSQL As String
        tempSQL = "DELETE * FROM Cities WHERE City='" & strCity & "' AND StateID='" & strState & "';"
        BeginTrans
        MySynonDatabase.Execute tempSQL
        CommitTrans
        InfoMsg "The city has been successfully removed from the state of " & strCity & ".", "Record deleted"
        txtID.Text = ""
        'txtDes.Text
        getCitiesForState strState
    End If
End If
End Sub

Private Sub cmdEdit_Click()
If lblhidden.Caption <> "" Then
    changeMode Editing
    isAdding = False
Else
    InfoMsg "Please select a city to edit.", "Missing selection"
End If
End Sub

Private Sub cmdSave_Click()
If txtID.Text = "" Then
    Err.Clear
    ValidMsg "Please enter a city name.", "Missing value"
    txtID.SetFocus
Else
    Dim saveSQL As String
    If isAdding = True Then
        saveSQL = "INSERT INTO Cities VALUES ('" & txtID.Text & "','" & lblhidden.Caption & "');"
    Else
        saveSQL = "UPDATE Cities SET City='" & txtID.Text & "' WHERE City='" & tmpID & "';"
    End If
    On Error GoTo ErrHandler
    BeginTrans
    
    MySynonDatabase.Execute saveSQL
    CommitTrans
    If isAdding = True Then
        InfoMsg "New city record has been successfully added.", "Record saved"
    Else
        InfoMsg "City information has been successfully updated.", "Record saved"
    End If
    getCitiesForState lblhidden.Caption
    changeMode Viewing
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
lblNotes.Caption = "The list of cities here are for data entry purposes." & vbCrLf & "Add/edit/delete the cities here to make it available for selection." & _
vbCrLf & "Cities have no ID but will be converted to upper-case. Each city must be unique."
changeMode Viewing
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCities = Nothing
End Sub

Private Sub lvCities_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        txtID.Text = .Text
    End If
End With
End Sub

Private Sub tvState_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Tag = "State" Then
    getCitiesForState Node.Key
    lblhidden.Caption = Node.Key
    txtID.Text = ""
Else
    lvCities.ListItems.Clear
    txtID.Text = ""
    lblhidden.Caption = ""
End If
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

