VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCountries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Countries"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
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
   Icon            =   "frmCountries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
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
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Click here to close this window."
      Top             =   4920
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
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Click here to delete the selected country."
      Top             =   4920
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
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Click here to edit the selected country."
      Top             =   4920
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
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "Click here to add a new country."
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvCountries 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4260
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
      Left            =   3000
      TabIndex        =   7
      ToolTipText     =   "Click here to save any changes."
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
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Click here to cancel adding or editing."
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmCountries.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Country ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmCountries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tCountry
    id As String
    description As String
End Type

Dim currCountry As tCountry, tempCountry As tCountry
Dim isAdding As Boolean
Private Sub cmdAdd_Click()
tempCountry.id = txtID.Text
tempCountry.description = txtDescription.Text
txtID.Text = ""
txtDescription.Text = ""
FormMode Editing
isAdding = True
End Sub

Private Sub cmdCancel_Click()
txtID.Text = tempCountry.id
txtDescription.Text = tempCountry.description
FormMode Viewing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtID.Text = "" Then
    ValidMsg "Please select an item first.", "No item selected"
Else
    If MsgBox("Are you sure you want to delete this country?", vbYesNo + vbQuestion, "Delete record") = vbYes Then
        Dim tmpSQL As String
        tmpSQL = "DELETE * FROM Countries WHERE CountryID='" & txtID.Text & "';"
        BeginTrans
        On Error GoTo ErrHandler
        MySynonDatabase.Execute tmpSQL
        CommitTrans
        InfoMsg "The country record has been deleted.", "Record deleted"
        getCountries
        txtID.Text = ""
        txtDescription.Text = ""
    End If
End If
ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub cmdEdit_Click()
If txtID.Text <> "" Then
    tempCountry.id = txtID.Text
    tempCountry.description = txtDescription.Text
    FormMode Editing
    isAdding = False
Else
    ValidMsg "Please select a country.", "No item selected"
End If
End Sub

Private Sub cmdSave_Click()
Dim saveSQL As String
If txtID.Text = "" Then
    ValidMsg "Please enter a country ID.", "Missing value"
    txtID.SetFocus
ElseIf txtDescription = "" Then
    ValidMsg "Please enter a description.", "Missing description"
    txtDescription.SetFocus
Else
    If isAdding = True Then
        saveSQL = "INSERT INTO Countries VALUES('" & txtID.Text & "','" & txtDescription.Text & "');"
    Else
        saveSQL = "UPDATE Countries SET CountryID='" & txtID.Text & "',CountryName='" & txtDescription.Text & "' WHERE CountryID='" & tempCountry.id & "';"
    End If
        BeginTrans
        MySynonDatabase.Execute saveSQL
        CommitTrans
        If isAdding = True Then
            InfoMsg "The new country has been added to the database successfully.", "Record saved"
        Else
            InfoMsg "The country detail has been updated successfully.", "Record updated"
        End If
        getCountries
        FormMode Viewing
End If
ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
With lvCountries
    .ColumnHeaders.add , , "Country ID"
    .ColumnHeaders.add , , "Description"
    .ColumnHeaders(1).width = 0.25 * .width
    .ColumnHeaders(2).width = 0.7 * .width
End With
DisableClose Me, True
getCountries
lblNote.Caption = "The list of countries here are for data entry purposes." & vbCrLf & "Add/edit/delete the countries here to make it available for selection." & _
vbCrLf & "The ID will be automatically converted to upper case letters. Each country ID must be unique."
FormMode Viewing
End Sub

Private Sub takeFromDB(ByRef strRecordset As Recordset)
strRecordset("CountryID") = currCountry.id
strRecordset("CountryName") = currCountry.description
End Sub

Private Sub showOnList()
With lvCountries
    .ListItems.add , , currCountry.id
    .ListItems(.ListItems.Count).SubItems(1) = currCountry.description
End With
End Sub

Private Sub getCountries()
Dim tempSQL As String
Dim tempRS As Recordset

lvCountries.ListItems.Clear
tempSQL = "SELECT * FROM Countries ORDER BY Countries.CountryID;"
RSOpen tempRS, tempSQL, dbOpenSnapshot

While Not tempRS.EOF
    'takeFromDB (tempRS)
    currCountry.id = tempRS("CountryID")
    currCountry.description = tempRS("CountryName")
    showOnList
    tempRS.MoveNext
Wend
tempRS.Close
Set tempRS = Nothing
End Sub

Private Sub FormMode(ByVal tmpFormMode As ModeStatus)
Select Case tmpFormMode
    Case Editing
        txtID.Enabled = True
        txtDescription.Enabled = True
        cmdAdd.Visible = False
        cmdEdit.Visible = False
        cmdDelete.Visible = False
        cmdClose.Visible = False
        lvCountries.Enabled = False
    Case Viewing
        txtID.Enabled = False
        txtDescription.Enabled = False
        cmdAdd.Visible = True
        cmdEdit.Visible = True
        cmdDelete.Visible = True
        cmdClose.Visible = True
        lvCountries.Enabled = True
End Select
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCountries = Nothing
End Sub

Private Sub lvCountries_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        txtID.Text = .Text
        txtDescription.Text = .SubItems(1)
    End If
End With
End Sub

Private Sub txtDescription_GotFocus()
SelText txtDescription
End Sub

Private Sub txtID_GotFocus()
SelText txtID
End Sub

Private Sub txtID_LostFocus()
CapCon txtID
End Sub
