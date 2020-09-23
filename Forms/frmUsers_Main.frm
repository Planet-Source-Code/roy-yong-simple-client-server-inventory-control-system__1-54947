VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers_Main 
   Caption         =   "Users Management Console"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView list_Users 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "To reset the password of an account, right click on a user and select Reset Password."
      Top             =   1800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
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
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
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
      Left            =   1080
      Picture         =   "frmUsers_Main.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click here to remove the selected user."
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
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
      Picture         =   "frmUsers_Main.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click here to add a new user."
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdProperties 
      Appearance      =   0  'Flat
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
      Left            =   2160
      Picture         =   "frmUsers_Main.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click here to view the properties of the selected user."
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmUsers_Main.frx":23E8
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUsers_Main.frx":2CB2
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnu_Reset 
         Caption         =   "&Reset Password..."
      End
   End
End
Attribute VB_Name = "frmUsers_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
frmUsers_Add.Show vbModal
End Sub

Private Sub cmdProperties_Click()
If list_Users.SelectedItem.Selected Then
    Load frmUsers_Properties
    frmUsers_Properties.getUserProp list_Users.SelectedItem.Text
    frmUsers_Properties.Show vbModal
End If
End Sub

Private Sub cmdRemove_Click()
With list_Users
    If .SelectedItem.Index > -1 Then
        If MsgBox("Are you sure you want to remove this user from the system?", vbQuestion + vbYesNo, "Remove user") = vbYes Then
            'Remove from the list
            Dim strTarget As String
            strTarget = .ListItems(.SelectedItem.Index).Text
            .ListItems.Remove (.SelectedItem.Index)
            strTarget = "DELETE Users.* FROM Users WHERE Users.username='" & strTarget & "';"
            MySynonDatabase.Execute strTarget
            If Err Then
                ErrorNotifier Err.Number, Err.description
            Else
                InfoMsg "User record has been successfully removed from the system.", "Remove successful"
                getUsers
            End If
        End If
    End If
End With
End Sub

Private Sub Form_Load()
FrmtListView list_Users
getUsers
cmdProperties.Enabled = False
cmdRemove.Enabled = False
frmMain.img32.MaskColor = RGB(255, 0, 255)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Shape1.width = Me.width
list_Users.Move list_Users.Left, list_Users.Top, Me.ScaleWidth - 225, Me.ScaleHeight - (list_Users.Top + 150)
End Sub

Private Sub FrmtListView(ByRef lstViewName As ListView)
With lstViewName
    .View = lvwReport
    .ColumnHeaders.add , , "Username", 1350
    .ColumnHeaders.add , , "EmployeeID", 1500
    .ColumnHeaders.add , , "Disabled", 900
    .ColumnHeaders.add , , "Locked Out", 900
    .ColumnHeaders.add , , "Status", 900
End With
End Sub

Public Sub getUsers()
Dim tempSQL As String
'Clear the list of users
list_Users.ListItems.Clear
'define query
tempSQL = "SELECT Users.Username, Users.EmployeeID, Users.isDisabled, Users.isLocked, Users.status " & _
          "FROM Users;"
Dim userRS As Recordset
On Error GoTo ErrHandler
'execute query
RSOpen userRS, tempSQL, dbOpenSnapshot
While Not userRS.EOF
    With list_Users
        .ListItems.add , , userRS("Username")
        .ListItems(.ListItems.Count).SubItems(1) = CStr(userRS("EmployeeID"))
        .ListItems(.ListItems.Count).SubItems(2) = CStr(CBool(userRS("isDisabled")))
        .ListItems(.ListItems.Count).SubItems(3) = CStr(CBool(userRS("isLocked")))
        .ListItems(.ListItems.Count).SubItems(4) = CStr(userRS("status"))
    End With
    userRS.MoveNext
Wend

userRS.Close
Set userRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub list_Users_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With list_Users '// change to the name of the list view
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    .SortKey = iCur
    iLast = iCur
End With
End Sub

Private Sub list_Users_DblClick()
With list_Users
    If .ListItems.Count > 0 Then
        If .SelectedItem.Selected = True Then
            Load frmUsers_Properties
            frmUsers_Properties.getUserProp .SelectedItem.Text
            frmUsers_Properties.Show vbModal
        End If
    End If
End With
End Sub

Private Sub list_Users_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        cmdProperties.Enabled = True
        cmdRemove.Enabled = True
    Else
        cmdProperties.Enabled = False
        cmdRemove.Enabled = False
    End If
End With
End Sub

Private Sub list_Users_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With list_Users
    If .ListItems.Count > 0 Then
        If .SelectedItem.Selected Then
            If Button = vbRightButton Then
                PopupMenu mnu_Options, vbPopupMenuLeftAlign
            End If
        End If
    End If
End With
End Sub

Private Sub mnu_Reset_Click()
With list_Users
    Dim strTemp As String
    strTemp = "Are you sure you want to reset the following selected user's account?" & vbCrLf & "User selected: " & .SelectedItem.Text & vbCrLf & "The username and password will be the same."
    If MsgBox(strTemp, vbQuestion + vbYesNoCancel, "Reset password") = vbYes Then
        strTemp = "UPDATE Users SET Users.Password=Users.Username, Users.MustChange=True WHERE Users.Username='" & .SelectedItem.Text & "';"
        MySynonDatabase.Execute strTemp
        InfoMsg "The user account has been successfully reset and would be required to change his/her password during the next logon.", "Account updated"
    End If
End With
End Sub
