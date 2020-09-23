VERSION 5.00
Begin VB.Form frmUsers_Properties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Properties"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   6360
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
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
      TabIndex        =   19
      ToolTipText     =   "Click here to save any changes."
      Top             =   5880
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
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Account Status:"
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   4455
      Begin VB.CheckBox chkStatus 
         Caption         =   "Account is locked out."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   4215
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Account is disabled."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "User must change password on next logon."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Accessibility:"
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4455
      Begin VB.CheckBox chkAccess 
         Caption         =   "Reporting System"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "General Administration System"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Human Resource System"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Delivery Order System"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Accounts Receivable System"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Accounts Payable System"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information:"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmbEmployeeID 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "User name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Employee ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmUsers_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum AccessCodes
'Enumerated values used to refer to the array of checkboxes representing the modules of the system
    APS '0
    ARS
    DOS
    HRS
    Admin
    REP
End Enum

Private Sub cmdApply_Click()
If (Len(cmbEmployeeID.Text) = 0) Then
    ValidMsg "Please enter an Employee ID", "Missing values"
    cmbEmployeeID.SetFocus
ElseIf (Len(txtUsername.Text) = 0) Then
    ValidMsg "Please enter a user name.", "Missing values"
    txtUsername.SetFocus
ElseIf (Len(txtPassword.Text) = 0) Then
    ValidMsg "Please enter a password.", "Missing values"
    txtPassword.SetFocus
Else
    Dim userRS As Recordset
    Dim tempSQL As String
    Dim accessName As AccessCodes
    tempSQL = "SELECT * FROM Users WHERE Users.Username = '" & txtUsername.Text & "';"
    On Error GoTo ErrHandler
    RSOpen userRS, tempSQL, dbOpenDynaset
    If Not userRS.EOF Then
        userRS.Edit 'Assumes username is fixed, begin editing in record
        userRS("Password") = txtPassword.Text
        'set account status
        If chkStatus(0).Value = vbChecked Then
            userRS("MustChange") = True
        Else
            userRS("MustChange") = False
        End If
        If chkStatus(1).Value = vbChecked Then
            userRS("isDisabled") = True
        Else
            userRS("isDisabled") = False
        End If
        If chkStatus(2).Value = vbChecked Then
            userRS("isLocked") = True
        Else
            userRS("isLocked") = False
        End If
        'set accessbility
        accessName = Admin
        If chkAccess(accessName) = vbChecked Then
            userRS("gotAdmin") = True
        Else
            userRS("gotAdmin") = False
        End If
        accessName = APS
        If chkAccess(accessName) = vbChecked Then
            userRS("gotAPS") = True
        Else
            userRS("gotAPS") = False
        End If
        accessName = ARS
        If chkAccess(accessName) = vbChecked Then
            userRS("gotARS") = True
        Else
            userRS("gotARS") = False
        End If
        accessName = DOS
        If chkAccess(accessName) = vbChecked Then
            userRS("gotDOS") = True
        Else
            userRS("gotDOS") = False
        End If
        accessName = HRS
        If chkAccess(accessName) = vbChecked Then
            userRS("gotHRS") = True
        Else
            userRS("gotHRS") = False
        End If
        accessName = REP
        If chkAccess(accessName) = vbChecked Then
            userRS("gotReport") = True
        Else
            userRS("gotReport") = False
        End If
        userRS.Update
        'userRS.Close
        'Set userRS = Nothing
        InfoMsg "User properties have been successfully saved. Changes will only be effective when the user logs on the next time.", "Save"
    End If
    userRS.Close
    Set userRS = Nothing
    Unload Me
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim tempSQL As String
tempSQL = "SELECT Employees.EmployeeID FROM Employees;"
Dim tempRS As Recordset
On Error Resume Next
RSOpen tempRS, tempSQL, dbOpenSnapshot
While Not tempRS.EOF
    cmbEmployeeID.addItem tempRS("EmployeeID")
    tempRS.MoveNext
Wend
tempRS.Close
Set tempRS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUsers_Properties = Nothing
End Sub

Private Sub txtPassword_GotFocus()
SelText txtPassword
End Sub

Private Sub txtUsername_GotFocus()
SelText txtUsername
End Sub

Public Sub getUserProp(ByVal AccUsername As String)
'This procedure obtain the record details of the user with the parameter passed by value
Dim userRS As Recordset
Dim tempSQL As String
tempSQL = "SELECT * FROM Users WHERE Users.Username='" & AccUsername & "';"
On Error GoTo ErrHandler
RSOpen userRS, tempSQL, dbOpenSnapshot
'get user info
cmbEmployeeID.Text = userRS("EmployeeID")
txtUsername.Text = userRS("Username")
txtPassword.Text = userRS("Password")
'get account status
If CBool(userRS("MustChange")) = True Then
    chkStatus(0).Value = vbChecked
Else
    chkStatus(0).Value = vbUnchecked
End If
If CBool(userRS("isDisabled")) = True Then
    chkStatus(1).Value = vbChecked
Else
    chkStatus(1).Value = vbUnchecked
End If
If CBool(userRS("isLocked")) = True Then
    chkStatus(2).Value = vbChecked
Else
    chkStatus(2).Value = vbUnchecked
End If
'Verify accessibility
Dim accessName As AccessCodes
accessName = APS
If CBool(userRS("gotAPS")) = True Then
    chkAccess(accessName).Value = vbChecked
Else
    chkAccess(accessName).Value = vbUnchecked
End If
accessName = ARS
If CBool(userRS("gotARS")) = True Then
    chkAccess(accessName).Value = vbChecked
Else
    chkAccess(accessName).Value = vbUnchecked
End If
accessName = DOS
If CBool(userRS("gotDOS")) = True Then
    chkAccess(accessName).Value = vbChecked
Else
    chkAccess(accessName).Value = vbUnchecked
End If
accessName = HRS
If CBool(userRS("gotHRS")) = True Then
    chkAccess(accessName).Value = vbChecked
Else
    chkAccess(accessName).Value = vbUnchecked
End If
accessName = Admin
If CBool(userRS("gotAdmin")) = True Then
    chkAccess(accessName).Value = vbChecked
Else
    chkAccess(accessName).Value = vbUnchecked
End If
accessName = REP
If CBool(userRS("gotReport")) = True Then
    chkAccess(accessName).Value = vbChecked
Else
    chkAccess(accessName).Value = vbUnchecked
End If

userRS.Close
Set userRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub
