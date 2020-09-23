VERSION 5.00
Begin VB.Form frmUsers_Add 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers_Add.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "User Information:"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cmbEmployeeID 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Employee ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "User name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Accessibility:"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4455
      Begin VB.CheckBox chkAccess 
         Caption         =   "Reporting System"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Accounts Payable System"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Accounts Receivable System"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Delivery Order System"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "Human Resource System"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "General Administration System"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Account Status:"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
      Begin VB.CheckBox chkStatus 
         Caption         =   "User must change password on next logon."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Account is disabled."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Account is locked out."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   4215
      End
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
      TabIndex        =   15
      ToolTipText     =   "Click here to close this window without saving any changes."
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
      Left            =   2280
      TabIndex        =   14
      ToolTipText     =   "Click here to save any changes."
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "frmUsers_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
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
    
    userRS.AddNew  'Assumes username is fixed, begin editing in record
    userRS("username") = txtUsername.Text
    userRS("Password") = txtPassword.Text
    userRS("EmployeeID") = cmbEmployeeID.Text
    userRS("Status") = "OFFLINE"
    userRS("lastChange") = Format$(Now(), "mm/dd/yyyy")
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
    userRS.Close
    Set userRS = Nothing
    InfoMsg "Username: " & txtUsername.Text & vbCrLf & "Employee ID: " & cmbEmployeeID.Text & vbCrLf & "User account has been successfully created.", "Record saved"
    tempSQL = "INSERT INTO Internal_Transaction VALUES('" & Format$(Now(), "dd/mm/yyyy") & "','User " & txtUsername.Text & " account has been created','" & CurrentUser.strUsername & "');"
    MySynonDatabase.Execute tempSQL
    frmUsers_Main.getUsers
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
DisableClose Me, True
tempSQL = "SELECT Employees.EmployeeID FROM Employees;"
FillCombo cmbEmployeeID, tempSQL, "EmployeeID"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUsers_Add = Nothing
End Sub

Private Sub txtPassword_GotFocus()
SelText txtPassword
End Sub

Private Sub txtUsername_GotFocus()
SelText txtUsername
End Sub

