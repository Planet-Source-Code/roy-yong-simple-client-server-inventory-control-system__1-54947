VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   2595
   ClientLeft      =   4380
   ClientTop       =   4245
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
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
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Click to login when username and password has been entered correctly."
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
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
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Click here to exit the program."
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Menu mnu_Menu_Option 
      Caption         =   "&Option"
      Begin VB.Menu mnu_Option_Database 
         Caption         =   "&Database Settings"
      End
      Begin VB.Menu mnu_Dash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Option_Exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xTrials As Byte
Private Sub cmdExit_Click()
'Asks the user if he/she wishes to quit
If MsgBox("Are you sure you want to exit now?", vbQuestion + vbYesNo, "Exit") = vbYes Then
    Unload Me
End If
End Sub

Private Sub cmdLogin_Click()
Dim strPass As String, strUser As String, tempSQL As String
Dim LoginRS As Recordset
'Assign values to variables
strUser = txtUsername.Text
strPass = txtPassword.Text
'Then load recordset
tempSQL = "SELECT * FROM Users WHERE Users.Username = '" & strUser & "';"
Screen.MousePointer = 11
On Error GoTo ErrHandler
ConnDB
If isOpen = True Then
    RSOpen LoginRS, tempSQL, dbOpenDynaset
    WriStatus "Searching for username......"
    If Not LoginRS.EOF Then
        'Username exist. Check password
        WriStatus "Username found......Verifying password......"
        If LoginRS("Password") <> strPass Then
            'Wrong password
            xTrials = xTrials + 1
            Screen.MousePointer = 0
            ValidMsg "Incorrect password. Please try again." & vbCrLf & _
            "Please note that password is case-sensitive.", "Invalid entry"
            WriStatus ""
            txtPassword.SetFocus
            If xTrials > 3 Then
                CriticalMsg "Unauthorised access detected. The account with the username '" & strUser & "' will be suspended." & _
                vbCrLf & "The program will shut down now.", "Unauthorised access."
                LoginRS.Edit
                LoginRS("isLocked") = True
                LoginRS.Update
                LoginRS.Close
                Set LoginRS = Nothing
                Unload Me
            End If
        ElseIf LoginRS("isDisabled") = True Then
            LoginRS.Close
            Set LoginRS = Nothing
            WriStatus ""
            Screen.MousePointer = 0
            InfoMsg "Your account has been disabled. Please contact system administrator.", "Account disabled"
            Unload Me
        Else
            'Check if account locked.
            WriStatus "Password matched......Checking account status......"
            If LoginRS("isLocked") = True Then
                LoginRS.Close
                Set LoginRS = Nothing
                WriStatus ""
                Screen.MousePointer = 0
                InfoMsg "Your account has been locked. Please contact system administrator. The program will shut down now.", "Account suspended"
                Exit Sub
            Else
                'Account is OK
                WriStatus "Account status excellent......Assigning values......"
                With CurrentUser
                    .isLocked = LoginRS("isLocked")
                    .isDisabled = LoginRS("isDisabled")
                    .lastPassword = LoginRS("lastChange")
                    .mustChange = LoginRS("MustChange")
                    .prvlgAdmin = LoginRS("gotAdmin")
                    .prvlgAPS = LoginRS("gotAPS")
                    .prvlgARS = LoginRS("gotARS")
                    .prvlgDOS = LoginRS("gotDOS")
                    .prvlgHRS = LoginRS("gotHRS")
                    .prvlgReport = LoginRS("gotReport")
                    .strPassword = LoginRS("Password")
                    .strUsername = LoginRS("Username")
                End With
                WriStatus "Assignment complete......Beginning session logging......"
                'Update user status
                tempSQL = "UPDATE Users SET Users.status = 'ONLINE' " & _
                        "WHERE ((Users.Username='" & strUser & "'));"
                MySynonDatabase.Execute tempSQL
                'Add to system log
                tempSQL = "INSERT INTO Logging VALUES ('" & CurrentUser.strUsername & "','User logged on' ,'" & FormatDateTime(Now(), vbLongTime) & "','" & Format$(Now(), "dd/mm/yyyy") & "');"
                MySynonDatabase.Execute tempSQL
                WriStatus "Logging complete......Loading data......"
                LoginRS.Close
                Set LoginRS = Nothing
                Screen.MousePointer = 0
                frmMain.Show
                Unload Me
            End If
        End If
    Else 'No such user name
        Screen.MousePointer = 0
        ValidMsg "Invalid username. Please try again.", "Invalid entry"
        WriStatus ""
        txtUsername.SetFocus
    End If
End If
ErrHandler:
If Err.Number <> 0 Then
    'ErrorNotifier Err.Number, Err.Description
    'CriticalMsg "The database cannot be located or missing. Please try to locate the database using the database settings under Options.", "Database missing"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
xTrials = 0
lblNote.Caption = "Only authorised users are allowed to login." & vbCrLf & "If you forgot your password, please contact the system administrator immediately."
DisableClose Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing
End Sub

Private Sub mnu_Option_Database_Click()
frmLogin_Settings.Show vbModal
End Sub

Private Sub mnu_Option_Exit_Click()
Call cmdExit_Click
End Sub

Private Sub txtPassword_Change()
CheckFields
End Sub

Private Sub txtPassword_GotFocus()
SelText txtPassword
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdLogin_Click
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
onlyPassword KeyAscii
End Sub

Private Sub txtUsername_Change()
CheckFields
End Sub

Private Sub txtUsername_GotFocus()
SelText txtUsername
End Sub

Private Sub CheckFields()
If (Len(txtUsername.Text) = 0) Or (Len(txtPassword.Text) = 0) Then
    cmdLogin.Enabled = False
Else
    cmdLogin.Enabled = True
End If
End Sub

Private Sub WriStatus(ByVal strStatus As String)
lblStatus.Caption = strStatus
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdLogin_Click
End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
onlyPassword KeyAscii
End Sub
