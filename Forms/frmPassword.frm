VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
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
   ScaleHeight     =   2505
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   3
      ToolTipText     =   "Click here to save your new password."
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Re-enter your new password here."
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter your new password here."
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Enter your old password here."
      Top             =   960
      Width           =   2295
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
      TabIndex        =   4
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPassword.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm password:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "New password:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Old password:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mode_Name As form_Condition
Public Property Get getFormMode() As form_Condition
    getFormMode = mode_Name
End Property

Public Property Let setFormMode(ByVal strFormCondition As form_Condition)
    mode_Name = strFormCondition
    If strFormCondition = force_Change Then
        cmdSave.Left = cmdCancel.Left
    Else
        cmdSave.Left = cmdCancel.Left - (cmdSave.width + 105)
    End If
End Property

'Checks if password is match
Private Sub saveChanges()
    If txtOld.Text <> CurrentUser.strPassword Then
        ValidMsg "Invalid password. Please re-enter your correct old password.", "Invalid password"
        txtOld.SetFocus
    ElseIf Len(txtNew.Text) < 6 Then
        ValidMsg "Please enter a password of 6 or more characters in length.", "Missing entry"
        txtNew.SetFocus
    ElseIf txtConfirm.Text <> txtNew.Text Then
        ValidMsg "Please ensure that your new password is the same as the confirmation password.", "Invalid password"
        txtConfirm.SetFocus
    Else
        'Begin saving process
        On Error GoTo ErrHandler
        Screen.MousePointer = 11
        Dim tempSQL As String
        tempSQL = "UPDATE Users SET Password='" & txtConfirm.Text & "', lastChange='" & Format$(Now(), "dd/mm/yyyy") & "', MustChange=False " & _
                    "WHERE Users.Username='" & CurrentUser.strUsername & "';"
        MySynonDatabase.Execute tempSQL
        Screen.MousePointer = 0
        
        'inform user
        InfoMsg "Your password has been successfully changed.", "Record updated"
        Unload Me
    End If

ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Unable to save your password. Please try again." & vbCrLf & _
    "If you see this message twice, please contact the system administrator.", "Error occurred"
    Exit Sub
End If
End Sub

Private Sub checkEntry()
If (Len(txtOld.Text) = 0) Or (Len(txtNew.Text) = 0) Or (Len(txtConfirm.Text) = 0) Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
saveChanges
End Sub

Private Sub Form_Load()
setFormMode = mild_Change
lblNote.Caption = "Re-enter your current password before entering the new password." & vbCrLf & "Please remember that password is case-sensitive." & _
vbCrLf & "Keys not allowed: !, @, #, $, %, ^, &, *, (, ), -, +, =, {, }, [, ], |, \, :, ;, ', "", <, >, ?, /, ~, `"
End Sub

Private Sub txtConfirm_Change()
checkEntry
End Sub

Private Sub txtConfirm_GotFocus()
SelText txtConfirm
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
onlyPassword KeyAscii
End Sub

Private Sub txtNew_Change()
checkEntry
End Sub

Private Sub txtNew_GotFocus()
SelText txtNew
End Sub

Private Sub txtNew_KeyPress(KeyAscii As Integer)
onlyPassword KeyAscii
End Sub

Private Sub txtOld_Change()
checkEntry
End Sub

Private Sub txtOld_GotFocus()
SelText txtOld
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
onlyPassword KeyAscii
End Sub
