VERSION 5.00
Begin VB.Form frmAdmin_SQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Console"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdmin_SQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Click here to execute the command line."
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Click here to close this window."
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmAdmin_SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExecute_Click()
If txtSQL.Text = "" Then
    ValidMsg "Please enter some SQL statement.", "Empty field"
    txtSQL.SetFocus
Else
    Dim tempSQL As String
    tempSQL = txtSQL.Text
    On Error GoTo ErrHandler
    MySynonDatabase.Execute tempSQL
    InfoMsg "SQL statement successful.", "Execution completed"
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
DisableClose Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAdmin_SQL = Nothing
End Sub
