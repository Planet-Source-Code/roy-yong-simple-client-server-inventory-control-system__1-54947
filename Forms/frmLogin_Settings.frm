VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin_Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Settings"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
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
   ScaleHeight     =   1080
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cBrowse 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
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
      Left            =   7320
      TabIndex        =   1
      ToolTipText     =   "Click here to browse for the database."
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      TabIndex        =   2
      ToolTipText     =   "Click here to save the changes."
      Top             =   600
      Width           =   1215
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
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Database Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogin_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
Dim tPathName As String
With cBrowse
    .DialogTitle = "Locate"
    .Flags = cdlOFNHideReadOnly 'cdlOFNReadOnly
    
    .DefaultExt = "|MDB Files|*.mdb|All Files|*.*|"
    .InitDir = App.Path
    .ShowOpen
    
    tPathName = .FileName
    txtLocation.Text = tPathName
End With
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If txtLocation.Text = "" Then
    ValidMsg "Please enter the path name of the database.", "Missing string"
    txtLocation.SetFocus
Else
    On Error GoTo ErrHandler
    AddToINI "Database", "Path", txtLocation.Text, pathFileSettings
    DBLocation = txtLocation.Text
    Unload Me
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
    Exit Sub
End If
End Sub

Private Sub Form_Load()
'Load the path from ini file.
txtLocation.Text = GetFromINI("Database", "Path", "", pathFileSettings)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin_Settings = Nothing
End Sub

Private Sub txtLocation_GotFocus()
SelText txtLocation
End Sub
