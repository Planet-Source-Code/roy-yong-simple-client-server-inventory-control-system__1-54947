VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTickers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticker Management"
   ClientHeight    =   6210
   ClientLeft      =   945
   ClientTop       =   1515
   ClientWidth     =   10095
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
   Icon            =   "frmTickers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7200
      TabIndex        =   13
      ToolTipText     =   "Click here to delete the current selected message."
      Top             =   5760
      Width           =   1335
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
      Left            =   8640
      TabIndex        =   14
      ToolTipText     =   "Click here to close this window."
      Top             =   5760
      Width           =   1335
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
      Left            =   5760
      TabIndex        =   12
      ToolTipText     =   "Click here to edit the current selected message."
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox cmbUser 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4440
      Width           =   2295
   End
   Begin VB.ComboBox created 
      Height          =   315
      Index           =   2
      ItemData        =   "frmTickers.frx":08CA
      Left            =   7920
      List            =   "frmTickers.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.ComboBox created 
      Height          =   315
      Index           =   0
      ItemData        =   "frmTickers.frx":08CE
      Left            =   6480
      List            =   "frmTickers.frx":0932
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox created 
      Height          =   315
      Index           =   1
      ItemData        =   "frmTickers.frx":09B4
      Left            =   7200
      List            =   "frmTickers.frx":09DF
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Month"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox txtMsg 
      Height          =   1215
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4800
      Width           =   4455
   End
   Begin VB.CheckBox chkSpec 
      Caption         =   "Publish on specific date:"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvTicker 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4895
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
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   4920
      Width           =   4935
      Begin VB.ComboBox yyyy 
         Height          =   315
         ItemData        =   "frmTickers.frx":0A15
         Left            =   2760
         List            =   "frmTickers.frx":0A17
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox dd 
         Height          =   315
         ItemData        =   "frmTickers.frx":0A19
         Left            =   1320
         List            =   "frmTickers.frx":0A7D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox mm 
         Height          =   315
         ItemData        =   "frmTickers.frx":0AFF
         Left            =   2040
         List            =   "frmTickers.frx":0B2A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Month"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Publish Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1095
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
      Left            =   8640
      TabIndex        =   16
      Top             =   5760
      Width           =   1335
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
      Left            =   7200
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   23
      Top             =   120
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmTickers.frx":0B60
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblHidden 
      Height          =   255
      Left            =   8880
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Created By:"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Date Created:"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Message Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmTickers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tTicker
    title As String
    message As String
    id As String
    dc As String
    mc As String
    yc As String
    user As String
    isspec As Boolean
    dt As String
    mt As String
    yt As String
End Type
Private currQuery As String
Dim tmpTick As tTicker

Public Sub getTickers(ByVal tSQL As String)
Dim tRS As Recordset
currQuery = tSQL
With lvTicker
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.add , , "Date Created"
    .ColumnHeaders.add , , "Title"
    .ColumnHeaders.add , , "Message"
    .ColumnHeaders.add , , "Publish Date"
    .ColumnHeaders.add , , "User"
    
    RSOpen tRS, tSQL, dbOpenSnapshot
    While Not tRS.EOF
        .ListItems.add , , tRS("dateCreated")
        .ListItems(.ListItems.Count).SubItems(1) = tRS("msgTitle")
        .ListItems(.ListItems.Count).SubItems(2) = tRS("msgText")
        .ListItems(.ListItems.Count).SubItems(3) = IIf(IsNull(tRS("dateToBeShown")), "", tRS("dateToBeShown"))
        .ListItems(.ListItems.Count).SubItems(4) = tRS("username")
        .ListItems(.ListItems.Count).Tag = tRS("tickerID")
        tRS.MoveNext
    Wend
    tRS.Close
    Set tRS = Nothing
End With
ErrHandler:
If Err.Number <> 0 Then
    'Check if table does not exist or syntax errors
    ErrorNotifier Err.Number, Err.description
    Exit Sub
End If
End Sub

Private Sub setFormMode(ByVal tMode As ModeStatus)
Select Case tMode
    Case Editing
        lvTicker.Enabled = False
        txtTitle.Enabled = True
        txtMsg.Enabled = True
        created(0).Enabled = True
        created(1).Enabled = True
        created(2).Enabled = True
        cmbUser.Enabled = True
        chkSpec.Enabled = True
        dd.Enabled = True
        mm.Enabled = True
        yyyy.Enabled = True
        Frame1.Enabled = True
        cmdEdit.Visible = False
        cmdClose.Visible = False
        cmdDelete.Visible = False
    Case Viewing
        lvTicker.Enabled = True
        txtTitle.Enabled = False
        txtMsg.Enabled = False
        created(0).Enabled = False
        created(1).Enabled = False
        created(2).Enabled = False
        cmbUser.Enabled = False
        chkSpec.Enabled = False
        dd.Enabled = False
        mm.Enabled = False
        yyyy.Enabled = False
        Frame1.Enabled = False
        cmdEdit.Visible = True
        cmdClose.Visible = True
        cmdDelete.Visible = True
End Select
End Sub

Private Sub chkSpec_Click()
If chkSpec.Value = vbChecked Then
    Frame1.Enabled = True
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
setFormMode Viewing
created(0).Text = tmpTick.dc
created(1).Text = tmpTick.mc
created(2).Text = tmpTick.yc
lblhidden.Caption = tmpTick.id
txtTitle.Text = tmpTick.title
txtMsg.Text = tmpTick.message
cmbUser.Text = tmpTick.user
If tmpTick.dt = "" Then
    dd.ListIndex = 0
Else
    dd.Text = tmpTick.dt
End If
If tmpTick.mt = "" Then
    mm.ListIndex = 0
Else
    mm.Text = tmpTick.mt
End If
If tmpTick.yt = "" Then
    yyyy.ListIndex = 0
Else
    yyyy.Text = tmpTick.yt
End If
If dd.Text <> "" Then
    chkSpec.Value = vbChecked
Else
    chkSpec.Value = vbUnchecked
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If lblhidden.Caption <> "" Then
    If MsgBox("Are you sure you want to delete this ticker message.", vbYesNo + vbQuestion, "Delete message") = vbYes Then
        Dim tempSQL As String
        tempSQL = "DELETE FROM Ticker WHERE tickerID=" & lblhidden.Caption & ";"
        MySynonDatabase.Execute tempSQL
        InfoMsg "The ticker has been deleted.", "Ticker deleted"
        getTickers currQuery
        If lvTicker.ListItems.Count > 0 Then
            lvTicker.ListItems(lvTicker.ListItems.Count).Selected = True
        Else
            txtTitle.Text = ""
            txtMsg.Text = ""
            cmbUser.Text = ""
            dd.Text = ""
            mm.Text = ""
            yyyy.Text = ""
            created(0).Text = ""
            created(1).Text = ""
            created(2).Text = ""
            chkSpec.Value = Unchecked
            
        End If
    End If
End If
End Sub

Private Sub cmdEdit_Click()
If txtTitle.Text <> "" Then
    setFormMode Editing
    tmpTick.dc = created(0).Text
    tmpTick.mc = created(1).Text
    tmpTick.yc = created(2).Text
    tmpTick.id = lblhidden.Caption
    tmpTick.title = txtTitle.Text
    tmpTick.message = txtMsg.Text
    tmpTick.user = cmbUser.Text
    If chkSpec.Value = vbUnchecked Then
        tmpTick.dt = ""
        tmpTick.mt = ""
        tmpTick.yt = ""
    Else
        tmpTick.dt = dd.Text
        tmpTick.mt = mm.Text
        tmpTick.yt = yyyy.Text
    End If
Else
    InfoMsg "Please select a ticker to be edited.", "Missing selection"
    lvTicker.SetFocus
End If
End Sub

Private Sub Form_Load()
lblNotes.Caption = "Use the ticker manager to add messages to remind you of your personal daily task." & vbCrLf & "Also a useful tool for administrators to communicate with other users publicly."
Dim i As Integer
yyyy.addItem ""
For i = 0 To 5
    yyyy.addItem Format$(Year(Now()) + i, "0000")
    created(2).addItem Format$(Year(Now()) - 2 + i, "0000")
Next i
cmbUser.addItem "", 0
cmbUser.addItem "GENERAL"
FillCombo cmbUser, "SELECT Username FROM Users", "Username"
setFormMode Viewing
If CurrentUser.prvlgAdmin = True Then
    getTickers "SELECT * FROM Ticker WHERE username='" & CurrentUser.strUsername & "' OR username='GENERAL';"
Else
    getTickers "SELECT * FROM Ticker WHERE username='" & CurrentUser.strUsername & "';"
    cmbUser.Locked = True
End If
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmTickers = Nothing
End Sub

Private Sub lvTicker_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        txtTitle.Text = .SubItems(1)
        txtMsg.Text = .SubItems(2)
        created(0).Text = Left$(.Text, 2)
        created(1).Text = Right$(Left$(.Text, 5), 2)
        created(2).Text = Right$(.Text, 4)
        lblhidden.Caption = .Tag
        If .SubItems(3) = "" Then
            chkSpec.Value = vbUnchecked
            dd.ListIndex = 0
            mm.ListIndex = 0
            yyyy.ListIndex = 0
        Else
            chkSpec.Value = vbChecked
            dd.Text = Left$(.SubItems(3), 2)
            mm.Text = Right$(Left$(.SubItems(3), 5), 2)
            yyyy.Text = Right$(.SubItems(3), 4)
        End If
        cmbUser.Text = .SubItems(4)
    End If
End With
End Sub

Private Sub txtMsg_GotFocus()
SelText txtMsg
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

Private Sub txtTitle_GotFocus()
SelText txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

