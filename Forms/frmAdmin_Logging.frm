VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmin_Logging 
   Caption         =   "View Logs"
   ClientHeight    =   4110
   ClientLeft      =   240
   ClientTop       =   450
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdmin_Logging.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7770
   WindowState     =   2  'Maximized
   Begin VB.PictureBox splitPane 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2280
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3135
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   0
      Width           =   130
   End
   Begin MSComctlLib.TreeView tvLog 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5106
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvLogging 
      Height          =   2895
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5106
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
End
Attribute VB_Name = "frmAdmin_Logging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SPLIT_WIDTH As Long = 130 'width of the splitter bar
Private Const MIN_WINDOW As Long = 10 'Minimum size For any frame created by splitter bars
'This form provides a navigation tree to administrator
'to view log files on user access as well as other critical logs
Private Sub fillTree()
'Add the nodes to the tree
With tvLog.Nodes
    tvLog.ImageList = frmMain.img16
    .add , , "Root", "System", "server"
    .add "Root", tvwChild, "Branch01", "Systems Log", "arrows"
    .add "Root", tvwChild, "Branch02", "Errors Log", "exclaimation"
    .add "Root", tvwChild, "Branch03", "Internal Transactions", "arrows"
End With
End Sub

Private Sub getDetails(ByVal nodeName As String)
Dim tempSQL As String
tempSQL = ""
Select Case nodeName
    Case "Systems Log"
        tempSQL = "SELECT * FROM Logging;"
    Case "Errors Log"
        tempSQL = "SELECT * FROM Error_Logging;"
    Case "Internal Transactions"
        tempSQL = "SELECT * FROM Internal_Transaction;"
End Select
If Not tempSQL = "" Then
    Dim viewRS As Recordset
    On Error GoTo ErrHandler
    RSOpen viewRS, tempSQL, dbOpenSnapshot
    'format the listview
    With lvLogging
        .View = lvwReport
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        'Format headers according to number of fields
        Dim i As Integer
        For i = 0 To viewRS.Fields.Count - 1
            If viewRS.Fields(i).Name = "isIn" Then
                .ColumnHeaders.add , , "In"
            Else
                .ColumnHeaders.add , , viewRS.Fields(i).Name
            End If
        Next i
        
        'Loop through records and add to list view
        While Not viewRS.EOF
            For i = 0 To viewRS.Fields.Count - 1
                If i = 0 Then
                    .ListItems.add , , IIf(IsNull(viewRS.Fields(i).Value), "", viewRS.Fields(i).Value)
                Else
                    If viewRS.Fields(i).Name = "isIn" Then
                        .ListItems(.ListItems.Count).SubItems(i) = IIf((viewRS.Fields(i).Value = True), "In", "Out")
                    Else
                        .ListItems(.ListItems.Count).SubItems(i) = IIf(IsNull(viewRS.Fields(i).Value), "", viewRS.Fields(i).Value)
                    End If
                End If
            Next i
            viewRS.MoveNext
        Wend
        viewRS.Close
        Set viewRS = Nothing
    End With
End If
ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
Move 0, 0
fillTree
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim height1 As Long, width1 As Long
Dim X As Long, Y As Long, width As Long, height As Long
tvLog.height = Me.ScaleHeight - 185
lvLogging.height = tvLog.height
splitPane.height = Me.height
lvLogging.width = Me.width - (tvLog.width + 450)
End Sub

Private Sub lvLogging_Resize()
Dim x1 As Long
Dim x2 As Long
Dim y1 As Long
    
    On Error Resume Next
    With lvLogging
        .Left = splitPane.Left + splitPane.width + 10
        .width = Me.ScaleWidth - (.Left + 185)
    
        tvLog_Resize
    End With
End Sub

Private Sub tvLog_Resize()
With tvLog
    .width = splitPane.Left - (.Left + 10)
End With
End Sub

Private Sub splitPane_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With splitPane
    If Button = vbLeftButton Then
        splitPane.Appearance = 1
        splitPane.BackColor = vbButtonShadow 'change the splitter colour
        splitPane.Drag vbDropEffectMove
        splitPane.Move (splitPane.Left - (SPLIT_WIDTH \ 2)) + X, 0
    End If
End With
End Sub

Private Sub splitPane_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If splitPane.BackColor = vbButtonShadow Then
        splitPane.Move (splitPane.Left - (SPLIT_WIDTH \ 2)) + X, 0
End If
End Sub

Private Sub splitPane_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If splitPane.BackColor = vbButtonShadow Then
        splitPane.Drag vbDropEffectNone
        splitPane.BackColor = &H8000000F  'restore splitter colour
        splitPane.Move (splitPane.Left - (SPLIT_WIDTH \ 2)) + X, tvLog.Top
        splitPane.Appearance = 0

        'Set the absolute Boundaries
        Dim lAbsLeft As Long
        Dim lAbsRight As Long
        lAbsLeft = tvLog.Left + 500
        lAbsRight = lvLogging.Left + lvLogging.width - (SPLIT_WIDTH + 500)


        Select Case splitPane.Left
            Case Is < lAbsLeft 'the pane is too thin
            splitPane.Move lAbsLeft, 0
            Case Is > lAbsRight 'the pane is too wide
            splitPane.Move lAbsRight, 0
        End Select

    lvLogging_Resize
End If
End Sub

Private Sub tvLog_Click()
With tvLog
If .SelectedItem.Selected = True Then
    getDetails (.SelectedItem.Text)
End If
End With
End Sub
