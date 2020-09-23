VERSION 5.00
Begin VB.Form frmTicker_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Public Ticker Message"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
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
   ScaleHeight     =   3075
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSpec 
      Caption         =   "Publish on specific date:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
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
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "Click here to close this window."
      Top             =   2640
      Width           =   1215
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
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "Click here to add the new ticker message."
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtMsg 
      Height          =   1215
      Left            =   1560
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
      Begin VB.ComboBox mm 
         Height          =   315
         ItemData        =   "frmTicker_New.frx":0000
         Left            =   2040
         List            =   "frmTicker_New.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Month"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox dd 
         Height          =   315
         ItemData        =   "frmTicker_New.frx":0061
         Left            =   1320
         List            =   "frmTicker_New.frx":00C5
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox yyyy 
         Height          =   315
         ItemData        =   "frmTicker_New.frx":0147
         Left            =   2760
         List            =   "frmTicker_New.frx":0149
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Publish Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Message Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmTicker_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkSpec_Click()
If chkSpec.Value = vbChecked Then
    Frame1.Enabled = True
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub cmdAdd_Click()
If txtTitle.Text = "" Then
    ValidMsg "Please enter a title for the message.", "Missing value"
    txtTitle.SetFocus
ElseIf txtMsg.Text = "" Then
    ValidMsg "Please enter some message text.", "Missing value"
    txtMsg.SetFocus
Else
    If (chkSpec.Value = vbChecked) Then
        If isDateValid(CByte(dd.Text), CByte(mm.Text), CInt(yyyy.Text)) = False Then
            ValidMsg "Please select a valid date for the ticker to be published.", "Invalid date"
            dd.SetFocus
        End If
    Else
        Dim tckRS As Recordset
        Dim tckSQL As String
        tckSQL = "SELECT * FROM Ticker"
        On Error GoTo ErrHandler
        Set tckRS = MySynonDatabase.OpenRecordset(tckSQL, dbOpenDynaset, dbAppendOnly)
            tckRS.AddNew
            tckRS("msgTitle") = Trim$(txtTitle.Text)
            tckRS("msgText") = LTrim$(txtMsg.Text)
            tckRS("dateCreated") = Format$(Now(), "dd/mm/yyyy")
            If chkSpec.Value = vbChecked Then
                tckRS("dateToBeShown") = Format$(dd.Text, "00") & "/" & Format$(mm.Text, "00") & "/" & Format$(yyyy.Text, "0000")
            End If
            tckRS("username") = "GENERAL"
            tckRS.Update
        tckRS.Close
        Set tckRS = Nothing
        InfoMsg "The public ticker message has been created.", "Record added"
        Unload Me
    End If
End If
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Unable to add message. Please check your spellings and spacings and try again.", "Error found"
    Exit Sub
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 5
    yyyy.addItem Format$(Year(Now()) + i, "0000")
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmTicker_New = Nothing
End Sub

Private Sub txtMsg_GotFocus()
SelText txtMsg
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
tickerKeys KeyAscii
End Sub

Private Sub txtTitle_GotFocus()
SelText txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
tickerKeys KeyAscii
End Sub
