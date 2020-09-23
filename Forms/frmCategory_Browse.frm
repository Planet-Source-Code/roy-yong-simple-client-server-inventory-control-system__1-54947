VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategory_Browse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categories"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
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
   ScaleHeight     =   5520
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
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
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Click here to add a new category."
      Top             =   5040
      Width           =   975
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
      Left            =   5760
      TabIndex        =   6
      ToolTipText     =   "Click here to close this window."
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdModify 
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
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Click here to edit the category."
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   4
      Top             =   4560
      Width           =   5535
   End
   Begin VB.TextBox txtCategoryID 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvCategory 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5530
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
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
      Left            =   4680
      TabIndex        =   8
      ToolTipText     =   "Click here to save any changes."
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
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
      TabIndex        =   9
      ToolTipText     =   "Click here to cancel any changes."
      Top             =   5040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmCategory_Browse.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   120
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Category ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "frmCategory_Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isAdding As Boolean
Private Sub adjustFormMode(ByVal strMode As ModeStatus)
If strMode = Editing Then
    lvCategory.Enabled = False
    txtCategoryID.Enabled = True
    txtDescription.Enabled = True
    cmdAdd.Visible = False
    cmdModify.Visible = False
    cmdClose.Visible = False
    cmdSave.Visible = True
    cmdCancel.Visible = True
Else
    lvCategory.Enabled = True
    txtCategoryID.Enabled = False
    txtDescription.Enabled = False
    cmdAdd.Visible = True
    cmdModify.Visible = True
    cmdClose.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
    isAdding = False
End If
End Sub

Private Sub loadCategories()
With lvCategory
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Category ID"
    .ColumnHeaders.add , , "Description", 4850
    .ListItems.Clear
    
    Dim tempSQL As String
    tempSQL = "SELECT * FROM Categories;"
    
    Dim CatRS As Recordset
    On Error GoTo ErrHandler
    RSOpen CatRS, tempSQL, dbOpenSnapshot
    While Not CatRS.EOF
        .ListItems.add , , CatRS("CategoryID")
        .ListItems(.ListItems.Count).SubItems(1) = IIf(IsNull(CatRS("Description")), "", CatRS("Description"))
        CatRS.MoveNext
    Wend
    CatRS.Close
    Set CatRS = Nothing
End With

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub assignValues(ByVal strCategoryID As String, ByVal strDescription As String)
txtCategoryID.Text = strCategoryID
txtDescription.Text = strDescription
End Sub

Private Sub cmdAdd_Click()
isAdding = True
adjustFormMode Editing
txtCategoryID.Text = ""
txtDescription.Text = ""
End Sub

Private Sub cmdCancel_Click()
adjustFormMode Viewing
'Return the modified values to its original
    txtCategoryID.Text = txtCategoryID.Tag
    txtDescription.Text = txtDescription.Tag
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
If txtCategoryID.Text = "" Then
    InfoMsg "Please select an item first.", "Missing selection"
Else
    txtCategoryID.Tag = txtCategoryID.Text
    txtDescription.Tag = txtDescription.Text
    isAdding = False
    adjustFormMode Editing
End If
End Sub

Private Sub cmdSave_Click()
If txtCategoryID.Text = "" Then
    Err.Clear
    ValidMsg "Please enter a category ID for the category.", "Missing values"
    txtCategoryID.SetFocus
ElseIf txtDescription.Text = "" Then
    Err.Clear
    ValidMsg "Please enter a description for the category.", "Missing values"
    txtDescription.SetFocus
Else
    'Done with validation
    'Check if is adding or modifying
    'update by execution of sql statement
    Dim tempSQL As String
    On Error GoTo ErrHandler
    If isAdding = True Then
        tempSQL = "INSERT INTO Categories VALUES ('" & txtCategoryID.Text & "','" & txtDescription.Text & "');"
        MySynonDatabase.Execute tempSQL
        insertLog "Category ID: " & txtCategoryID.Text & " has been created."
        InfoMsg "Category ID: " & txtCategoryID.Text & vbCrLf & "Description: " & txtDescription.Text & vbCrLf & "Record has been successfully created.", "New Category"
    Else
        tempSQL = "UPDATE Categories SET CategoryID='" & txtCategoryID.Text & "', Description='" & txtDescription.Text & "' WHERE CategoryID='" & txtCategoryID.Tag & "';"
        MySynonDatabase.Execute tempSQL
        insertLog "Category ID: " & txtCategoryID.Text & " has been updated."
        InfoMsg "Category ID: " & txtCategoryID.Text & vbCrLf & "Description: " & txtDescription.Text & vbCrLf & "Record has been successfully updated.", "Record modified"
    End If
    adjustFormMode Viewing
    loadCategories
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
lblNotes.Caption = "To add a category, click on 'Add' and begin entering the details of the new category. " & _
"To edit a category, select the category you wish to edit from the list and click on 'Edit'." & vbCrLf & _
"Please note that category ID would be converted to upper-case automatically."

loadCategories
adjustFormMode Viewing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCategory_Browse = Nothing
End Sub

Private Sub lvCategory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lvCategory '// change to the name of the list view
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    .SortKey = iCur
    iLast = iCur
End With

End Sub

Private Sub lvCategory_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        assignValues .Text, .SubItems(1)
    Else
        txtCategoryID.Text = ""
        txtDescription.Text = ""
    End If
End With
End Sub

Private Sub txtCategoryID_GotFocus()
SelText txtCategoryID
End Sub

Private Sub txtCategoryID_LostFocus()
CapCon txtCategoryID
End Sub

Private Sub txtDescription_GotFocus()
SelText txtDescription
End Sub

Private Sub txtDescription_LostFocus()
txtDescription.Text = StrConv(txtDescription.Text, vbProperCase)
End Sub
