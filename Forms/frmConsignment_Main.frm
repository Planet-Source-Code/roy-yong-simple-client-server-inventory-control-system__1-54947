VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsignment_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consignment Contracts"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
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
   ScaleHeight     =   5280
   ScaleWidth      =   9105
   StartUpPosition =   1  'CenterOwner
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
      Left            =   7800
      TabIndex        =   7
      ToolTipText     =   "Click here to delete the selected consignment contract."
      Top             =   4320
      Width           =   1095
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
      Left            =   7800
      TabIndex        =   6
      ToolTipText     =   "Click here to edit the selected contract."
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
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
      Left            =   7800
      TabIndex        =   5
      ToolTipText     =   "Click here to create a new consignment contract."
      Top             =   3360
      Width           =   1095
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
      Left            =   7800
      TabIndex        =   8
      ToolTipText     =   "Click here to close this window."
      Top             =   4800
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvContracts 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4048
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
   Begin VB.TextBox txtContract 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox cmbCustomer 
      Height          =   315
      ItemData        =   "frmConsignment_Main.frx":0000
      Left            =   1920
      List            =   "frmConsignment_Main.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3840
      Width           =   2775
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
      Left            =   7800
      TabIndex        =   9
      Top             =   4320
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
      Left            =   7800
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblHidden 
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "(Leave blank if no expiry.)"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Expiry Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Commencement Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Contract No:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Customer ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmConsignment_Main.frx":0004
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   11
      Top             =   120
      Width           =   8175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmConsignment_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TContract
    contractNo As String
    customerID As String
    startDate As String
    endDate As String
End Type

Private tempContract As TContract, currContract As TContract
Dim isAdding As Boolean
Private Sub getContracts()
Dim tempRS As Recordset
Dim tempSQL As String
lvContracts.ListItems.Clear
tempSQL = "SELECT Contracts.ContractNo, Customers.Name, Contracts.StartDate, Contracts.ExpireDate " & _
        "FROM Contracts INNER JOIN Customers ON Contracts.CustomerID = Customers.CustomerID;"

RSOpen tempRS, tempSQL, dbOpenSnapshot
While Not tempRS.EOF
    With lvContracts
        .ListItems.add , , tempRS("ContractNo")
        .ListItems(.ListItems.Count).SubItems(1) = tempRS("Name")
        .ListItems(.ListItems.Count).SubItems(2) = tempRS("StartDate")
        .ListItems(.ListItems.Count).SubItems(3) = tempRS("ExpireDate")
    End With
    tempRS.MoveNext
Wend
tempRS.Close
Set tempRS = Nothing
End Sub
Private Sub setListFormat()
With lvContracts
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.add , , "Contract No", 1150
    .ColumnHeaders.add , , "Customer ID", 5200
    .ColumnHeaders.add , , "Start Date", 960
    .ColumnHeaders.add , , "Expiry Date", 1150
End With
End Sub

Private Sub setFormMode(ByVal strModeStatus As ModeStatus)
Select Case strModeStatus
    Case Editing
        txtContract.Enabled = True
        cmbCustomer.Enabled = True
        txtStart.Enabled = True
        txtEnd.Enabled = True
        lvContracts.Enabled = False
        cmdEdit.Visible = False
        cmdDelete.Visible = False
        cmdClose.Visible = False
        cmdNew.Visible = False
    Case Viewing
        txtContract.Enabled = False
        cmbCustomer.Enabled = False
        txtStart.Enabled = False
        txtEnd.Enabled = False
        lvContracts.Enabled = True
        cmdDelete.Visible = True
        cmdEdit.Visible = True
        cmdClose.Visible = True
        cmdNew.Visible = True
End Select
End Sub

Private Sub getContractValues()
txtContract.Text = lvContracts.SelectedItem.Text
cmbCustomer.Text = lvContracts.SelectedItem.SubItems(1)
txtStart.Text = lvContracts.SelectedItem.SubItems(2)
txtEnd.Text = lvContracts.SelectedItem.SubItems(3)
End Sub

Private Sub storeValues()
currContract.contractNo = txtContract.Text
currContract.customerID = cmbCustomer.Text
currContract.endDate = txtEnd.Text
currContract.startDate = txtStart.Text
End Sub

Private Sub restoreValues()
With currContract
    txtContract.Text = .contractNo
    If .customerID = "" Then
        cmbCustomer.ListIndex = 0
    Else
        cmbCustomer.Text = .customerID
    End If
    txtStart.Text = .startDate
    txtEnd.Text = .endDate
End With
End Sub

Private Sub showContract()
With currContract
    txtContract.Text = .contractNo
    cmbCustomer.Text = .customerID
    txtStart.Text = .startDate
    txtEnd.Text = .endDate
End With
End Sub
Private Function isSame(ByRef strTempVar As TContract) As Boolean
With strTempVar
    If (.contractNo <> tempContract.contractNo) Or (.customerID <> tempContract.customerID) Or _
    (.endDate <> tempContract.endDate) Or (.startDate <> tempContract.startDate) Then
        isSame = False
    Else
        isSame = True
    End If
End With
End Function

Private Sub cmbCustomer_Click()
If Not cmbCustomer.Text = "" Then
    Dim tempRS As Recordset
    RSOpen tempRS, "SELECT CustomerID FROM Customers WHERE Name='" & cmbCustomer.Text & "'", dbOpenSnapshot
    If Not tempRS.EOF Then
        cmbCustomer.Tag = tempRS("CustomerID")
    End If
    tempRS.Close
    Set tempRS = Nothing
End If
End Sub

Private Sub cmdCancel_Click()
restoreValues
setFormMode Viewing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtContract.Text <> "" Then
    If MsgBox("Are you sure you want to remove this contract?" & vbCrLf & "The entire inventory assigned to this consignment contract will be removed.", vbYesNoCancel + vbQuestion, "Delete contract") = vbYes Then
        Dim delSQL As String
        delSQL = "DELETE * FROM Contracts WHERE ContractNo='" & txtContract.Text & "'"
        MySynonDatabase.Execute delSQL
        insertLog "Consignment No: " & txtContract.Text & " has been deleted."
        InfoMsg "The contract has been successfully removed.", "Record deleted"
        getContracts
    End If
Else
    ValidMsg "There are no existing contracts to be deleted.", "No contract available"
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub cmdEdit_Click()
If lvContracts.ListItems.Count > 0 Then
    If lvContracts.SelectedItem.Selected = True Then
        storeValues
        setFormMode Editing
        isAdding = False
    End If
End If
End Sub

Private Sub cmdNew_Click()
setFormMode Editing
storeValues
txtContract.Enabled = False
'cmbCustomer.Text = ""
txtStart.Text = ""
txtEnd.Text = ""
isAdding = True
End Sub

Private Sub cmdSave_Click()
If cmbCustomer.Text = "" Then
    Err.Clear
    ValidMsg "Please select a customer.", "Missing selection"
    cmbCustomer.SetFocus
ElseIf txtStart.Text = "" Then
    Err.Clear
    ValidMsg "Please enter a beginning date.", "Missing value"
    txtStart.SetFocus
Else
    Dim conRS As Recordset, tmpRS As Recordset
    Dim newID As Long
    BeginTrans
    Set tmpRS = MySynonDatabase.OpenRecordset("SELECT DataValue FROM Misc WHERE DataType='CONSIGN'", dbOpenDynaset, dbDenyRead + dbDenyWrite)
    newID = tmpRS("DataValue")
    If isAdding = True Then
        Set conRS = MySynonDatabase.OpenRecordset("SELECT * FROM Contracts", dbOpenDynaset, dbDenyWrite)
        conRS.AddNew
        conRS("ContractNo") = newID
    Else
        Set conRS = MySynonDatabase.OpenRecordset("SELECT * FROM Contracts WHERE ContractNo='" & lblHidden.Caption & "';", dbOpenDynaset, dbDenyWrite)
        conRS.Edit
        conRS("ContractNo") = txtContract.Text
    End If
    conRS("CustomerID") = cmbCustomer.Tag
    conRS("StartDate") = txtStart.Text
    conRS("ExpireDate") = txtEnd.Text
    conRS.Update
    'Update new key
    tmpRS.Edit
    tmpRS("DataValue") = newID + 1
    tmpRS.Update
    
    'Insert into systems log
    insertLog "Consignment contract no: " & newID & IIf((isAdding = True), " created.", " updated.")
    'Close recordsets and free memory
    tmpRS.Close
    conRS.Close
    CommitTrans
    Set tmpRS = Nothing
    Set conRS = Nothing
    If isAdding = True Then
        InfoMsg "Contract No: " & newID & vbCrLf & "The new contract has been successfully created.", "Record saved"
    Else
        InfoMsg "The contract has been successfully updated.", "Record saved"
    End If
    setFormMode Viewing
    getContracts
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, "An error has occurred while saving the data. Record has not been added into the database. Please try again." & _
    "The error might have occurred because another person has just updated the same record or table."
End If
End Sub

Private Sub Form_Load()
setFormMode Viewing
cmbCustomer.addItem ""
FillCombo cmbCustomer, "SELECT Name FROM Customers", "Name"
setListFormat
getContracts
lblNotes.Caption = "It is strongly advised that these settings are left as default. Only administrators are aware of the changes made here."
End Sub

Private Sub lvContracts_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Selected Then
    getContractValues
    lblHidden.Caption = Item.Text
Else
    lblHidden.Caption = ""
End If
End Sub

Private Sub txtContract_GotFocus()
SelText txtContract
End Sub
Private Sub cmbCustomer_GotFocus()
SelText cmbCustomer
End Sub

Private Sub txtContract_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc("/") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtStart_GotFocus()
SelText txtStart
End Sub
Private Sub txtEnd_GotFocus()
SelText txtEnd
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc("/") Then
    OnlyNum KeyAscii
End If
End Sub
