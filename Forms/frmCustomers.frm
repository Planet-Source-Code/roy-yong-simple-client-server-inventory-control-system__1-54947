VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomers 
   Caption         =   "Customers"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10560
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   15
      Top             =   4560
      Width           =   6495
   End
   Begin MSComctlLib.ListView list_History 
      Height          =   1695
      Left            =   120
      TabIndex        =   39
      Top             =   6120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2990
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
      Left            =   9360
      TabIndex        =   22
      ToolTipText     =   "Click here to close this window."
      Top             =   5280
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
      Left            =   9360
      TabIndex        =   21
      ToolTipText     =   "Click here to edit the selected customer."
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtBalance 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   5400
      Width           =   975
   End
   Begin MSComCtl2.UpDown udTerm 
      Height          =   285
      Left            =   4335
      TabIndex        =   18
      Top             =   5400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtTerm"
      BuddyDispid     =   196613
      OrigLeft        =   4440
      OrigTop         =   4560
      OrigRight       =   4680
      OrigBottom      =   4815
      Increment       =   30
      Max             =   360
      Min             =   30
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTerm 
      Height          =   285
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtFax2 
      Height          =   285
      Index           =   1
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   14
      Text            =   "00000000"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtFax1 
      Height          =   285
      Index           =   1
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "00000000"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtPhone2 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "00000000"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtPhone1 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "00000000"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtFax2 
      Height          =   285
      Index           =   0
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   13
      Text            =   "000"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtFax1 
      Height          =   285
      Index           =   0
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "000"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtPhone2 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "000"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtPhone1 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "000"
      Top             =   3840
      Width           =   495
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.ComboBox cmbCity 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1920
      Width           =   9015
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3240
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1560
      Width           =   7215
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin MSComctlLib.ListView list_Customers 
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2355
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
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
      Left            =   9360
      TabIndex        =   24
      ToolTipText     =   "Click here to cancel any changes."
      Top             =   5280
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
      Left            =   9360
      TabIndex        =   23
      ToolTipText     =   "Click here to save any changes."
      Top             =   4800
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tb 
      Height          =   2175
      Left            =   0
      TabIndex        =   42
      Top             =   5760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "days"
      Height          =   255
      Left            =   4680
      TabIndex        =   43
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblHidden 
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   6000
      X2              =   6120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   2040
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "Current Balance:"
      Height          =   255
      Index           =   13
      Left            =   5400
      TabIndex        =   38
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Term:"
      Height          =   255
      Index           =   12
      Left            =   2760
      TabIndex        =   37
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Limit:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   36
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax (2):"
      Height          =   255
      Index           =   10
      Left            =   4200
      TabIndex        =   35
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax (1):"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   34
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phone (2):"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   33
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phone (1):"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   32
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   31
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Country:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   30
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "State:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   26
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Customer ID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Options_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_Options_Edit 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_Bar_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options_Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Options_Cancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnu_Bar_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_DO 
      Caption         =   "&Delivery Order"
      Begin VB.Menu mnu_DO_New 
         Caption         =   "&New Delivery Order"
      End
      Begin VB.Menu mnu_DO_Payment 
         Caption         =   "&Payment"
      End
   End
   Begin VB.Menu mnu_Account 
      Caption         =   "&Account"
      Begin VB.Menu mnu_Acc_Adjust 
         Caption         =   "&Adjustments"
      End
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCity_GotFocus()
If cmbCity.Text = "[PLEASE SELECT ONE]" Then
    cmbCity.Text = ""
End If
SelText cmbCity
End Sub

Private Sub cmbCity_LostFocus()
CapCon cmbCity
End Sub

Private Sub cmbCountry_Click()
FillComboState cmbState, cmbCountry.Text
End Sub

Private Sub cmbCountry_GotFocus()
If cmbCountry.Text = "[PLEASE SELECT ONE]" Then
    cmbCountry.Text = ""
End If
SelText cmbCountry
End Sub

Private Sub cmbCountry_LostFocus()
CapCon cmbCountry
End Sub

Private Sub cmbState_GotFocus()
If cmbState.Text = "[PLEASE SELECT ONE]" Then
    cmbState.Text = ""
End If
SelText cmbState
End Sub

Private Sub cmbState_LostFocus()
CapCon cmbState
End Sub

Private Sub cmdCancel_Click()
FormMode Viewing
getCustomerInfo lblHidden.Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If txtCustomerID.Text <> "" Then
    FormMode Editing
Else
    InfoMsg "Please select a customer first.", "Missing selection"
End If
End Sub

Private Sub cmdSave_Click()
'Do nothing
If txtCustomerID.Text = "" Then
    ValidMsg "Please enter a customer ID.", "Missing customer ID"
    txtCustomerID.SetFocus
ElseIf txtName.Text = "" Then
    ValidMsg "Please enter a customer name.", "Missing name"
    txtName.SetFocus
ElseIf txtAddress.Text = "" Then
    ValidMsg "Please enter an address for the customer.", "Missing address"
    txtAddress.SetFocus
ElseIf cmbCountry.Text = "" Then
    ValidMsg "Please select or enter a country for the address.", "Missing country"
    cmbCountry.SetFocus
ElseIf cmbState.Text = "" Then
    ValidMsg "Please select or enter a state for the address.", "Missing state"
    cmbState.SetFocus
ElseIf cmbCity.Text = "" Then
    ValidMsg "Please select or enter a city for the address.", "Missing city"
    cmbCity.SetFocus
ElseIf txtZip.Text = "" Then
    ValidMsg "Please enter a zip code for the address.", "Missing zip code"
    txtZip.SetFocus
ElseIf txtPhone1(0).Text = "" Then
    ValidMsg "Please enter the a valid phone number or leave it in 0s.", "Missing phone number"
    txtPhone1(0).SetFocus
ElseIf txtPhone1(1).Text = "" Then
    ValidMsg "Please enter the a valid phone number or leave it in 0s.", "Missing phone number"
    txtPhone1(1).SetFocus
ElseIf txtPhone2(0).Text = "" Then
    ValidMsg "Please enter the a valid phone number or leave it in 0s.", "Missing phone number"
    txtPhone2(0).SetFocus
ElseIf txtPhone2(1).Text = "" Then
    ValidMsg "Please enter the a valid phone number or leave it in 0s.", "Missing phone number"
    txtPhone1(1).SetFocus
ElseIf txtFax1(0).Text = "" Then
    ValidMsg "Please enter the a valid fax number or leave it in 0s.", "Missing fax number"
    txtFax1(0).SetFocus
ElseIf txtFax1(1).Text = "" Then
    ValidMsg "Please enter the a valid fax number or leave it in 0s.", "Missing fax number"
    txtFax1(1).SetFocus
ElseIf txtFax2(0).Text = "" Then
    ValidMsg "Please enter the a valid fax number or leave it in 0s.", "Missing fax number"
    txtFax2(0).SetFocus
ElseIf txtFax2(1).Text = "" Then
    ValidMsg "Please enter the a valid fax number or leave it in 0s.", "Missing fax number"
    txtFax2(1).SetFocus
ElseIf Val(txtLimit.Text) > 999999 Then
    ValidMsg "Please ensure that the credit limit do not exceed $999,999.00.", "Invalid credit limit"
    txtLimit.SetFocus
Else
    'Done with validation
    'Continue with saving
    Dim custRS As Recordset
    RSOpen custRS, "SELECT * FROM Customers WHERE CustomerID='" & lblHidden.Caption & "';", dbOpenDynaset
    If Not custRS.EOF Then
        custRS.Edit
        custRS("CustomerID") = txtCustomerID.Text
        custRS("Name") = txtName.Text
        custRS("Address") = txtAddress.Text
        custRS("Country") = cmbCountry.Text
        custRS("City") = cmbCity.Text
        custRS("State") = cmbState.Text
        custRS("Zip") = txtZip.Text
        custRS("ACPhone1") = txtPhone1(0).Text
        custRS("ACPhone2") = txtPhone2(0).Text
        custRS("ACFax1") = txtFax1(0).Text
        custRS("ACFax2") = txtFax2(0).Text
        custRS("Phone1") = txtPhone1(1).Text
        custRS("Phone2") = txtPhone2(1).Text
        custRS("Fax1") = txtFax1(1).Text
        custRS("Fax2") = txtFax2(1).Text
        custRS("Email") = txtEmail.Text
        custRS("CreditLimit") = Format$(txtLimit.Text, "#,##0.00")
        custRS("CreditTerm") = txtTerm.Text
        custRS.Update
        
        'Insert into systems log
        insertLog "Customer ID: " & lblHidden.Caption & " account has been updated."
        InfoMsg "Customer ID: " & lblHidden.Caption & " account has been successfully updated.", "Record saved"
        FormMode Viewing
    End If
    'Close recordsets and free memory
    custRS.Close
    Set custRS = Nothing
End If
End Sub

Private Sub Form_Load()
FillComboCountry cmbCountry
FormMode Viewing
Me.WindowState = vbMaximized
getCustomers
DisableClose frmCustomers, True
With tb
    .Tabs.Clear
    .Tabs.add , , "Delivery"
    .Tabs.add , , "Payment"
    .Tabs.add , , "All"
    .Tabs(1).Selected = True
End With
If CurrentUser.prvlgAdmin = True Then
    mnu_Account.Visible = True
Else
    mnu_Account.Visible = False
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
list_Customers.width = Me.ScaleWidth - list_Customers.Left * 2
list_History.width = Me.ScaleWidth - list_History.Left * 2
tb.width = Me.ScaleWidth - tb.Left * 2
tb.height = Me.ScaleHeight - tb.Left * 5
list_History.height = Me.ScaleHeight - list_History.Left * 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCustomers = Nothing
End Sub

Private Sub list_Customers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With list_Customers '// change to the name of the list view
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    .SortKey = iCur
    iLast = iCur
End With
End Sub

Private Sub list_Customers_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
If .Selected Then
    getCustomerInfo .Text
End If
End With
End Sub

Private Sub mnu_Acc_Adjust_Click()
If lblHidden.Caption <> "" Then
    Load frmAdjustment
    frmAdjustment.setAccType "Customer"
    frmAdjustment.Tag = lblHidden.Caption
    frmAdjustment.Show vbModal
End If
End Sub

Private Sub mnu_Account_Click()
If lblHidden.Caption = "" Then
    mnu_Acc_Adjust.Enabled = False
Else
    mnu_Acc_Adjust.Enabled = True
End If
End Sub

Private Sub mnu_DO_New_Click()
frmDelivery.Show , frmMain
End Sub

Private Sub mnu_DO_Payment_Click()
frmCustomer_Payment.Show vbModal
End Sub

Private Sub mnu_Options_Cancel_Click()
Call cmdCancel_Click
End Sub

Private Sub mnu_Options_Edit_Click()
Call cmdEdit_Click
End Sub

Private Sub mnu_Options_Exit_Click()
Call cmdClose_Click
End Sub

Private Sub mnu_Options_New_Click()
frmCustomer_New.Show vbModal
End Sub

Private Sub mnu_Options_Save_Click()
Call cmdSave_Click
End Sub

Private Sub tb_Click()
If tb.SelectedItem.Selected = True Then
    getCustomerHistory tb.SelectedItem.Caption
End If
End Sub

Private Sub txtAddress_GotFocus()
SelText txtAddress
End Sub

Private Sub txtAddress_LostFocus()
CapCon txtAddress
End Sub

Private Sub txtBalance_GotFocus()
SelText txtBalance
End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtCustomerID_GotFocus()
SelText txtCustomerID
End Sub

Private Sub txtFax1_GotFocus(Index As Integer)
SelText txtFax1(Index)
End Sub

Private Sub txtFax1_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtFax2_GotFocus(Index As Integer)
SelText txtFax2(Index)
End Sub

Private Sub txtFax2_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtLimit_GotFocus()
SelText txtLimit
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtName_GotFocus()
SelText txtName
End Sub

Private Sub txtName_LostFocus()
CapCon txtName
End Sub

Private Sub txtPhone1_GotFocus(Index As Integer)
SelText txtPhone1(Index)
End Sub

Private Sub txtPhone1_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtPhone2_GotFocus(Index As Integer)
SelText txtPhone2(Index)
End Sub

Private Sub txtPhone2_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtTerm_GotFocus()
SelText txtTerm
End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtZip_GotFocus()
SelText txtZip
End Sub

Private Sub txtZip_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Public Sub getCustomers()
'Obtain list of customers and display them on the list view control
'Format list
With list_Customers
    .View = lvwReport
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Customer ID"
    .ColumnHeaders.add , , "Company Name", 4000
'Define query
Dim custSQL As String
custSQL = "SELECT Customers.CustomerID, Customers.Name, Customers.CurrentBalance, Customers.CreditLimit FROM Customers ORDER BY Customers.CustomerID;"

Dim custRS As Recordset

On Error GoTo ErrHandler
'Open recordset
RSOpen custRS, custSQL, dbOpenSnapshot
While Not custRS.EOF
    'Run through all records and add them to list
    .ListItems.add , , custRS("CustomerID") ' (IIf((custRS("CurrentBalance") > custRS("CreditLimit")) , imgList.Image(1), imgList.Image(2)))
    .ListItems(.ListItems.Count).SubItems(1) = custRS("Name")
    custRS.MoveNext
Wend
custRS.Close
Set custRS = Nothing
End With
FormMode Viewing

ErrHandler:
If Err.Number <> 0 Then
    Screen.MousePointer = 0
    CriticalMsg "Unable to load list of customers. Please contact system administrator.", "Warning"
    Exit Sub
End If
End Sub

Private Sub FormMode(ModeName As ModeStatus)
'sets the current mode of the form
Select Case ModeName
Case Editing
    list_Customers.Enabled = False
    txtCustomerID.Enabled = True
    txtName.Enabled = True
    txtAddress.Enabled = True
    cmbCountry.Enabled = True
    cmbState.Enabled = True
    cmbCity.Enabled = True
    txtZip.Enabled = True
    txtPhone1(0).Enabled = True
    txtPhone1(1).Enabled = True
    txtPhone2(0).Enabled = True
    txtPhone2(1).Enabled = True
    txtFax1(0).Enabled = True
    txtFax1(1).Enabled = True
    txtFax2(0).Enabled = True
    txtFax2(1).Enabled = True
    txtLimit.Enabled = True
    txtEmail.Enabled = True
    udTerm.Enabled = True
    cmdEdit.Visible = False
    cmdClose.Visible = False
    mnu_Options_Exit.Enabled = False
    mnu_Options_Edit.Enabled = False
    mnu_Options_New.Enabled = False
    mnu_Options_Save.Enabled = True
    mnu_Options_Cancel.Enabled = True
Case Viewing
    list_Customers.Enabled = True
    txtCustomerID.Enabled = False
    txtName.Enabled = False
    txtAddress.Enabled = False
    cmbCountry.Enabled = False
    cmbState.Enabled = False
    cmbCity.Enabled = False
    txtZip.Enabled = False
    txtPhone1(0).Enabled = False
    txtPhone1(1).Enabled = False
    txtPhone2(0).Enabled = False
    txtPhone2(1).Enabled = False
    txtFax1(0).Enabled = False
    txtFax1(1).Enabled = False
    txtFax2(0).Enabled = False
    txtFax2(1).Enabled = False
    txtLimit.Enabled = False
    txtEmail.Enabled = False
    udTerm.Enabled = False
    cmdEdit.Visible = True
    cmdClose.Visible = True
    mnu_Options_Exit.Enabled = True
    mnu_Options_Edit.Enabled = True
    mnu_Options_New.Enabled = True
    mnu_Options_Save.Enabled = False
    mnu_Options_Cancel.Enabled = False
End Select
End Sub

Private Sub getCustomerInfo(ByVal strCustomerID As String)
'Obtains the particular customer information and place them into respective controls
Dim InfoSQL As String
InfoSQL = "SELECT * FROM Customers WHERE Customers.CustomerID='" & strCustomerID & "';"

Dim InfoRS As Recordset

On Error Resume Next
RSOpen InfoRS, InfoSQL, dbOpenSnapshot
If Not InfoRS.EOF Then
    txtCustomerID.Text = strCustomerID
    lblHidden.Caption = InfoRS("CustomerID")
    txtName.Text = InfoRS("Name")
    txtAddress.Text = InfoRS("Address")
    cmbCountry.Text = InfoRS("Country")
    cmbState.Text = InfoRS("State")
    cmbCity.Text = InfoRS("City")
    txtZip.Text = IIf(IsNull(InfoRS("Zip")), "", InfoRS("Zip"))
    txtPhone1(0).Text = InfoRS("ACPhone1")
    txtPhone1(1).Text = InfoRS("Phone1")
    txtPhone2(0).Text = InfoRS("ACPhone2")
    txtPhone2(1).Text = InfoRS("Phone2")
    txtFax1(0).Text = InfoRS("ACFax1")
    txtFax1(1).Text = InfoRS("Fax1")
    txtFax2(0).Text = InfoRS("ACFax2")
    txtFax2(1).Text = InfoRS("Fax2")
    txtEmail.Text = IIf(IsNull(InfoRS("Email")), "", InfoRS("Email"))
    txtLimit.Text = Format$(InfoRS("CreditLimit"), "#,##0.00")
    txtTerm.Text = InfoRS("CreditTerm")
    txtBalance.Text = Format$(InfoRS("CurrentBalance"), "#,##0.00")
End If
InfoRS.Close
Set InfoRS = Nothing
list_History.ListItems.Clear
tb.Tabs(0).Selected = True
End Sub

Private Sub getCustomerHistory(ByVal strSection As String)
'Gets the history of the customer based on the parameter passed
Dim hisSQL As String, strCon As String
Dim hisRS As Recordset
strCon = ""
Select Case strSection
    Case "Payment"
        strCon = "credit"
        hisSQL = "SELECT notes, date, credit FROM cust_transactions WHERE CustomerID='" & lblHidden.Caption & "'"
    Case "Delivery"
        strCon = "debit"
        hisSQL = "SELECT notes, date, debit FROM cust_transactions WHERE CustomerID='" & lblHidden.Caption & "'"
    Case "All"
        strCon = "All"
        hisSQL = "SELECT notes,date,debit,credit FROM cust_transactions WHERE CustomerID='" & lblHidden.Caption & "'"
End Select
With list_History
    .View = lvwReport
    'Clear the history contents
    .ColumnHeaders.Clear
    .ListItems.Clear
    'Re-format the properties
    .ColumnHeaders.add , , "Date"
    .ColumnHeaders.add , , "Description"
    .ColumnHeaders(2).width = 4470
    If strSection = "All" Then
        .ColumnHeaders.add , , "Debit" '2
        .ColumnHeaders.add , , "Credit" '3
        .ColumnHeaders.add , , "Balance" '4
    Else
        .ColumnHeaders.add , , "Amount" '2
    End If
End With
Dim currBalance As Double
currBalance = 0
If strCon <> "" Then
    On Error GoTo ErrHandler
    RSOpen hisRS, hisSQL, dbOpenSnapshot
    While Not hisRS.EOF
        With list_History.ListItems
            If strCon = "All" Then
                .add , , hisRS("date")
                .Item(.Count).SubItems(1) = hisRS("notes")
                .Item(.Count).SubItems(2) = Format$(hisRS("debit"), "#,##0.00")
                .Item(.Count).SubItems(3) = Format$(hisRS("credit"), "#,##0.00")
                currBalance = currBalance + hisRS("debit") - hisRS("credit")
                .Item(.Count).SubItems(4) = Format$(currBalance, "#,##0.00")
                
            ElseIf strCon = "debit" Then
                If hisRS(strCon) > 0 Then
                    .add , , hisRS("date")
                    .Item(.Count).SubItems(1) = hisRS("notes")
                    .Item(.Count).SubItems(2) = Format$(hisRS("debit"), "#,##0.00")
                End If
            Else
                If hisRS(strCon) Then
                    .add , , hisRS("date")
                    .Item(.Count).SubItems(1) = hisRS("notes")
                    .Item(.Count).SubItems(2) = Format$(hisRS("credit"), "#,##0.00")
                End If
            End If
        End With
        hisRS.MoveNext
    Wend
    hisRS.Close
    Set hisRS = Nothing
End If
ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
    Exit Sub
End If
End Sub
Private Sub FieldCheck()
If (Len(txtCustomerID.Text) = 0) Or (Len(txtName.Text) = 0) Or (Len(txtAddress.Text) = 0) Or _
(Len(cmbCountry.Text) = 0) Or (Len(cmbState.Text) = 0) Or (Len(cmbCity.Text) = 0) Or (Len(txtZip.Text) = 0) Or _
(Len(txtPhone1(0).Text) = 0) Or (Len(txtPhone1(1).Text) = 0) Or (Len(txtPhone2(0).Text) = 0) Or (Len(txtPhone2(1).Text) = 0) Or _
(Len(txtFax1(0).Text) = 0) Or (Len(txtFax1(1).Text) = 0) Or (Len(txtFax2(0).Text) = 0) Or (Len(txtFax2(1).Text) = 0) Or _
(Len(txtLimit.Text) = 0) Then
    cmdSave.Enabled = False
    mnu_Options_Save.Enabled = False
Else
    cmdSave.Enabled = True
    mnu_Options_Save.Enabled = True
End If
End Sub

Private Sub printInvoice(ByVal strCustID As String)
Dim tSQL As String
tSQL = "SELECT Delivery.DOnumber, Delivery.Date, Delivery.Status, Sum(([D_Details].[Quantity]*[D_Details].[SalePrice])+[Delivery].[Charges]) AS Total " & _
"FROM (Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID) INNER JOIN D_Details ON Delivery.DOnumber = D_Details.DOnumber " & _
"Where (((Customers.CustomerID) = '" & strCustID & "') And ((D_Details.isInvoiced) = True)) " & _
"GROUP BY Delivery.DOnumber, Delivery.Date, Delivery.Status " & _
"HAVING (((Delivery.Status)='INVOICED'));"


End Sub
