VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSuppliers 
   Caption         =   "Suppliers"
   ClientHeight    =   7635
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuppliers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10545
   Begin MSComctlLib.ListView list_History 
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TabStrip tb 
      Height          =   1935
      Left            =   0
      TabIndex        =   25
      Top             =   5760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSupplierID 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3240
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1560
      Width           =   7215
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1920
      Width           =   9015
   End
   Begin VB.ComboBox cmbCity 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   3000
      Width           =   3735
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "00000"
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtPhone1 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "000"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtPhone2 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   12
      Text            =   "000"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtFax1 
      Height          =   285
      Index           =   0
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   10
      Text            =   "000"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtFax2 
      Height          =   285
      Index           =   0
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "000"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtPhone1 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "00000000"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtPhone2 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "00000000"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtFax1 
      Height          =   285
      Index           =   1
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "00000000"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtFax2 
      Height          =   285
      Index           =   1
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   15
      Text            =   "00000000"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtTerm 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5400
      Width           =   495
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
      ToolTipText     =   "Click here to edit the selected supplier."
      Top             =   4800
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
      Left            =   9360
      TabIndex        =   22
      ToolTipText     =   "Click here to close this window."
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   16
      Top             =   4560
      Width           =   6495
   End
   Begin MSComCtl2.UpDown udTerm 
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Top             =   5400
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtBalance"
      BuddyDispid     =   196622
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
   Begin MSComctlLib.ListView list_Suppliers 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
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
      BorderStyle     =   1
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
      Left            =   9360
      TabIndex        =   23
      ToolTipText     =   "Click here to save changes."
      Top             =   4800
      Width           =   1095
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
      ToolTipText     =   "Click here to cancel editing."
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "days"
      Height          =   255
      Left            =   4800
      TabIndex        =   43
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier ID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   42
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   41
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   40
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   39
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "State:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   38
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Country:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   36
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phone (1):"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   35
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phone (2):"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   34
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax (1):"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   33
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax (2):"
      Height          =   255
      Index           =   10
      Left            =   4200
      TabIndex        =   32
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Limit:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   31
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Term:"
      Height          =   255
      Index           =   12
      Left            =   2760
      TabIndex        =   30
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Current Balance:"
      Height          =   255
      Index           =   13
      Left            =   5400
      TabIndex        =   29
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   2040
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      X1              =   6000
      X2              =   6120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblHidden 
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4560
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
      Begin VB.Menu Bar01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options_Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Options_Cancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu Bar02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_PO 
      Caption         =   "&Purchase Order"
      Begin VB.Menu mnu_PO_new 
         Caption         =   "&New Purchase Order"
      End
   End
   Begin VB.Menu mnu_Account 
      Caption         =   "&Account"
      Begin VB.Menu mnu_Acc_Adjust 
         Caption         =   "&Adjustment"
      End
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub list_Suppliers_ItemClick(ByVal Item As MSComctlLib.ListItem)
If list_Suppliers.ListItems.Count > 0 Then
    If list_Suppliers.SelectedItem.Selected = True Then
        getSupplierInfo list_Suppliers.SelectedItem.Text
    End If
End If
End Sub

Private Sub mnu_Acc_Adjust_Click()
If lblhidden.Caption <> "" Then
    Load frmAdjustment
    frmAdjustment.setAccType "Supplier"
    frmAdjustment.Tag = lblhidden.Caption
    frmAdjustment.Show vbModal
End If
End Sub

Private Sub mnu_Account_Click()
If lblhidden.Caption = "" Then
    mnu_Acc_Adjust.Enabled = False
Else
    mnu_Acc_Adjust.Enabled = True
End If
End Sub

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
getSupplierInfo lblhidden.Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If Len(txtSupplierID.Text) > 0 Then
    FormMode Editing
Else
    InfoMsg "Please select a supplier first.", "Missing selection"
End If
End Sub

Private Sub cmdSave_Click()
'Do nothing
If txtSupplierID.Text = "" Then
    ValidMsg "Please enter a supplier ID.", "Missing supplier ID"
    txtSupplierID.SetFocus
ElseIf txtName.Text = "" Then
    ValidMsg "Please enter a supplier name.", "Missing name"
    txtName.SetFocus
ElseIf txtAddress.Text = "" Then
    ValidMsg "Please enter an address for the supplier.", "Missing address"
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
    Dim suppRS As Recordset
    RSOpen suppRS, "SELECT * FROM Suppliers WHERE SupplierID='" & lblhidden.Caption & "';", dbOpenDynaset
    If Not suppRS.EOF Then
        suppRS.Edit
        suppRS("SupplierID") = txtSupplierID.Text
        suppRS("Name") = txtName.Text
        suppRS("Address") = txtAddress.Text
        suppRS("Country") = cmbCountry.Text
        suppRS("City") = cmbCity.Text
        suppRS("State") = cmbState.Text
        suppRS("Zip") = txtZip.Text
        suppRS("ACPhone1") = txtPhone1(0).Text
        suppRS("ACPhone2") = txtPhone2(0).Text
        suppRS("ACFax1") = txtFax1(0).Text
        suppRS("ACFax2") = txtFax2(0).Text
        suppRS("Phone1") = txtPhone1(1).Text
        suppRS("Phone2") = txtPhone2(1).Text
        suppRS("Fax1") = txtFax1(1).Text
        suppRS("Fax2") = txtFax2(1).Text
        suppRS("Email") = txtEmail.Text
        suppRS("CreditLimit") = Format$(txtLimit.Text, "#,##0.00")
        suppRS("CreditTerm") = txtTerm.Text
        suppRS.Update
        
        'Insert into systems log
        insertLog "Supplier ID: " & lblhidden.Caption & " account has been updated."
        InfoMsg "Supplier ID: " & lblhidden.Caption & " account has been successfully updated.", "Record saved"
        FormMode Viewing
    End If
    'Close recordsets and free memory
    suppRS.Close
    Set suppRS = Nothing
End If
End Sub

Private Sub Form_Load()
FillComboCountry cmbCountry
FormMode Viewing
Me.WindowState = vbMaximized
getSuppliers
With tb
    .Tabs.Clear
    .Tabs.add , , "Purchases"
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
list_Suppliers.width = Me.ScaleWidth - list_Suppliers.Left * 2
list_History.width = Me.ScaleWidth - list_History.Left * 2
tb.width = Me.ScaleWidth - tb.Left * 2
tb.height = Me.ScaleHeight - tb.Left * 5
list_History.height = Me.ScaleHeight - list_History.Left * 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSuppliers = Nothing
End Sub

Private Sub list_Suppliers_DblClick()
If list_Suppliers.ListItems.Count > 0 Then
    If list_Suppliers.SelectedItem.Selected = True Then
        getSupplierInfo list_Suppliers.SelectedItem.Text
    End If
End If
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
frmSupplier_New.Show vbModal
End Sub

Private Sub mnu_Options_Save_Click()
Call cmdSave_Click
End Sub

Private Sub mnu_PO_new_Click()
frmPurchase.Show , frmMain
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

Private Sub txtName_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

Private Sub txtSupplierID_GotFocus()
SelText txtSupplierID
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
SelText txtPhone2
End Sub

Private Sub txtPhone2_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtSupplierID_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
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

Private Sub tb_Click()
If tb.SelectedItem.Selected = True Then
    getSupplierHistory tb.SelectedItem.Caption
End If
End Sub

Public Sub getSuppliers()
'Obtain list of Suppliers and display them on the list view control
'Format list
With list_Suppliers
    .View = lvwReport
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Supplier ID"
    .ColumnHeaders.add , , "Company Name", 4000
    
    'Define query
    Dim suppSQL As String
    suppSQL = "SELECT Suppliers.SupplierID, Suppliers.Name, Suppliers.CurrentBalance, Suppliers.CreditLimit FROM Suppliers ORDER BY Suppliers.SupplierID;"
    
    Dim suppRS As Recordset
    
    On Error GoTo ErrHandler
    'Open recordset
    RSOpen suppRS, suppSQL, dbOpenSnapshot
    While Not suppRS.EOF
        'Run through all records and add them to list
        .ListItems.add , , suppRS("SupplierID") ' (IIf((suppRS("CurrentBalance") > suppRS("CreditLimit")) , imgList.Image(1), imgList.Image(2)))
        .ListItems(.ListItems.Count).SubItems(1) = suppRS("Name")
        suppRS.MoveNext
    Wend
    suppRS.Close
    Set suppRS = Nothing

FormMode Viewing
End With

ErrHandler:
If Err.Number <> 0 Then
    Screen.MousePointer = 0
    CriticalMsg "Unable to load list of Suppliers. Please contact system administrator.", "Warning"
    Exit Sub
End If
End Sub

Private Sub FormMode(strModeName As ModeStatus)
If strModeName = Editing Then
    list_Suppliers.Enabled = False
    txtSupplierID.Enabled = True
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
Else
    list_Suppliers.Enabled = True
    txtSupplierID.Enabled = False
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
End If
End Sub

Private Sub getSupplierInfo(ByVal strSupplierID As String)
'Obtains the particular Supplier information and place them into respective controls
Dim InfoSQL As String
InfoSQL = "SELECT * FROM Suppliers WHERE Suppliers.SupplierID='" & strSupplierID & "';"

Dim InfoRS As Recordset

On Error Resume Next
RSOpen InfoRS, InfoSQL, dbOpenSnapshot
If Not InfoRS.EOF Then
    txtSupplierID.Text = strSupplierID
    lblhidden.Caption = InfoRS("SupplierID")
    txtName.Text = InfoRS("Name")
    txtAddress.Text = InfoRS("Address")
    cmbCountry.Text = InfoRS("Country")
    cmbState.Text = InfoRS("State")
    cmbCity.Text = InfoRS("City")
    txtZip.Text = InfoRS("Zip")
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

Private Sub FieldCheck()
If (Len(txtSupplierID.Text) = 0) Or (Len(txtName.Text) = 0) Or (Len(txtAddress.Text) = 0) Or _
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
Private Sub getSupplierHistory(ByVal strSection As String)
'Gets the history of the customer based on the parameter passed
Dim hisSQL As String, strCon As String
Dim hisRS As Recordset
strCon = ""
Select Case strSection
    Case "Payment"
        strCon = "credit"
        hisSQL = "SELECT notes, date, credit FROM supp_transactions WHERE SupplierID='" & lblhidden.Caption & "'"
    Case "Purchases"
        strCon = "debit"
        hisSQL = "SELECT notes, date, debit FROM supp_transactions WHERE SupplierID='" & lblhidden.Caption & "'"
    Case "All"
        strCon = "All"
        hisSQL = "SELECT notes,date,credit, debit FROM supp_transactions WHERE SupplierID='" & lblhidden.Caption & "'"
End Select
With list_History
    .View = lvwReport
    'Clear the history contents
    .ColumnHeaders.Clear
    .ListItems.Clear
    'Re-format the properties
    .ColumnHeaders.add , , "Date"
    .ColumnHeaders.add , , "Description", 4470
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
                currBalance = currBalance + hisRS("credit") - hisRS("debit")
                .Item(.Count).SubItems(4) = Format$(currBalance, "#,##0.00")
            ElseIf strCon = "debit" Then
                If hisRS(strCon) > 0 Then
                    .add , , hisRS("date")
                    .Item(.Count).SubItems(1) = hisRS("notes")
                    .Item(.Count).SubItems(2) = Format$(hisRS("debit"), "#,##0.00")
                End If
            Else
                If hisRS(strCon) > 0 Then
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
