VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduct_Browse 
   Caption         =   "Inventory Control"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   1275
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
   Icon            =   "frmProduct_Browse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10560
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
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
      Left            =   9600
      TabIndex        =   6
      ToolTipText     =   "Click here to begin adjustment."
      Top             =   6600
      Width           =   615
   End
   Begin VB.OptionButton optIn 
      Alignment       =   1  'Right Justify
      Caption         =   "In / +"
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      ToolTipText     =   "Select this if adding stock to the selected product."
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optOut 
      Alignment       =   1  'Right Justify
      Caption         =   "Out / -"
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      ToolTipText     =   "Select this if deducting stock of the selected product."
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtering:"
      Height          =   1335
      Left            =   8400
      TabIndex        =   22
      Top             =   120
      Width           =   2055
      Begin VB.ComboBox cmbFilter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the categories."
         Top             =   480
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar pbBar 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "By Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar bar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   7185
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Product Info:"
      Height          =   3615
      Left            =   8400
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
      Begin VB.TextBox lblQuantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox lblReorder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox lblMin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblUnitPrice 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Price of the selected product."
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblLocation 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Location of the selected product."
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Available Quantity:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Reorder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Minimum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblCategoryID 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Category ID of the selected product."
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblBrand 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDescription 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Description of the selected product."
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Product ID of the selected product."
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lvInventory 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9763
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjustments:"
      Height          =   1815
      Left            =   8400
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   5
         ToolTipText     =   "Enter the amount of stock for this transaction of the selected product."
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Refresh 
         Caption         =   "&Refresh List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_Options_Add 
         Caption         =   "Add &New Product"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_Bar_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Options_Close 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnu_Delivery 
      Caption         =   "&Delivery"
      Begin VB.Menu mnu_New 
         Caption         =   "New &Delivery Order"
      End
      Begin VB.Menu mnu_Delivery_Add 
         Caption         =   "&Add To Delivery Order"
         Begin VB.Menu mnu_ExistingDO 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnu_Bar_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_New_PO 
         Caption         =   "New &Purchase Order"
      End
      Begin VB.Menu mnu_Purchase_Add 
         Caption         =   "&Add To Purchase Order"
         Begin VB.Menu mnu_ExistingPO 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnu_bar_03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Consign 
         Caption         =   "&Insert Into Consignment"
         Begin VB.Menu mnu_Insert 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmProduct_Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
End Function

Public Sub formatListView()
With lvInventory
    .View = lvwReport
    .ColumnHeaders.add , , "ID", 960
    .ColumnHeaders.add , , "Description", 3500
    .ColumnHeaders.add , , "Brand", 1500
    .ColumnHeaders.add , , "Category ID", 1500
    .ColumnHeaders.add , , "Quantity", 900
    .ColumnHeaders.add , , "Min Level", 900
    .ColumnHeaders.add , , "Reorder Level", 1200
    .ColumnHeaders.add , , "Location", 1500
    If getSettings("allowPrice") = "TRUE" Then
        .ColumnHeaders.add , , "Unit Price", 900
    End If
End With

End Sub
Public Sub showListofProducts(ByVal paramSQL As String)
lvInventory.ListItems.Clear

Dim listRS As Recordset

RSOpen listRS, paramSQL, dbOpenSnapshot
'On Error GoTo ErrHandler
While Not listRS.EOF
    setPercentage listRS.PercentPosition
    With lvInventory
        .ListItems.add , , listRS("ProductID")
        .ListItems(.ListItems.Count).SubItems(1) = IIf(IsNull(listRS("Description")), "", listRS("Description"))
        .ListItems(.ListItems.Count).SubItems(2) = IIf(IsNull(listRS("Brand")), "", listRS("Brand"))
        .ListItems(.ListItems.Count).SubItems(3) = listRS("CategoryID")
        .ListItems(.ListItems.Count).SubItems(4) = IIf(IsNull(listRS("Quantity")), "0", listRS("Quantity"))
        .ListItems(.ListItems.Count).SubItems(5) = IIf(IsNull(listRS("MinLevel")), "0", listRS("MinLevel"))
        .ListItems(.ListItems.Count).SubItems(6) = IIf(IsNull(listRS("ReorderLevel")), "0", listRS("ReorderLevel"))
        .ListItems(.ListItems.Count).SubItems(7) = IIf(IsNull(listRS("Location")), "", listRS("Location"))
        If getSettings("allowPrice") = "TRUE" Then 'If allowed to see price
            .ListItems(.ListItems.Count).SubItems(8) = Format$(IIf(IsNull(listRS("UnitPrice")), "0", listRS("UnitPrice")), "#,##0.00")
        End If
    End With
    listRS.MoveNext
Wend
listRS.Close
Set listRS = Nothing
setPercentage 0

ErrHandler:
If Err.Number <> 0 Then
    'Possible errors occuring during runtime
    ErrorNotifier Err.Number, Err.description
    setPercentage 0
End If
End Sub

Private Sub setPercentage(ByVal valPercent As Single)
pbBar.Value = valPercent
If valPercent = 0 Then
    lblPercent.Caption = ""
Else
    lblPercent.Caption = Format$(valPercent, "00.00\%")
End If
End Sub

Public Sub getValues()
'Obtain values from list view for further references
With lvInventory
    lblID.Caption = .SelectedItem.Text
    lblDescription.Caption = .SelectedItem.SubItems(1)
    lblBrand.Caption = .SelectedItem.SubItems(2)
    lblCategoryID.Caption = .SelectedItem.SubItems(3)
    lblQuantity.Text = .SelectedItem.SubItems(4)
    lblMin.Text = .SelectedItem.SubItems(5)
    lblReorder.Text = .SelectedItem.SubItems(6)
    lblLocation.Caption = .SelectedItem.SubItems(7)
    If getSettings("allowPrice") = "TRUE" Then
        lblUnitPrice.Caption = .SelectedItem.SubItems(8)
    End If
End With
End Sub

Public Sub clearValues()
'Clears the values within product info section
lblID.Caption = ""
lblDescription.Caption = ""
lblBrand.Caption = ""
lblCategoryID.Caption = ""
lblQuantity.Text = ""
lblMin.Text = ""
lblReorder.Text = ""
lblLocation.Caption = ""
lblUnitPrice.Caption = ""
optIn.Value = False
optOut.Value = False
txtQuantity.Text = ""
End Sub

Public Sub setStatus(ByVal strMessage As String)
bar.Panels(bar.Panels.Count).Text = strMessage
End Sub

Private Sub cmbFilter_Click()
If (IsNull(cmbFilter.Text) = True) Or (cmbFilter.Text <> "") Then
    showListofProducts "SELECT * FROM Products WHERE CategoryID='" & cmbFilter.Text & "';"
    clearValues
End If
End Sub

Private Sub cmdDone_Click()
If lblDescription.Caption = "" Then 'Indicates no product selected
    ValidMsg "Please select a product first.", "No product selected."
    lvInventory.SetFocus
'check if options selected
ElseIf ((optIn.Value = False) And (optOut.Value = False)) Then
    ValidMsg "Please select an option of adding or deducting stock.", "Missing selection"
    optIn.SetFocus
ElseIf Len(txtQuantity.Text) = 0 Then 'check if quantity entered.
    ValidMsg "Please enter a quantity of the selected stock.", "Missing entry"
    txtQuantity.SetFocus
ElseIf ((Val(txtQuantity.Text) < 1) Or (Val(txtQuantity.Text) > 30000)) Then
    ValidMsg "Please enter a value for quantity between 0 and 30000.", "Invalid value"
    txtQuantity.SetFocus
ElseIf (optOut.Value = True) And (Val(txtQuantity.Text) > Val(lblQuantity.Text)) Then
    ValidMsg "Please enter a value that do not exceed the amount available in store.", "Invalid value"
    txtQuantity.SetFocus
Else
    'Begin saving
    Dim tempSQL As String, tmpMessage As String
    Dim oldQty As Integer, tmpQty As Integer
    Dim tRS As Recordset
    tmpQty = CInt(txtQuantity.Text)
    On Error GoTo ErrHandler
    BeginTrans
    tempSQL = "SELECT Quantity FROM Products WHERE ProductID='" & lblID.Caption & "';"
    Set tRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset, dbDenyWrite)
    If Not tRS.EOF Then
        tRS.Edit
        oldQty = CInt(tRS("Quantity"))
        If optIn.Value = True Then
            tRS("Quantity") = oldQty + tmpQty
            tempSQL = "INSERT INTO Internal_Transaction VALUES('" & Format$(Now(), "dd/mm/yyyy") & "','" & lblID.Caption & "',True," & tmpQty & ");"
        Else
            tRS("Quantity") = oldQty - tmpQty
            tempSQL = "INSERT INTO Internal_Transaction VALUES('" & Format$(Now(), "dd/mm/yyyy") & "','" & lblID.Caption & "',False," & tmpQty & ");"
        End If
        tRS.Update
        'updates the transaction log
        MySynonDatabase.Execute tempSQL
        CommitTrans
        'inform user through status bar
        setStatus "Product ID: " & lblID.Caption & " has been updated!"
        'Clear the screen
        clearValues
        Call cmbFilter_Click
        tRS.Close
        Set tRS = Nothing
    End If
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description & vbNewLine & "The changes have not been made."
End If
End Sub

Private Sub Form_Load()
formatListView
FillCombo cmbFilter, "SELECT Categories.CategoryID FROM Categories;", "CategoryID"
cmbFilter.ListIndex = 0
Call cmbFilter_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
lvInventory.width = Me.width - (480 + Frame1.width)
lvInventory.height = Me.ScaleHeight - (250 + bar.height)
Frame1.Left = Me.width - (250 + Frame1.width)
optIn.Left = Frame1.Left + 120
optOut.Left = optIn.Left
optIn.Top = Frame1.Top + 300
optOut.Top = optIn.Top + optIn.height + 105
cmdDone.Left = Frame1.Left + 1200
cmdDone.Top = Frame1.Top + 1320
Frame2.Left = Frame1.Left
Frame3.Left = Frame1.Left
bar.Panels(1).width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmProduct_Browse = Nothing
End Sub

Private Sub lvInventory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lvInventory '// change to the name of the list view
    Static iLast As Integer, iCur As Integer
    .Sorted = True
    iCur = ColumnHeader.Index - 1
    If iCur = iLast Then .SortOrder = IIf(.SortOrder = 1, 0, 1)
    .SortKey = iCur
    iLast = iCur
End With
End Sub

Private Sub lvInventory_DblClick()
With lvInventory
    If .ListItems.Count > 0 Then
        If .SelectedItem.Selected Then
            Load frmProduct_Details
            frmProduct_Details.getProductDetails lblID.Caption
            frmProduct_Details.Show vbModal
        End If
    End If
End With
End Sub

Private Sub lvInventory_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        getValues
    Else
        clearValues
    End If
End With
End Sub

Private Sub lvInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    If lvInventory.SelectedItem.Selected Then
        PopupMenu mnu_Delivery, vbPopupMenuLeftAlign
    End If
End If
End Sub

Private Sub mnu_Delivery_Click()
If lblID.Caption = "" Then
    mnu_Delivery_Add.Enabled = False
    mnu_Purchase_Add.Enabled = False
Else
    mnu_Delivery_Add.Enabled = True
    mnu_Purchase_Add.Enabled = True
    'Deal with delivery orders and purchase orders
    Dim i As Integer, j As Integer, frm As Form
    For i = mnu_ExistingDO.LBound To mnu_ExistingDO.UBound
        If i = 0 Then
            mnu_ExistingDO(i).Caption = "--{NONE}--"
        Else
            Unload mnu_ExistingDO(i)
        End If
    Next i
    For j = mnu_ExistingPO.LBound To mnu_ExistingPO.UBound
        If j = 0 Then
            mnu_ExistingPO(j).Caption = "--{NONE}--"
        Else
            Unload mnu_ExistingPO(j)
        End If
    Next j
    'Loads the number of menu equal to the current number of DO/PO forms
    i = 0
    j = 0
    For Each frm In Forms
        If Left$(frm.Tag, 2) = "DO" Then
            If i > 0 Then
                Load mnu_ExistingDO(i)
            End If
            mnu_ExistingDO(i).Caption = "New Delivery Order - " & frm.Tag
            mnu_ExistingDO(i).Tag = frm.Tag
            i = i + 1
        ElseIf Left$(frm.Tag, 2) = "PO" Then
            If j > 0 Then
                Load mnu_ExistingPO(j)
            End If
            mnu_ExistingPO(i).Caption = "New Purchase Order - " & frm.Tag
            mnu_ExistingPO(i).Tag = frm.Tag
            j = j + 1
        End If
    Next
    'Deal with consignments
    Dim invRS As Recordset
    For i = mnu_Insert.LBound To mnu_Insert.UBound
        If i <> 0 Then
            Unload mnu_Insert(i)
        End If
    Next i
    RSOpen invRS, "SELECT Contracts.ContractNo, Customers.Name FROM Customers INNER JOIN Contracts ON Customers.CustomerID=Contracts.CustomerID", dbOpenSnapshot
    i = 0
    While Not invRS.EOF
        mnu_Insert(i).Caption = invRS("Name")
        mnu_Insert(i).Tag = invRS("ContractNo")
        invRS.MoveNext
        If Not invRS.EOF Then
            i = i + 1
            Load mnu_Insert(i)
        End If
    Wend
    invRS.Close
    Set invRS = Nothing
End If

End Sub

Private Sub mnu_ExistingDO_Click(Index As Integer)
If mnu_ExistingDO(Index).Caption <> "--{NONE}--" Then
    frmSelected.add lvInventory.SelectedItem, mnu_ExistingDO(Index).Tag
End If
End Sub

Private Sub mnu_ExistingPO_Click(Index As Integer)
If mnu_ExistingPO(Index).Caption <> "--{NONE}--" Then
    frmSelectedForPO.add lvInventory.SelectedItem, mnu_ExistingPO(Index).Tag
End If
End Sub

Private Sub mnu_Insert_Click(Index As Integer)
If lblID.Caption <> "" Then
    If mnu_Insert.Count > 0 Then
        If MsgBox("Are you sure you want to add this product into " & mnu_Insert(Index).Caption & " consignment contract? " & vbCrLf & "The initial quantity will be 0 and it may be removed from the consignment anytime.", vbYesNo + vbQuestion, "Add to consignment") = vbYes Then
            Dim insertSQL As String, insertRS As Recordset
            With lvInventory.SelectedItem
                On Error Resume Next
                RSOpen insertRS, "SELECT ProductID FROM C_Details WHERE ProductID='" & .Text & "' AND ContractNo='" & mnu_Insert(Index).Tag & "';", dbOpenSnapshot
                insertRS.MoveFirst
                insertRS.MoveLast
                If insertRS.RecordCount = 0 Then
                    insertSQL = "INSERT INTO C_Details VALUES ('" & mnu_Insert(Index).Tag & "','" & .Text & "','',0,0,0,'')"
                    MySynonDatabase.Execute insertSQL
                    If Err.Number <> 0 Then
                        CriticalMsg "Unable to add product into consignment. Please try again.", "Error found"
                    Else
                        InfoMsg "Product has been successfully inserted into the consignment.", "Record inserted"
                    End If
                Else 'Already in there
                    ValidMsg "The product is already available in the consignment inventory.", "Record exist"
                End If
                insertRS.Close
                Set insertRS = Nothing
            End With
        End If
    End If
End If
End Sub

Private Sub mnu_New_Click()
Dim f As frmDelivery
Set f = New frmDelivery
Load f
f.Show , frmMain
End Sub

Private Sub mnu_New_PO_Click()
Dim p As frmPurchase
Set p = New frmPurchase
Load p
p.Show , frmMain

End Sub

Private Sub mnu_Options_Add_Click()
frmProduct_New.Show vbModal
End Sub

Private Sub mnu_Options_Close_Click()
Unload Me
End Sub

Private Sub mnu_Refresh_Click()
cmbFilter_Click
End Sub

Private Sub txtQuantity_GotFocus()
SelText txtQuantity
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtQuantity_LostFocus()
If txtQuantity.Text = "" Then
    txtQuantity.Text = "0"
End If
End Sub


