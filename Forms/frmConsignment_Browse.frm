VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsignment_Browse 
   Caption         =   "Consignment Inventory -"
   ClientHeight    =   8355
   ClientLeft      =   240
   ClientTop       =   1095
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
   Icon            =   "frmConsignment_Browse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10545
   Begin VB.Frame Frame1 
      Caption         =   "Adjustments:"
      Height          =   2295
      Left            =   8280
      TabIndex        =   4
      Top             =   5760
      Width           =   2175
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Index           =   1
         ItemData        =   "frmConsignment_Browse.frx":08CA
         Left            =   720
         List            =   "frmConsignment_Browse.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Month"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Index           =   0
         ItemData        =   "frmConsignment_Browse.frx":0926
         Left            =   120
         List            =   "frmConsignment_Browse.frx":0987
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Click here to change the date for the adjustment."
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optIn 
         Alignment       =   1  'Right Justify
         Caption         =   "In / +"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Select this if adding stock to the selected product."
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optOut 
         Alignment       =   1  'Right Justify
         Caption         =   "Out / -"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Select this if deducting stock of the selected product."
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   10
         ToolTipText     =   "Enter the amount of stock for this transaction of the selected product."
         Top             =   1320
         Width           =   615
      End
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
         Left            =   1200
         TabIndex        =   11
         ToolTipText     =   "Click here to begin adjustment."
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Product Info:"
      Height          =   3615
      Left            =   8280
      TabIndex        =   15
      Top             =   2040
      Width           =   2175
      Begin VB.TextBox lblMin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox lblReorder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox lblQuantity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblDescription 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblBrand 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblCategoryID 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1815
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
         TabIndex        =   24
         Top             =   1680
         Width           =   855
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
         TabIndex        =   23
         Top             =   1680
         Width           =   855
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
         TabIndex        =   22
         Top             =   2280
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
         TabIndex        =   21
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblLocation 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblUnitPrice 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtering:"
      Height          =   1335
      Left            =   8280
      TabIndex        =   1
      Top             =   720
      Width           =   2175
      Begin VB.ComboBox cmbFilter 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmConsignment_Browse.frx":0A07
         Left            =   120
         List            =   "frmConsignment_Browse.frx":0A09
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select the category."
         Top             =   480
         Width           =   1935
      End
      Begin MSComctlLib.ProgressBar pbBar 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label6 
         Caption         =   "By Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar bar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   8100
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvInventory 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8705
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
   Begin VB.Label lblEnd 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   32
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   31
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10575
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Refresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_Bar_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmConsignment_Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldQuery As String
Public Sub getConsignmentDetails(ByVal strContractNo As String)
Dim detailRS As Recordset
Me.Tag = strContractNo
RSOpen detailRS, "SELECT * FROM Contracts WHERE ContractNo='" & strContractNo & "';", dbOpenSnapshot
If Not detailRS.EOF Then
    lblNo.Caption = "Contract No: " & strContractNo
    lblStart.Caption = "Start Date: " & detailRS("StartDate")
    lblEnd.Caption = "End Date: " & IIf(IsNull(detailRS("ExpireDate")), "", detailRS("ExpireDate"))
    detailRS.Close
    Set detailRS = Nothing
    getListOfStocks "SELECT C_Details.*, Products.Description, Products.CategoryID, Products.Brand, Products.UnitPrice FROM Products INNER JOIN C_Details ON Products.ProductID=C_Details.ProductID WHERE ContractNo='" & Me.Tag & "' AND CategoryID='" & cmbFilter.Text & "';"
End If
End Sub

Private Sub getChosen()
'Obtain values from list view for further references
With lvInventory
    lblID.Caption = .SelectedItem.Text
    lblDescription.Caption = .SelectedItem.SubItems(1)
    lblBrand.Caption = .SelectedItem.SubItems(3)
    lblCategoryID.Caption = .SelectedItem.SubItems(4)
    lblQuantity.Text = .SelectedItem.SubItems(5)
    lblMin.Text = .SelectedItem.SubItems(6)
    lblReorder.Text = .SelectedItem.SubItems(7)
    lblLocation.Caption = .SelectedItem.SubItems(8)
    If getSettings("allowPrice") = "TRUE" Then
        lblUnitPrice.Caption = .SelectedItem.SubItems(9)
    End If
End With
End Sub

Private Sub clearChosen()
lblID.Caption = ""
lblDescription.Caption = ""
lblBrand.Caption = ""
lblCategoryID.Caption = ""
lblQuantity.Text = ""
lblMin.Text = ""
lblReorder.Text = ""
lblLocation.Caption = ""
If getSettings("allowPrice") = "TRUE" Then
    lblUnitPrice.Caption = ""
End If
End Sub

Private Sub getListOfStocks(ByVal strCustom As String)
Dim listRS As Recordset
lvInventory.ListItems.Clear
oldQuery = strCustom
RSOpen listRS, strCustom, dbOpenSnapshot
lblPercent.Caption = "0%"
pbBar.Value = 0
While Not listRS.EOF
    With lvInventory.ListItems
        pbBar.Value = listRS.PercentPosition
        lblPercent.Caption = listRS.PercentPosition & "%"
        .add , , listRS("ProductID")
        .Item(.Count).SubItems(1) = listRS("CustRef")
        .Item(.Count).SubItems(2) = listRS("Description")
        .Item(.Count).SubItems(3) = listRS("Brand")
        .Item(.Count).SubItems(4) = listRS("CategoryID")
        .Item(.Count).SubItems(5) = listRS("Quantity")
        .Item(.Count).SubItems(6) = listRS("MinLevel")
        .Item(.Count).SubItems(7) = listRS("ReorderLevel")
        .Item(.Count).SubItems(8) = IIf(IsNull(listRS("Location")), "", listRS("Location"))
        If getSettings("allowPrice") = "TRUE" Then
            .Item(.Count).SubItems(9) = Format$(listRS("UnitPrice"), "#,##0.00")
        End If
    End With
    listRS.MoveNext
Wend
pbBar.Value = 0
lblPercent.Caption = ""
listRS.Close
Set listRS = Nothing

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub setListFormat()
Dim i As Integer
With lvInventory
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Product ID"
    .ColumnHeaders.add , , "Customer Ref"
    .ColumnHeaders.add , , "Description"
    .ColumnHeaders.add , , "Brand"
    .ColumnHeaders.add , , "Category ID"
    .ColumnHeaders.add , , "Qty On Hand"
    .ColumnHeaders.add , , "Minimum Level"
    .ColumnHeaders.add , , "Reorder Level"
    .ColumnHeaders.add , , "Location"
    If getSettings("allowPrice") = "TRUE" Then
        .ColumnHeaders.add , , "Unit Price"
    End If
    For i = 1 To .ColumnHeaders.Count
        Select Case i
        Case 5, 6
            .ColumnHeaders(i).width = 900
        Case 7
            .ColumnHeaders(i).width = 1200
        Case 2
            .ColumnHeaders(i).width = 2000
        End Select
    Next i

End With
End Sub

Private Sub newTransaction()
'Specify by default today's date
cmbDate(0).Text = Format$(Day(Now()), "00")
cmbDate(1).Text = Format$(Month(Now()), "00")
cmbDate(2).Text = Year(Now())

txtQuantity.Text = ""
optIn.Value = False
optOut.Value = False

End Sub

Private Sub cmbFilter_Click()
If cmbFilter.Text <> "" Then
    getListOfStocks "SELECT C_Details.*, Products.Description, Products.CategoryID, Products.Brand, Products.UnitPrice FROM Products INNER JOIN C_Details ON Products.ProductID=C_Details.ProductID WHERE ContractNo='" & Me.Tag & "' AND CategoryID='" & cmbFilter.Text & "';"
End If
End Sub

Private Sub cmdDone_Click()
If lblID.Caption = "" Then
    Err.Clear
    ValidMsg "Please select a product first.", "Missing selection"
    lvInventory.SetFocus
ElseIf txtQuantity.Text = "" Then
    Err.Clear
    ValidMsg "Please enter a quantity.", "Missing values"
    txtQuantity.SetFocus
ElseIf ((Val(txtQuantity.Text) < 0) Or (Val(txtQuantity.Text) > Val(lblQuantity.Text))) Then
    Err.Clear
    ValidMsg "Please enter a quantity between 0 and " & lblQuantity.Text & ".", "Invalid value"
    txtQuantity.SetFocus
ElseIf ((optIn.Value = False) And (optOut.Value = False)) Then
    Err.Clear
    ValidMsg "Please select the type of transaction.", "Missing choice"
    optIn.SetFocus
ElseIf isDateValid(CByte(cmbDate(0).Text), CByte(cmbDate(1).Text), CInt(cmbDate(2).Text)) = False Then
    Err.Clear
    ValidMsg "The selected date is invalid. Please try again.", "Invalid date"
    cmbDate(0).SetFocus
Else
    Dim consignRS As Recordset, oldQty As Integer, currQty As Integer, tempSQL As String
    currQty = CInt(txtQuantity.Text)
    On Error GoTo ErrHandler
    BeginTrans
    Set consignRS = MySynonDatabase.OpenRecordset("SELECT * FROM C_Details WHERE ProductID='" & lblID.Caption & "'", dbOpenDynaset, dbDenyWrite + dbDenyRead)
    If Not consignRS.EOF Then
        oldQty = consignRS("Quantity")
        consignRS.Edit
        If optOut.Value = True Then
            consignRS("Quantity") = oldQty - currQty
        Else
            consignRS("Quantity") = oldQty + currQty
        End If
        consignRS.Update
        'consignRS.LockEdits = True
        consignRS.Close
        Set consignRS = Nothing
        If optOut.Value = False Then
            tempSQL = "INSERT INTO External_Transaction VALUES ('" & Format$(Now(), "dd/mm/yyyy") & "','" & Me.Tag & "','" & lblID.Caption & "', " & currQty & " ,False)"
        Else
            tempSQL = "INSERT INTO External_Transaction VALUES ('" & Format$(Now(), "dd/mm/yyyy") & "','" & Me.Tag & "','" & lblID.Caption & "'," & currQty & ",True)"
        End If
        MySynonDatabase.Execute tempSQL
        Set consignRS = Nothing
        CommitTrans
        InfoMsg "Transaction has been successfully recorded.", "Record saved"
        newTransaction
        getListOfStocks oldQuery
    End If
End If

ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 3
    cmbDate(2).addItem Format$(Year(Now()) + i)
Next i
FillCombo cmbFilter, "SELECT CategoryID FROM Categories", "CategoryID"
cmbFilter.ListIndex = 0
setListFormat
newTransaction
'bar.width = Me.width
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
Frame1.Left = Me.width - (Frame1.width + 180)
Frame2.Left = Frame1.Left
Frame3.Left = Frame1.Left
lvInventory.width = Me.width - (Frame1.width + lvInventory.Left + 250)
lvInventory.height = Me.height - (lvInventory.Top + bar.height + Shape1.height + 125)
bar.Panels(1).width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmConsignment_Browse = Nothing
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

Private Sub lvInventory_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Selected Then
    getChosen
Else
    clearChosen
End If
End Sub

Private Sub mnu_Close_Click()
Unload Me
End Sub

Private Sub txtQuantity_GotFocus()
SelText txtQuantity
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub
