VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduct_Details 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6960
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
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvHistory 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2566
      View            =   3
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
   Begin MSComctlLib.TabStrip tb 
      Height          =   2175
      Left            =   0
      TabIndex        =   23
      Top             =   3360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3836
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
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
      TabIndex        =   10
      ToolTipText     =   "Click here to close this window."
      Top             =   2880
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
      Left            =   5760
      TabIndex        =   9
      ToolTipText     =   "Click here to begin editing the product details."
      Top             =   2400
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
      Left            =   5760
      TabIndex        =   12
      ToolTipText     =   "Click here to cancel editing."
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox cmbBrand 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtReorder 
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtMinimum 
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtQuantity 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      ToolTipText     =   "Quantity cannot be changed here. Adjustments are made through Inventory window."
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtProductID 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "Administrator privilege required in order to edit the product ID."
      Top             =   120
      Width           =   1455
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
      Left            =   5760
      TabIndex        =   11
      ToolTipText     =   "Click here to save changes."
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Product ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Reorder Level:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Minimum Level:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Unit Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Brand:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Category ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmProduct_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FormMode(ByVal strMode As ModeStatus)
'Determine the status of the controls
If strMode = Editing Then
    If CurrentUser.prvlgAdmin = True Then
        txtProductID.Locked = False
        txtQuantity.Locked = False
    End If
    txtDescription.Enabled = True
    cmbCategory.Enabled = True
    cmbBrand.Enabled = True
    txtUnitPrice.Enabled = True
    txtMinimum.Enabled = True
    txtReorder.Enabled = True
    txtLocation.Enabled = True
    lvHistory.Enabled = False
    
    cmdCancel.Visible = True
    cmdSave.Visible = True
    cmdEdit.Visible = False
    cmdClose.Visible = False
Else
    txtProductID.Locked = True
    txtQuantity.Locked = True
    txtDescription.Enabled = False
    cmbCategory.Enabled = False
    cmbBrand.Enabled = False
    txtUnitPrice.Enabled = False
    txtMinimum.Enabled = False
    txtReorder.Enabled = False
    txtLocation.Enabled = False
    lvHistory.Enabled = True
    cmdCancel.Visible = False
    cmdSave.Visible = False
    cmdEdit.Visible = True
    cmdClose.Visible = True
    
End If
End Sub

Private Sub setTabs()
With tb
    .Tabs.Clear
    .Tabs.add , "Sales", "Sales"
    .Tabs.add , "Purchases", "Purchases"
    .Tabs.add , "Transactions", "Transactions"
    If CurrentUser.prvlgAdmin = True Then
        .Tabs.add , "PricesOnly", "Unit Price"
        .Tabs.add , "CostOnly", "Unit Cost"
    End If
End With
End Sub

Private Sub showRecords(ByRef strRecordset As Recordset)
Dim i As Integer
With lvHistory
    .ColumnHeaders.Clear
    .ListItems.Clear
    For i = 0 To strRecordset.Fields.Count - 1
        If strRecordset.Fields(i).Name = "qty" Then
            .ColumnHeaders.add , , "Quantity", 970
        ElseIf strRecordset.Fields(i).Name = "isIn" Then
            .ColumnHeaders.add , , "Type", 900
        ElseIf strRecordset.Fields(i).Name = "ProductID" Then
            .ColumnHeaders.add , , "ProductID", 0
        Else
            .ColumnHeaders.add , , strRecordset.Fields(i).Name
        End If
    Next i
    While Not strRecordset.EOF
        For i = 0 To strRecordset.Fields.Count - 1
            If i = 0 Then
                .ListItems.add , , IIf(IsNull(strRecordset.Fields(i).Value), "", strRecordset.Fields(i).Value)
            Else
                If strRecordset.Fields(i).Name = "SalePrice" Then
                    .ListItems(.ListItems.Count).SubItems(i) = Format$(IIf(IsNull(strRecordset.Fields(i).Value), "0.00", strRecordset.Fields(i).Value), "#,##0.00")
                ElseIf strRecordset.Fields(i).Name = "isIn" Then
                    If strRecordset("isIn") = True Then
                        .ListItems(.ListItems.Count).SubItems(i) = "In"
                    Else
                        .ListItems(.ListItems.Count).SubItems(i) = "Out"
                    End If
                Else
                    .ListItems(.ListItems.Count).SubItems(i) = IIf(IsNull(strRecordset.Fields(i).Value), "", strRecordset.Fields(i).Value)
                End If
            End If
        Next i
        strRecordset.MoveNext
    Wend
End With
End Sub

Public Sub getProductDetails(ByVal strProductID As String)
Dim getRS As Recordset, getSQL As String
Me.Caption = strProductID
getSQL = "SELECT * FROM Products WHERE Products.ProductID='" & strProductID & "';"
RSOpen getRS, getSQL, dbOpenSnapshot
If Not getRS.EOF Then
    txtProductID.Text = getRS("ProductID")
    cmbCategory.Text = getRS("CategoryID")
    cmbBrand.Text = getRS("Brand")
    txtDescription.Text = getRS("Description")
    txtQuantity.Text = getRS("Quantity")
    txtMinimum.Text = getRS("MinLevel")
    txtReorder.Text = IIf(IsNull(getRS("ReorderLevel")), "0", getRS("ReorderLevel"))
    txtUnitPrice.Text = Format$(getRS("UnitPrice"), "#,##0.00")
    txtLocation.Text = IIf(IsNull(getRS("Location")), "", getRS("Location"))
End If
getRS.Close
Set getRS = Nothing
End Sub

Private Sub Form_Load()
FillCombo cmbCategory, "SELECT Categories.CategoryID FROM Categories ORDER BY CategoryID ASC;", "CategoryID"
FillCombo cmbBrand, "SELECT DISTINCT Products.Brand FROM Products ORDER BY Brand ASC;", "Brand"
setTabs
FormMode Viewing
End Sub

Private Sub tb_Click()
Dim tempSQL As String
tempSQL = ""
With tb
    Select Case .SelectedItem.Key
        Case "Sales"
            tempSQL = "SELECT Delivery.Date, Delivery.DOnumber, Customers.Name, D_Details.Quantity, D_Details.UnitLabel, D_Details.SalePrice " & _
                    "FROM Customers INNER JOIN (Delivery INNER JOIN D_Details ON Delivery.DOnumber = D_Details.DOnumber) ON Customers.CustomerID = Delivery.CustomerID " & _
                    "WHERE (((D_Details.ProductID)='" & Me.Caption & "') AND ((D_Details.isInvoiced)=True));"
        Case "Purchases"
        
        Case "Transactions"
            tempSQL = "SELECT * FROM internal_transaction WHERE ProductID='" & Me.Caption & "';"
        Case "PricesOnly"
            tempSQL = "SELECT DISTINCT Delivery.Date, D_Details.SalePrice " & _
                        "FROM Delivery INNER JOIN D_Details ON Delivery.DOnumber = D_Details.DOnumber " & _
                        "WHERE (((D_Details.ProductID)='" & Me.Caption & "'));"
        Case "CostOnly"
            tempSQL = "SELECT DISTINCT Purchase.Date, P_Details.UnitPrice " & _
                        "FROM Purchase INNER JOIN P_Details ON Purchase.poNumber = P_Details.poNumber " & _
                        "WHERE (((P_Details.ProductID)='" & Me.Caption & "'));"
    End Select
    
    If tempSQL <> "" Then
        Dim reportRS As Recordset
        RSOpen reportRS, tempSQL, dbOpenSnapshot
        showRecords reportRS
        reportRS.Close
        Set reportRS = Nothing
    End If
End With
End Sub

Private Sub cmbBrand_LostFocus()
CapCon cmbBrand
End Sub

Private Sub cmbCategory_LostFocus()
CapCon cmbCategory
End Sub

Private Sub cmdCancel_Click()
txtProductID = txtProductID.Tag
txtDescription.Text = txtDescription.Tag
cmbBrand.Text = cmbBrand.Tag
cmbCategory.Text = cmbCategory.Tag
txtMinimum.Text = txtMinimum.Tag
txtReorder.Text = txtReorder.Tag
txtLocation.Text = txtLocation.Tag
txtUnitPrice.Text = txtUnitPrice.Tag

FormMode Viewing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
txtProductID.Tag = txtProductID
txtDescription.Tag = txtDescription.Text
cmbBrand.Tag = cmbBrand.Text
cmbCategory.Tag = cmbCategory.Text
txtMinimum.Tag = txtMinimum.Text
txtReorder.Tag = txtReorder.Text
txtLocation.Tag = txtLocation.Text
txtUnitPrice.Tag = txtUnitPrice.Text

FormMode Editing
End Sub

Private Sub cmdSave_Click()
If txtDescription.Text = "" Then
    ValidMsg "Please provide a description for this product.", "Missing description"
    txtDescription.SetFocus
ElseIf cmbCategory.Text = "" Then
    ValidMsg "Please select a category.", "Missing category"
    cmbCategory.SetFocus
ElseIf txtMinimum.Text = "" Then
    ValidMsg "Please enter a minimum level for this product.", "Missing minimum level"
    txtMinimum.SetFocus
ElseIf Val(txtMinimum.Text) > 30000 Then
    ValidMsg "Please enter a value between 0 to 30000 for the minimum level.", "Invalid value"
    txtMinimum.SetFocus
ElseIf Val(txtReorder.Text) > 30000 Then
    ValidMsg "Please enter a value between 0 to 30000 for the reorder level.", "Invalid value"
    txtReorder.SetFocus
Else
    Dim pRS As Recordset, pSQL As String
    pSQL = "SELECT * FROM Products WHERE ProductID='" & Me.Caption & "';"
    RSOpen pRS, pSQL, dbOpenDynaset
    If Not pRS.EOF Then
        pRS.Edit
        pRS("ProductID") = txtProductID.Text
        pRS("Description") = txtDescription.Text
        pRS("Brand") = cmbBrand.Text
        pRS("CategoryID") = cmbCategory.Text
        pRS("MinLevel") = txtMinimum.Text
        pRS("ReorderLevel") = txtReorder.Text
        pRS("Location") = txtLocation.Text
        pRS("UnitPrice") = txtUnitPrice.Text
        pRS.Update
        insertLog "Product ID: " & Me.Caption & " has been updated."
        InfoMsg "The product details have been successfully updated.", "Record saved"
    End If
    pRS.Close
    Set pRS = Nothing
    FormMode Viewing
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmProduct_Details = Nothing
End Sub

Private Sub txtDescription_GotFocus()
SelText txtDescription
End Sub

Private Sub txtDescription_LostFocus()
CapCon txtDescription
End Sub

Private Sub txtLocation_GotFocus()
SelText txtLocation
End Sub

Private Sub txtMinimum_GotFocus()
SelText txtMinimum
End Sub

Private Sub txtMinimum_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtMinimum_LostFocus()
If txtMinimum.Text = "" Then
    txtMinimum.Text = "0"
End If
End Sub

Private Sub txtProductID_GotFocus()
SelText txtProductID
End Sub

Private Sub txtQuantity_GotFocus()
SelText txtQuantity
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtReorder_GotFocus()
SelText txtReorder
End Sub

Private Sub txtReorder_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtReorder_LostFocus()
If txtReorder.Text = "" Then
    txtReorder.Text = "0"
End If
End Sub

Private Sub txtUnitPrice_GotFocus()
SelText txtUnitPrice
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtUnitPrice_LostFocus()
If txtUnitPrice.Text = "" Then
    txtUnitPrice.Text = "0"
End If
txtUnitPrice.Text = Format$(txtUnitPrice.Text, "#,##0.00")
End Sub
