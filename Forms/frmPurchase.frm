VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   5265
   ClientLeft      =   -135
   ClientTop       =   435
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPurchase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton cash 
      Caption         =   "Cash"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Click here if purchasing item by cash"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton credit 
      Caption         =   "Credit"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Click here if purchasing items by credit"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox total 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "This is the total cost if available."
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Cart"
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
      TabIndex        =   11
      ToolTipText     =   "Click here to clear all items from the cart."
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove..."
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
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "Click here to remove a selected item."
      Top             =   4800
      Width           =   1215
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
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "Click here when you wish to proceed to save this Purchase Order"
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvCart 
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "This cart contains the items that would be purchased."
      Top             =   3120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
      Caption         =   "Purchase Method:"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7575
      Begin VB.Frame Frame2 
         Caption         =   "Creditor Details:"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3735
         Begin VB.TextBox ref 
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   5
            ToolTipText     =   "Enter any references here. Eg: Quotation number, etc."
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox supplier 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Select a supplier."
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "Reference:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Supplier:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox employee 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select an employee responsible for this purchase order."
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox notes 
         Height          =   975
         Left            =   4680
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         ToolTipText     =   "Enter any remark for this Purchase Order"
         Top             =   960
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker datepk 
         Height          =   315
         Left            =   4680
         TabIndex        =   6
         ToolTipText     =   "Select the issued date for this Purchase Order"
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19726336
         CurrentDate     =   38148
      End
      Begin VB.Label Label6 
         Caption         =   "Emp ID:"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Remark:"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPurchase.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNotes"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   14
      Top             =   120
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8000
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numItems As Integer
Dim max_Item As Integer
Dim purchaseSum As Single

Private Sub cash_Click()
Frame2.Enabled = False
supplier.ListIndex = -1
ref.Text = ""
End Sub

Private Sub cmdClear_Click()
If numItems > 0 Then
    If MsgBox("Are you sure you want to clear the cart? Every item in the cart will be removed.", vbYesNo + vbQuestion, "Clear cart") = vbYes Then
        lvCart.ListItems.Clear
        purchaseSum = 0
        numItems = 0
    End If
Else
    InfoMsg "No item in cart.", "Cart empty"
End If
End Sub

Private Sub cmdDone_Click()
If (credit.Value = True) And supplier.Text = "" Then
    ValidMsg "Please select a supplier.", "Missing supplier"
    supplier.SetFocus
ElseIf numItems = 0 Then
    ValidMsg "Please ensure at least 1 item is in the cart.", "No item in cart"
Else
    Dim poRS As Recordset, tmpRS As Recordset
    Dim tmpSQL As String
    Dim newPOID As Long, i As Integer
    On Error GoTo ErrHandler
    BeginTrans
    tmpSQL = "SELECT DataValue FROM Misc WHERE DataType='PO';"
    Set tmpRS = MySynonDatabase.OpenRecordset(tmpSQL, dbOpenDynaset, dbDenyRead + dbDenyWrite)
    newPOID = Format$(tmpRS("DataValue"), "000000")
    'Create new purchase order record
    tmpSQL = "SELECT * FROM Purchase;"
    Set poRS = MySynonDatabase.OpenRecordset(tmpSQL, dbOpenDynaset, dbDenyRead + dbDenyWrite)
    poRS.AddNew
    poRS("poNumber") = newPOID
    If credit.Value = True Then
        poRS("SupplierID") = supplier.Tag
        poRS("isCash") = False
    Else
        poRS("isCash") = True
    End If
    poRS("Date") = Format$(datepk.Day, "00") & "/" & Format$(datepk.Month, "00") & "/" & Format$(datepk.Year, "0000")
    poRS("EmployeeID") = employee.Tag
    poRS("Notes") = notes.Text
    poRS("Ref") = ref.Text
    poRS.Update
    'Update details table
    tmpSQL = "SELECT * FROM P_Details;"
    Set poRS = MySynonDatabase.OpenRecordset(tmpSQL, dbOpenDynaset, dbAppendOnly)
    With lvCart
        For i = 1 To numItems
            poRS.AddNew
            poRS("poNumber") = CStr(newPOID)
            poRS("ProductID") = .ListItems(i).Text
            poRS("CustRef") = .ListItems(i).SubItems(2)
            poRS("Quantity") = CInt(.ListItems(i).SubItems(3))
            poRS("UnitLabel") = .ListItems(i).SubItems(4)
            poRS("UnitPrice") = CSng(.ListItems(i).SubItems(5))
            poRS.Update
        Next i
    End With
    'Update the next key
    tmpRS.Edit
    tmpRS("DataValue") = newPOID + 1
    tmpRS.Update
    CommitTrans
    tmpRS.Close
    poRS.Close
    Set tmpRS = Nothing
    Set poRS = Nothing
    
    InfoMsg "PO Number: " & newPOID & vbCrLf & "The purchase order has been successfully saved and ready for printing.", "Record saved"
    Unload Me
End If
ErrHandler:
If Err.Number <> 0 Then
    Rollback
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub cmdRemove_Click()
removeItem
End Sub

Private Sub credit_Click()
Frame2.Enabled = True
End Sub

Private Sub employee_Click()
If employee.Text <> "" Then
    employee.Tag = getEmpID(employee.Text)
End If
End Sub

Private Sub Form_Load()
'Insert notes here
lblNotes.Caption = "Red labels indicate required information. " & vbCrLf & "Items are added from the inventory window simply by right-clicking on them and selecting the corresponding purchase order."
'Set the properties of list view
With lvCart
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Product ID", 980
    .ColumnHeaders.add , , "Description", 2000
    .ColumnHeaders.add , , "Supp Ref"
    .ColumnHeaders.add , , "Quantity"
    .ColumnHeaders.add , , "Unit Label"
    .ColumnHeaders.add , , "Unit Price"
End With
FillCombo employee, "SELECT Name FROM Employees;", "Name"
FillCombo supplier, "SELECT Name FROM Suppliers;", "Name"
'Initialise variables
NumPOForm = NumPOForm + 1
Me.Caption = Me.Caption & " - " & NumPOForm
Me.Tag = "PO" & NumPOForm
max_Item = CInt(getSettings("cartSize"))
newPO
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
lblNotes.width = Me.ScaleWidth - (lblNotes.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'NumPOForm = NumPOForm - 1
Set frmPurchase = Nothing
End Sub

Private Sub newPO()
'Reset the date
datepk.Value = Now()
cash_Click
'Initialise values
ref.Text = ""
total.Text = "0.00"
numItems = 0
purchaseSum = 0
lvCart.ListItems.Clear
notes.Text = ""

supplier.ListIndex = -1
employee.ListIndex = -1
End Sub

Private Sub removeItem()
Dim i As Integer
With lvCart
If numItems > 0 Then
    If MsgBox("Are you sure you want to remove the selected item(s) from the cart?", vbQuestion + vbYesNo, "Remove item") = vbYes Then
        For i = 1 To numItems
            If .ListItems(i).Selected Then
                adjustSum ((CSng(.ListItems(i).SubItems(3)) * CSng(.ListItems(i).SubItems(5))) * -1)
                .ListItems.Remove .SelectedItem.Index
                numItems = numItems - 1
                displayTotal
            End If
        Next i
    End If
Else
    InfoMsg "No item in cart.", "Cart empty "
End If
End With
End Sub
Private Function getEmpID(ByVal strName As String) As String
Dim tmpRS As Recordset
RSOpen tmpRS, "SELECT EmployeeID FROM Employees WHERE Name='" & strName & "';", dbOpenSnapshot
If Not tmpRS.EOF Then
    getEmpID = tmpRS("EmployeeID")
Else
    getEmpID = ""
End If
tmpRS.Close
Set tmpRS = Nothing
End Function
Private Function getSuppID(ByVal strName As String) As String
Dim tmpRS As Recordset
RSOpen tmpRS, "SELECT SupplierID FROM Suppliers WHERE Name='" & strName & "';", dbOpenSnapshot
If Not tmpRS.EOF Then
    getSuppID = tmpRS("SupplierID")
Else
    getSuppID = ""
End If
tmpRS.Close
Set tmpRS = Nothing
End Function

Public Sub addItem(ByVal strID As String, ByVal strDes As String, ByVal strRef As String, ByVal qty As Integer, ByVal strLabel As String, ByVal strPrice As String)
If numItems > max_Item Then
    InfoMsg "Cart is full.", "Cart full"
Else
    With lvCart.ListItems
        .add , , strID
        .Item(.Count).SubItems(1) = strDes
        .Item(.Count).SubItems(2) = strRef
        .Item(.Count).SubItems(3) = qty
        .Item(.Count).SubItems(4) = strLabel
        .Item(.Count).SubItems(5) = Format$(strPrice, "#,##0.00")
        adjustSum CSng(strPrice * qty)
        displayTotal
        numItems = numItems + 1
    End With
End If

End Sub
Private Sub supplier_Click()
If supplier.Text <> "" Then
    supplier.Tag = getSuppID(supplier.Text)
End If
End Sub

Private Sub adjustSum(sngAmount As Single)
purchaseSum = purchaseSum + sngAmount
End Sub

Private Sub displayTotal()
total.Text = Format$(purchaseSum, "#,##0.00")
End Sub
