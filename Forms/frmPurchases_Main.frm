VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurchases_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Orders Management"
   ClientHeight    =   5280
   ClientLeft      =   -135
   ClientTop       =   435
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPurchases_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cash 
      Caption         =   "Credit Purchase"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Check this if purchase is done in credit terms"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   4080
      TabIndex        =   8
      Top             =   3000
      Width           =   3735
      Begin VB.TextBox ref 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Reference from Supplier."
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox supplier 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Supplier involved in this credit Purchase Order"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Reference:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Supplier ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
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
      Left            =   8040
      TabIndex        =   11
      ToolTipText     =   "Click here to create a new Purchase Order."
      Top             =   3360
      Width           =   1335
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
      Left            =   8040
      TabIndex        =   14
      ToolTipText     =   "Click here to close this window"
      Top             =   4800
      Width           =   1335
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
      Left            =   8040
      TabIndex        =   12
      ToolTipText     =   "Click here to edit the current selected Purchase Order"
      Top             =   3840
      Width           =   1335
   End
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
      Left            =   8040
      TabIndex        =   13
      ToolTipText     =   "Click here to delete the selected Purchase Order if any."
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   2
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Year."
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   1
      ItemData        =   "frmPurchases_Main.frx":08CA
      Left            =   1920
      List            =   "frmPurchases_Main.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Month."
      Top             =   3360
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   0
      ItemData        =   "frmPurchases_Main.frx":0926
      Left            =   1200
      List            =   "frmPurchases_Main.frx":0987
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Day."
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox notes 
      Height          =   855
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      ToolTipText     =   "Remarks made during the creation of the Purchase Order."
      Top             =   4320
      Width           =   2535
   End
   Begin VB.ComboBox employee 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Employee responsible for this Purchase Order."
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox po 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Purchase Order Number."
      Top             =   3000
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvPO 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "List of Purchase Orders."
      Top             =   960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3413
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
      Left            =   8040
      TabIndex        =   15
      ToolTipText     =   "Click here to save the changes."
      Top             =   4320
      Width           =   1335
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
      Left            =   8040
      TabIndex        =   16
      ToolTipText     =   "Click here to cancel editing."
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8760
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblPO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9120
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblSupp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Remark:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Issued By:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "PO Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPurchases_Main.frx":0A07
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNotes"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   17
      Top             =   120
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9435
   End
End
Attribute VB_Name = "frmPurchases_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cash_Click()
If cash.Value = vbChecked Then
    Frame1.Enabled = True
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
Dim k  As Integer
po.Text = lblPO.Caption
For k = 0 To 2
    cmbDate(k).Text = cmbDate(k).Tag
Next k
If cash.Tag = "No" Then
    cash.Value = vbChecked
Else
    cash.Value = vbUnchecked
End If
supplier.Text = lblSupp.Caption
employee.Text = lblEmp.Caption
ref.Text = ref.Tag
notes.Text = notes.Tag
setFormMode Viewing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If lvPO.ListItems.Count > 0 Then
    If po.Text <> "" Then
        If MsgBox("Are you sure you want to delete this purchase order?" & vbCrLf & "PO Number: " & lblPO.Caption & vbCrLf & "All its details would be deleted as well.", vbYesNo + vbQuestion, "Delete record") = vbYes Then
            'Proceed to delete the selected purchase order
            Dim tmpSQL As String
            tmpSQL = "DELETE * FROM Purchase WHERE poNumber='" & lblPO.Caption & "';"
            MySynonDatabase.Execute tmpSQL
            tmpSQL = "PO Number: " & lblPO.Caption & " has been deleted."
            insertLog tmpSQL
            InfoMsg tmpSQL, "Record deleted"
            clearFields
        End If
    Else
        InfoMsg "Please select a purchase order to be deleted.", "Missing selection"
    End If
Else
    InfoMsg "There are no purchase orders available to be deleted.", "No purchase orders available"
End If
End Sub

Private Sub cmdEdit_Click()
If lvPO.ListItems.Count > 0 Then
    If po.Text <> "" Then
        Dim k As Integer
        If cash.Value = vbUnchecked Then
            cash.Tag = "Yes"
        Else
            cash.Tag = "No"
        End If

        lblSupp.Caption = supplier.Text
        For k = 0 To 2
            cmbDate(k).Tag = cmbDate(k).Text
        Next k
        lblEmp.Caption = employee.Text
        notes.Tag = notes.Text
        ref.Tag = ref.Text
        setFormMode Editing
    Else
        InfoMsg "Please select a purchase order first.", "No purchase order selected"
    End If
Else
    InfoMsg "There are no purchase orders available.", "No purchase orders"
End If
End Sub

Private Sub cmdNew_Click()
Dim p As frmPurchase
Set p = New frmPurchase
Load p
p.Show , frmMain
End Sub

Private Sub cmdSave_Click()
If po.Text = "" Then
    ValidMsg "Please enter a purchase order number.", "Missing PO"
ElseIf isDateValid(CByte(cmbDate(0).Text), CByte(cmbDate(1).Text), CInt(cmbDate(2).Text)) = False Then
    ValidMsg "Please select a valid date.", "Invalid date"
    cmbDate(0).SetFocus
ElseIf employee.Text = "" Then
    ValidMsg "Please select an employee as the issuer of this purchase order.", "Missing employee"
    employee.SetFocus
ElseIf ((cash.Value = vbChecked) And (supplier.Text = "")) Then
    ValidMsg "Please select a supplier.", "Missing supplier"
    supplier.SetFocus
ElseIf ((cash.Value = vbChecked) And (ref.Text = "")) Then
    ValidMsg "Please enter a reference number or equivalent.", "Missing reference"
    ref.SetFocus
Else
    Dim saveRS As Recordset
    RSOpen saveRS, "SELECT  * FROM Purchase WHERE poNumber='" & lblPO.Caption & "';", dbOpenSnapshot
    If Not saveRS.EOF Then
        saveRS.Edit
        saveRS("poNumber") = po.Text
        saveRS("date") = cmbDate(0).Text & "/" & cmbDate(1).Text & "/" & cmbDate(2).Text
        saveRS("EmployeeID") = employee.Tag
        If cash.Value = vbChecked Then
            saveRS("isCash") = False
        Else
            saveRS("isCash") = True
        End If
        saveRS("Notes") = notes.Text
        saveRS("Ref") = ref.Text
        saveRS("supplier") = supplier.Tag
        saveRS.Update
        saveRS.Close
        Set saveRS = Nothing
        
        InfoMsg "PO Number: " & lblPO.Caption & " has been successfully updated.", "Record saved"
        getPO
    End If
End If
End Sub

Private Sub employee_Click()
If employee.Text <> "" Then
    employee.Tag = getEmpID(employee.Text)
End If
End Sub

Private Sub Form_Load()
DisableClose frmPurchases_Main, True
'Insert notes here
lblNotes.Caption = "Welcome to the purchase order management console. Please be careful in changing the details of these orders. " & vbCrLf & _
"Changes upon these documents may not reflect the truth in reality thus may cause undesirable outcomes and fatal errors. Ensure that you are " & _
"fully aware of what you are doing."

With lvPO.ColumnHeaders
    .Clear
    .add , , "PO No.", 880
    .add , , "Payment", 900
    .add , , "Name", 3000
    .add , , "Date"
    .add , , "Employee ID", 995
    .add , , "Reference"
    .add , , "Remark", 1500
End With
Dim j As Integer
For j = 0 To 5
    cmbDate(2).addItem Format$(Year(Now()) - 1 + j, "0000")
Next j
FillCombo employee, "SELECT Name FROM Employees;", "Name"
FillCombo supplier, "SELECT Name FROM Suppliers;", "Name"
getPO
setFormMode Viewing
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
lblNotes.width = Me.ScaleWidth - (lblNotes.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPurchases_Main = Nothing
End Sub

Private Sub lvPO_DblClick()
If lvPO.ListItems.Count > 0 Then
    If lvPO.SelectedItem.Selected Then
        Load frmPurchase_Details
        frmPurchase_Details.Tag = lvPO.SelectedItem.Text
        frmPurchase_Details.getDetails lvPO.SelectedItem.Text
        frmPurchase_Details.Show vbModal
    End If
End If
End Sub

Private Sub lvPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvPO.ListItems.Count > 0 Then
    If Item.Selected Then
        On Error Resume Next
        po.Text = Item.Text
        lblPO.Caption = Item.Text
        If Item.SubItems(1) = "Cash" Then
            cash.Value = vbUnchecked
            supplier.ListIndex = -1
            ref.Text = ""
        Else
            cash.Value = vbChecked
            supplier.Text = Item.SubItems(2)
            ref.Text = Item.SubItems(5)
        End If
        employee.Text = Item.SubItems(4)
        notes.Text = Item.SubItems(6)
        cmbDate(0).Text = Left$(Item.SubItems(3), 2)
        cmbDate(1).Text = Right$(Left$(Item.SubItems(3), 5), 2)
        cmbDate(2).Text = Right$(Item.SubItems(3), 4)
    End If
End If
End Sub

Private Sub notes_GotFocus()
SelText notes
End Sub

Private Sub po_GotFocus()
SelText po
End Sub

Private Sub po_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub po_LostFocus()
If po.Text <> "" Then
    po.Text = Format$(po.Text, "000000")
End If
End Sub

Private Sub ref_GotFocus()
SelText ref
End Sub

Private Sub supplier_Click()
If supplier.Text <> "" Then
    supplier.Tag = getSuppID(supplier.Text)
End If
End Sub

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

Private Sub setFormMode(ByVal strModeStatus As ModeStatus)
Dim i As Integer
If strModeStatus = Editing Then
    lvPO.Enabled = False
    po.Enabled = True
    For i = 0 To 2
        cmbDate(i).Enabled = True
    Next i
    cash.Enabled = True
    supplier.Enabled = True
    employee.Enabled = True
    ref.Enabled = True
    notes.Enabled = True
    cmdClose.Visible = False
    cmdNew.Visible = False
    cmdEdit.Visible = False
    cmdDelete.Visible = False
Else
    cash.Enabled = False
    lvPO.Enabled = True
    po.Enabled = False
    For i = 0 To 2
        cmbDate(i).Enabled = False
    Next i
    supplier.Enabled = False
    employee.Enabled = False
    notes.Enabled = False
    ref.Enabled = False
    cmdClose.Visible = True
    cmdNew.Visible = True
    cmdEdit.Visible = True
    cmdDelete.Visible = True
End If
cash_Click
End Sub

Private Sub getPO()
With lvPO.ListItems
    .Clear
    Dim getPORS As Recordset
    RSOpen getPORS, "SELECT Purchase.*, Suppliers.Name, Employees.Name As EmpName " & _
                    "FROM Suppliers INNER JOIN (Employees INNER JOIN Purchase ON Employees.EmployeeID = Purchase.EmployeeID) ON Suppliers.SupplierID = Purchase.SupplierID " & _
                    "ORDER BY Purchase.Date DESC;", dbOpenSnapshot
    While Not getPORS.EOF
        .add , , getPORS("poNumber")
        If getPORS("isCash") = True Then
            .Item(.Count).SubItems(1) = "Cash"
        Else
            .Item(.Count).SubItems(1) = "Credit"
        End If
        .Item(.Count).SubItems(2) = getPORS("Name")
        .Item(.Count).SubItems(3) = getPORS("Date")
        .Item(.Count).SubItems(4) = IIf(IsNull(getPORS("EmpName")), "", getPORS("EmpName"))
        .Item(.Count).SubItems(5) = IIf(IsNull(getPORS("Ref")), "", getPORS("Ref"))
        .Item(.Count).SubItems(6) = IIf(IsNull(getPORS("Notes")), "", getPORS("Notes"))
        getPORS.MoveNext
    Wend
    getPORS.Close
    Set getPORS = Nothing
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


Private Sub clearFields()
po.Text = ""
lblPO.Caption = ""
cash.Value = vbUnchecked
supplier.ListIndex = -1
ref.Text = ""
employee.ListIndex = -1
notes.Text = ""
cmbDate(0).ListIndex = -1
cmbDate(1).ListIndex = -1
cmbDate(2).ListIndex = -1
End Sub
