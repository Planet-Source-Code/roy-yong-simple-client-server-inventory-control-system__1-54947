VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoicing 
   Caption         =   "Invoicing"
   ClientHeight    =   6945
   ClientLeft      =   2040
   ClientTop       =   1455
   ClientWidth     =   9360
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
   ScaleHeight     =   6945
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   0
      ItemData        =   "frmInvoicing.frx":0000
      Left            =   1680
      List            =   "frmInvoicing.frx":0061
      Style           =   2  'Dropdown List
      TabIndex        =   14
      ToolTipText     =   "Day"
      Top             =   6480
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   1
      ItemData        =   "frmInvoicing.frx":00E1
      Left            =   2400
      List            =   "frmInvoicing.frx":0109
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Month"
      Top             =   6480
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   2
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Year"
      Top             =   6480
      Width           =   855
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
      TabIndex        =   10
      ToolTipText     =   "Click here to close this window."
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdInvoice 
      Caption         =   "&Invoice"
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
      TabIndex        =   9
      ToolTipText     =   "Click to invoice the customer account now."
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtDO 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtDet 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   6000
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvDet 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvDO 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
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
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmInvoicing.frx":013D
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "+"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lbldes 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   5640
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Additional Charges:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmInvoicing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sumDet As Single, sumTotal As Single
Private Sub getDO(ByVal strSQL As String)
Dim doRS As Recordset
With lvDO
    sumTotal = 0
    sumDet = 0
    lbldes.Caption = ""
    txtDO.Text = "0.00"
    txtDet.Text = "0.00"
    txtTotal.Text = "0.00"
    .ListItems.Clear
    lvDet.ListItems.Clear
    Debug.Print strSQL
    RSOpen doRS, strSQL, dbOpenSnapshot
    'On Error GoTo ErrHandler
    While Not doRS.EOF
        .ListItems.add , , doRS("DOnumber")
        .ListItems(.ListItems.Count).SubItems(1) = doRS("Name")
        .ListItems(.ListItems.Count).SubItems(2) = doRS("PONumber")
        .ListItems(.ListItems.Count).SubItems(3) = doRS("DelDate")
        .ListItems(.ListItems.Count).SubItems(4) = doRS("DelTime")
        .ListItems(.ListItems.Count).SubItems(5) = doRS("Status")
        .ListItems(.ListItems.Count).SubItems(6) = Format$(doRS("Charges"), "#,##0.00")
        .ListItems(.ListItems.Count).SubItems(7) = doRS("CustomerID")
        .ListItems(.ListItems.Count).Tag = doRS("Description")
        doRS.MoveNext
    Wend
    
    doRS.Close
    Set doRS = Nothing
End With
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "The query passed is invalid. Please try again.", "Error found"
    Exit Sub
End If
End Sub
Private Sub getDetails(ByVal strDOnumber As String)
If strDOnumber <> "" Then
    Dim gRS As Recordset, gSQL As String
    gSQL = "SELECT * FROM D_Details WHERE DOnumber='" & strDOnumber & "'"
    RSOpen gRS, gSQL, dbOpenSnapshot
    lvDet.ListItems.Clear
    While Not gRS.EOF
        With lvDet.ListItems
            .add , , gRS("ProductID")
            .Item(.Count).SubItems(1) = gRS("Description")
            .Item(.Count).SubItems(2) = gRS("CustRef")
            .Item(.Count).SubItems(3) = gRS("Quantity")
            .Item(.Count).SubItems(4) = gRS("UnitLabel")
            .Item(.Count).SubItems(5) = Format$(gRS("SalePrice"), "#,##0.00")
            'sumDet = sumDet + CSng(gRS("Quantity") * gRS("SalePrice"))
            .Item(.Count).Checked = True
        End With
        gRS.MoveNext
    Wend
    gRS.Close
    Set gRS = Nothing
End If
End Sub

Private Sub displayTotal()
Dim i As Integer
sumDet = 0
For i = 1 To lvDet.ListItems.Count
    If lvDet.ListItems(i).Checked = True Then
        sumDet = sumDet + CSng(lvDet.ListItems(i).SubItems(3) * lvDet.ListItems(i).SubItems(5))
    End If
Next i
txtDO.Text = Format$(lvDO.SelectedItem.SubItems(6), "#,##0.00")
lbldes.Caption = lvDO.SelectedItem.Tag
txtDet.Text = Format$(sumDet, "#,##0.00")
sumTotal = txtDO.Text + sumDet
txtTotal.Text = Format$(sumTotal, "#,##0.00")
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdInvoice_Click()
If lvDO.ListItems.Count > 0 Then
    If lvDO.SelectedItem.Selected Then
        If sumDet > 0 Then
            If MsgBox("Do you want to confirm to invoice this DO into the debtor's account?", vbYesNo + vbQuestion, "Invoice") = vbYes Then
                'Begin invoicing steps
                Screen.MousePointer = 11
                Dim tmpRS As Recordset
                Dim i As Integer
                'On Error GoTo ErrHandler
                
                RSOpen tmpRS, "SELECT * FROM cust_transactions", dbOpenDynaset
                tmpRS.AddNew
                tmpRS("date") = cmbDate(0).Text & "/" & cmbDate(1).Text & "/" & cmbDate(2).Text
                tmpRS("CustomerID") = lvDO.SelectedItem.SubItems(7)
                tmpRS("debit") = sumTotal
                tmpRS("DOnumber") = lvDO.SelectedItem.Text
                tmpRS("notes") = "Invoiced for DO - " & lvDO.SelectedItem.Text
                tmpRS.Update
                'BeginTrans
                'Update delivery order status
                RSOpen tmpRS, "SELECT Status FROM Delivery WHERE DOnumber='" & lvDO.SelectedItem.Text & "';", dbOpenDynaset
                tmpRS.Edit
                tmpRS("Status") = "INVOICED"
                tmpRS.Update
                For i = 1 To lvDet.ListItems.Count
                    'Check each item that has been invoiced.
                    RSOpen tmpRS, "SELECT isInvoiced FROM D_Details WHERE DOnumber='" & lvDO.SelectedItem.Text & "' AND ProductID='" & lvDet.ListItems(i).Text & "';", dbOpenDynaset
                    If lvDet.ListItems(i).Checked = True Then
                        tmpRS.Edit
                        tmpRS("isInvoiced") = True
                        tmpRS.Update
                    End If
                Next i
                'Free memory space
                tmpRS.Close
                Set tmpRS = Nothing
                insertLog "Customer ID: " & lvDO.SelectedItem.SubItems(7) & " has been invoiced with amount of $" & Format$(sumTotal, "#,##0.00")
                Screen.MousePointer = 0
                InfoMsg "Customer ID: " & lvDO.SelectedItem.SubItems(7) & vbCrLf & _
                "Customer account has been invoiced for: " & vbCrLf & _
                "Delivery Order No: " & lvDO.SelectedItem.Text & vbCrLf & _
                "Amount: " & Format$(sumTotal, "#,##0.00"), "Record updated"
                getDO "SELECT Delivery.DOnumber, Delivery.CustomerID, Customers.Name, Delivery.PONumber, Delivery.DelDate, Delivery.DelTime, Delivery.Status, Delivery.Charges, Delivery.Description FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE ((Delivery.Status) <> 'INVOICED');"
                frmMain.loadRecDeliveries
                Me.SetFocus
            End If
        Else
            ValidMsg "Please ensure at least an item is selected to be invoiced to the debtor's account.", "No item selected."
        End If
    Else
        ValidMsg "Please select a Delivery Order to be invoiced.", "Missing DO"
    End If
Else
    ValidMsg "No delivery order available to be invoiced.", "No DO"
End If

ErrHandler:
If Err.Number <> 0 Then
    'Rollback
    ErrorNotifier 1001, "An error has been encounted during the process of updating customer and DO records. No changes have been made."
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 5
    cmbDate(2).addItem Format$(Year(Now()) - 4 + i)
Next i
'set default today's date
cmbDate(0).Text = Format$(Day(Now()), "00")
cmbDate(1).Text = Format$(Month(Now()), "00")
cmbDate(2).Text = Format$(Year(Now()), "00")
lblNotes.Caption = "Carefully check the items that have been successfully delivered to the customers. " & _
"Uncheck those if not delivered. When you are done, click on 'Invoice' to credit the customer account."
'Format the list view properties
With lvDO.ColumnHeaders
    .Clear
    .add , , "DO No."
    .Item(1).width = 800
    .add , , "Customer"
    .Item(2).width = 2000
    .add , , "PO Number"
    .add , , "Delivery Date"
    .add , , "Delivery Time"
    .add , , "Status"
    .add , , "Charges"
    .add , , "Customer ID"
    .Item(8).width = 0
End With

With lvDet
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Product ID"
    .ColumnHeaders(1).width = 975
    .ColumnHeaders.add , , "Description"
    .ColumnHeaders.add , , "Cust Ref"
    .ColumnHeaders.add , , "Quantity"
    .ColumnHeaders(4).width = 900
    .ColumnHeaders.add , , "Unit Label"
    .ColumnHeaders.add , , "Unit Price"
End With
getDO "SELECT Delivery.DOnumber, Delivery.CustomerID, Customers.Name, Delivery.PONumber, Delivery.DelDate, Delivery.DelTime, Delivery.Status, Delivery.Charges, Delivery.Description FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE ((Delivery.Status) <> 'INVOICED');"
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
lvDO.width = Me.ScaleWidth - lvDO.Left * 2
lvDet.width = lvDO.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmInvoicing = Nothing
End Sub

Private Sub lvDet_ItemCheck(ByVal Item As MSComctlLib.ListItem)
displayTotal
End Sub

Private Sub lvDO_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected = True Then
        'Get the DO details
        Dim dRS As Recordset
        getDetails .Text
        displayTotal
    Else
        lvDet.ListItems.Clear
    End If
End With
End Sub
