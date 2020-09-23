VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Delivery Order"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDelivery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDelTime 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery.frx":058A
      Left            =   1920
      List            =   "frmDelivery.frx":059A
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2520
      Width           =   615
   End
   Begin VB.ComboBox cmbDelTime 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery.frx":05AE
      Left            =   1200
      List            =   "frmDelivery.frx":05FA
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2520
      Width           =   615
   End
   Begin VB.ComboBox cmbDelDate 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox cmbDelDate 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery.frx":065E
      Left            =   1920
      List            =   "frmDelivery.frx":0686
      Style           =   2  'Dropdown List
      TabIndex        =   30
      ToolTipText     =   "Month"
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox cmbDelDate 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery.frx":06BA
      Left            =   1200
      List            =   "frmDelivery.frx":071B
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2160
      Width           =   615
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
      Left            =   3960
      TabIndex        =   14
      ToolTipText     =   "Click here to remove all items in the cart."
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtDes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   11
      ToolTipText     =   "Description for charges."
      Top             =   5400
      Width           =   4815
   End
   Begin VB.TextBox txtSub 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtExtra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      MaxLength       =   12
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtGrand 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   5760
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
      Left            =   5280
      TabIndex        =   15
      ToolTipText     =   "Click here to remove the selected items."
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtREM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtAttn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtPO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox cmbEmployee 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
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
      Left            =   6600
      TabIndex        =   16
      ToolTipText     =   "Click here to save this Delivery Order."
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvCart 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3413
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
   Begin VB.ComboBox cmbEmployee 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox cmbCustomer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   2
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery.frx":079B
      Left            =   1920
      List            =   "frmDelivery.frx":07C3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Month"
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery.frx":07F7
      Left            =   1200
      List            =   "frmDelivery.frx":0858
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   7680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   34
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDelivery.frx":08D8
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label12 
      Caption         =   "Delivery Time:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Delivery Date:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Grand Total:"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Subtotal:"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Addtional Charges:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Remark:"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Attention:"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "PO Number:"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Delivered By:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Endorsed By:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Customer ID:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numItems As Integer
Dim max_Item As Integer
Dim currTotal As Single, addCharge As Single, subTotal As Single

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

Private Sub cmbEmployee_Click(Index As Integer)
If Not cmbEmployee(Index).Text = "" Then
    Dim tempRS As Recordset
    RSOpen tempRS, "SELECT EmployeeID FROM Employees WHERE Name='" & cmbEmployee(Index).Text & "'", dbOpenSnapshot
    If Not tempRS.EOF Then
        cmbEmployee(Index).Tag = tempRS("EmployeeID")
    End If
    tempRS.Close
    Set tempRS = Nothing
End If
End Sub

Private Sub cmdClear_Click()
If numItems > 0 Then
    If MsgBox("Are you sure you want to clear the cart? Every item in the cart will be removed.", vbYesNo + vbQuestion, "Clear cart") = vbYes Then
        lvCart.ListItems.Clear
        subTotal = 0
        numItems = 0
    End If
Else
    InfoMsg "No item in cart.", "Cart empty"
End If
End Sub

Private Sub cmdDone_Click()
If isDateValid(CByte(cmbDate(0).Text), CByte(cmbDate(1).Text), CInt(cmbDate(2).Text)) = False Then
    ValidMsg "Invalid date. Please try again", "Invalid value"
    cmbDate(0).SetFocus
ElseIf isDateValid(CByte(cmbDelDate(0).Text), CByte(cmbDelDate(1).Text), CInt(cmbDelDate(2).Text)) = False Then
    ValidMsg "Invalid delivery date. Please try again.", "Invalid value"
    cmbDelDate(0).SetFocus
ElseIf cmbCustomer.Text = "" Then
    ValidMsg "Please select a customer for this delivery order.", "Missing selection"
    cmbCustomer.SetFocus
ElseIf cmbEmployee(0).Text = "" Then
    ValidMsg "Please select an endorsement officer.", "Missing selection"
    cmbEmployee(0).SetFocus
ElseIf cmbEmployee(1).Text = "" Then
    ValidMsg "Please select a delivery officer.", "Missing selection"
    cmbEmployee(1).SetFocus
ElseIf ((CSng(txtExtra.Text) > 0) And (txtDes.Text = "")) Then
    ValidMsg "Please provide a description for the additional charges.", "Missing values"
    txtDes.SetFocus
ElseIf numItems = 0 Then
    ValidMsg "Please add at least an item into the cart.", "Cart empty"
Else
    Dim tempSQL As String, DONumber As String, DODate As String
    Dim delTime As String, delDate As String
    Dim smallRS As Recordset, i As Integer
    'obtain new delivery order
    'BeginTrans
    'On Error GoTo ErrHandler
    tempSQL = "SELECT * FROM Misc WHERE DataType='DLVR'"
    Set smallRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset, dbDenyRead + dbDenyWrite)
    DONumber = Format$(CLng(smallRS("DataValue")), "000000")
    'Update the new delivery order number
    'Assign values to variable
    DODate = cmbDate(0).Text & "/" & cmbDate(1).Text & "/" & cmbDate(2).Text
    delDate = cmbDelDate(0).Text & "/" & cmbDelDate(1).Text & "/" & cmbDelDate(2).Text
    delTime = cmbDelTime(0).Text & ":" & cmbDelTime(1).Text
    'begin insertion of data
    tempSQL = "INSERT INTO Delivery VALUES('" & DONumber & "','" & DODate & "'," & _
    "'" & cmbEmployee(0).Tag & "','" & cmbEmployee(1).Tag & "','" & txtpo.Text & "','" & cmbCustomer.Tag & "'," & _
    "'" & txtREM.Text & "','" & txtAttn.Text & "','" & delDate & "','" & delTime & "'," & _
    "'DELIVERING','" & txtExtra.Text & "','" & txtDes.Text & "')"
    MySynonDatabase.Execute tempSQL
    
    For i = 1 To lvCart.ListItems.Count 'Goes through list of item in cart and add to database
        tempSQL = "INSERT INTO D_Details VALUES ('" & DONumber & "','" & lvCart.ListItems(i).Text & "'," & _
        "'" & lvCart.ListItems(i).SubItems(1) & "','" & lvCart.ListItems(i).SubItems(2) & "','" & lvCart.ListItems(i).SubItems(3) & "'," & _
        "'" & lvCart.ListItems(i).SubItems(4) & "','" & lvCart.ListItems(i).SubItems(5) & "',FALSE)"
        MySynonDatabase.Execute tempSQL
    Next i
    
    smallRS.Edit
    smallRS("DataValue") = CLng(DONumber) + 1
    smallRS.Update
    smallRS.Close
    'Comment the codes above and uncomment the ones below to have table locking
    'Not sure if it is correctly done. If anyone knows how to do it better, please do modify it.
    'and send a copy of this module to me. - Thanks
    'tempSQL = "SELECT * FROM Delivery;"
    'Set smallRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset, dbdeny + dbDenyWrite)
    'smallRS.AddNew
    'smallRS("DOnumber") = DONumber
    'smallRS("Date") = DODate
    'smallRS("EmployeeID") = endoEmp(0)
    'smallRS("IssuedBy") = delEmp(0)
    'smallRS("PONumber") = txtPO.Text
    'smallRS("CustomerID") = cmbCustomer.Text
    'smallRS("Remark") = txtREM.Text
    'smallRS("Attn") = txtAttn.Text
    'smallRS("DelDate") = delDate
    'smallRS("DelTime") = delTime
    'smallRS("Status") = "DELIVERING"
    'smallRS("Charges") = txtExtra.Text
    'smallRS("Description") = txtDes.Text
    'smallRS.Update
    
   ' tempSQL = "SELECT * FROM D_Details;"
   ' Set smallRS = MySynonDatabase.OpenRecordset(tempSQL, dbOpenDynaset, dbDenyWrite)
    'For i = 1 To lvCart.ListItems.Count
    '    smallRS.AddNew
    '    smallRS("DOnumber") = DONumber
    '    smallRS("ProductID") = lvCart.ListItems(i).Text
    '    smallRS("Description") = lvCart.ListItems(i).SubItems(1)
    '    smallRS("CustRef") = lvCart.ListItems(i).SubItems(2)
    '    smallRS("Quantity") = lvCart.ListItems(i).SubItems(3)
    '    smallRS("UnitLabel") = lvCart.ListItems(i).SubItems(4)
    '    smallRS("SalePrice") = lvCart.ListItems(i).SubItems(5)
    '    smallRS.Update
    'Next i
    InfoMsg "The delivery order has been created and ready for print.", "Record saved"
    'CommitTrans
    Set smallRS = Nothing
    frmMain.loadRecDeliveries
    Unload Me
End If
ErrHandler:
If Err.Number <> 0 Then
    If Err.Number = 3197 Then
        'Locked record has been updated before hand. Possibly
        InfoMsg "One of the record has been modified. The record has not been saved. Please click 'OK' and try again.", "Record not saved"
        Exit Sub
    Else
        'Rollback
        ErrorNotifier Err.Number, Err.description
    End If
End If
End Sub

Private Sub cmdRemove_Click()
removeItem
End Sub

Private Sub Form_Load()
lblNotes.Caption = "Red labels indicate required information. Items are added from the inventory window simply by right-clicking on them and select 'Add to cart'."
NumDOForm = NumDOForm + 1
Me.Tag = "DO" & NumDOForm
Me.Caption = Me.Caption & " - " & NumDOForm
Dim i As Integer
'Add years
For i = 0 To 5
    cmbDate(2).addItem CStr(Year(Now()) + i)
    cmbDelDate(2).addItem CStr(Year(Now()) + i)
Next i
For i = 0 To 59
    cmbDelTime(1).addItem Format$(i, "00")
Next i

'Populate combo boxes
FillCombo cmbCustomer, "SELECT Name FROM Customers", "Name"
FillCombo cmbEmployee(0), "SELECT Name FROM Employees ORDER BY Name ASC", "Name"
FillCombo cmbEmployee(1), "SELECT Name FROM Employees ORDER BY Name ASC", "Name"

max_Item = CInt(getSettings("cartSize"))
newDelivery
End Sub

Private Sub newDelivery()
'Default date set today
cmbDate(0).Text = Format$(Day(Now()), "00")
cmbDate(1).Text = Format$(Month(Now()), "00")
cmbDate(2).Text = Year(Now())
cmbDelDate(0).Text = cmbDate(0).Text
cmbDelDate(1).Text = cmbDate(1).Text
cmbDelDate(2).Text = cmbDate(2).Text
cmbDelTime(0).Text = Format$(Now(), "hh")
cmbDelTime(1).Text = Right$(Format$(Now(), "hh:mm"), 2)
'Initialise variables
numItems = 0
txtSub.Text = "0.00"
txtExtra.Text = "0.00"
txtGrand.Text = "0.00"

'Clear cart
lvCart.ListItems.Clear
lvCart.ColumnHeaders.Clear

'Set column header
With lvCart.ColumnHeaders
    .add , , "Product ID", 960
    .add , , "Description", 2200
    .add , , "Cust Ref", 1200
    .add , , "Quantity", 900
    .add , , "Unit Label", 900
    .add , , "Unit Price", 900
End With
End Sub

Private Sub removeItem()
Dim i As Integer
With lvCart
If numItems > 0 Then
    If MsgBox("Are you sure you want to remove the selected item(s) from the cart?", vbQuestion + vbYesNo, "Remove item") = vbYes Then
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected Then
                subTotal = subTotal - (CInt(.ListItems(i).SubItems(3)) * CSng(.ListItems(i).SubItems(5)))
                .ListItems.Remove .SelectedItem.Index
                numItems = numItems - 1
            End If
        Next i
    End If
Else
    InfoMsg "No item in cart.", "Cart empty "
End If
End With
End Sub

Public Sub addItem(ByVal strID As String, ByVal strDes As String, ByVal strRef As String, ByVal strQty As Integer, ByVal strLabel As String, ByVal strPrice As Single)
If numItems > max_Item Then
    InfoMsg "Cart is full.", "Cart full"
Else
    With lvCart.ListItems
        .add , , strID
        .Item(.Count).SubItems(1) = strDes
        .Item(.Count).SubItems(2) = strRef
        .Item(.Count).SubItems(3) = strQty
        .Item(.Count).SubItems(4) = strLabel
        .Item(.Count).SubItems(5) = Format$(strPrice, "#,##0.00")
        subTotal = subTotal + (strPrice * strQty)
        displayTotal
        numItems = numItems + 1
    End With
End If
End Sub

Private Sub displayTotal()
txtSub.Text = Format$(subTotal, "#,##0.00")
txtGrand.Text = Format$(subTotal + CSng(txtExtra.Text), "#,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
'NumDOForm = NumDOForm - 1
Set frmDelivery = Nothing
End Sub

Private Sub txtExtra_Change()
displayTotal
End Sub

Private Sub txtExtra_GotFocus()
SelText txtExtra
End Sub

Private Sub txtExtra_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtExtra_LostFocus()
txtExtra.Text = Format$(txtExtra.Text, "#,##0.00")
End Sub

Private Sub txtGrand_GotFocus()
SelText txtGrand
End Sub

Private Sub txtGrand_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtGrand_LostFocus()
txtGrand.Text = Format$(txtGrand.Text, "#,##0.00")
End Sub

Private Sub txtSub_GotFocus()
SelText txtSub
End Sub

Private Sub txtSub_LostFocus()
txtSub.Text = Format$(txtSub.Text, "#,##0.00")
End Sub
