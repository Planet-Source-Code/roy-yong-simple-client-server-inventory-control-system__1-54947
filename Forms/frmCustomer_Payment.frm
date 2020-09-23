VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomer_Payment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Payment"
   ClientHeight    =   3705
   ClientLeft      =   -135
   ClientTop       =   435
   ClientWidth     =   7875
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
   ScaleHeight     =   3705
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker datepk 
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   19726339
      CurrentDate     =   38176
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
      Left            =   6600
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
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
      Left            =   5280
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvDO 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
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
   Begin VB.TextBox txtChq 
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Text            =   "000000"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtAmt 
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cmbCustomer 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Date:"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label owing 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Amount Owing:"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cheque No:"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Amount:"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Customer:"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmCustomer_Payment.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNotes"
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   6975
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
Attribute VB_Name = "frmCustomer_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCustomer_Click()
If cmbCustomer.Text <> "" Then
    Dim tempRS As Recordset
    RSOpen tempRS, "SELECT CustomerID FROM Customers WHERE Name='" & cmbCustomer.Text & "'", dbOpenSnapshot
    If Not tempRS.EOF Then
        cmbCustomer.Tag = tempRS("CustomerID")
    End If
    tempRS.Close
    Set tempRS = Nothing
    loadInvDO cmbCustomer.Tag
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If cmbCustomer.Text = "" Then
    ValidMsg "Please select a customer.", "Missing customer"
    cmbCustomer.SetFocus
ElseIf Val(txtAmt.Text) = 0 Then
    ValidMsg "Please enter an amount of more than $0.", "Invalid value"
    txtAmt.SetFocus
ElseIf txtChq.Text = "" Then
    ValidMsg "Please enter a cheque number.", "Missing cheque number"
    txtChq.SetFocus
ElseIf Val(txtAmt.Text) > Val(owing.Caption) Then
    ValidMsg "The amount entered is more than the amount owed. Please try again.", "Invalid amount"
    txtAmt.SetFocus
Else
    Dim savRS As Recordset
    Dim savSQL As String
    savSQL = "SELECT * FROM cust_transactions"
    Set savRS = MySynonDatabase.OpenRecordset(savSQL, dbOpenDynaset, dbAppendOnly)
    savRS.AddNew
    savRS("date") = Format$(datepk.Day, "00") & "/" & Format$(datepk.Month, "00") & "/" & Format$(datepk.Year, "0000")
    savRS("CustomerID") = cmbCustomer.Tag
    savRS("DOnumber") = lvDO.SelectedItem.Text
    savRS("credit") = txtAmt.Text
    savRS("notes") = "Payment with chq no: " & txtChq.Text
    savRS.Update
    
    RSOpen savRS, "SELECT Status FROM Delivery WHERE DOnumber='" & lvDO.SelectedItem.Text & "';", dbOpenDynaset
    savRS.Edit
    savRS("Status") = IIf(((Val(txtAmt.Text) < Val(owing.Caption))), "PARTIAL", "PAID")
    savRS.Update
    
    RSOpen savRS, "SELECT CurrentBalance FROM Customers WHERE CustomerID='" & cmbCustomer.Tag & "';", dbOpenDynaset
    savRS.Edit
    savRS("CurrentBalance") = savRS("CurrentBalance") - Val(txtAmt.Text)
    savRS.Update
    
    savRS.Close
    Set savRS = Nothing
    
    InfoMsg "Customer payment has been successfully updated.", "Record saved"
    newPayment
End If
End Sub

Private Sub newPayment()
cmbCustomer.ListIndex = -1
txtAmt.Text = "0.00"
txtChq.Text = "000000"
owing.Caption = "0.00"
lvDO.ListItems.Clear
End Sub

Private Sub Form_Load()
'Insert notes here
lblNotes.Caption = "Payment by customers are recorded here. Ensure all the required details have been entered correctly." & vbCrLf & _
"The date is the day the cheque is banked in and not the clearance date."
FillCombo cmbCustomer, "SELECT Name FROM Customers;", "Name"
lvDO.ColumnHeaders.add , , "DO Number", 1400
lvDO.ColumnHeaders.add , , "Amount", 1200
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
lblNotes.width = Me.ScaleWidth - (lblNotes.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCustomer_Payment = Nothing
End Sub

Private Function totalDOPaid(ByVal strDO As String, ByVal strCust As String) As Double
Dim totalRS As Recordset, totalSQL As String
totalSQL = "SELECT CustomerID, Sum(credit) AS SumPayment, DOnumber " & _
            "FROM cust_transactions WHERE CustomerID='" & strCust & "' AND DOnumber='" & strDO & "' " & _
            "GROUP BY CustomerID, DOnumber;"
On Error GoTo ErrHandler
RSOpen totalRS, totalSQL, dbOpenSnapshot
If Not totalRS.EOF Then
    totalDOPaid = totalRS("SumPayment")
Else
    totalDOPaid = 0
End If
totalRS.Close
Set totalRS = Nothing
ErrHandler:
If Err.Number <> 0 Then
    totalDOPaid = 0
End If
End Function

Private Sub loadInvDO(ByVal strCust As String)
'Loads the delivery orders that have been invoiced only
Dim invRS As Recordset, invSQL As String
invSQL = "SELECT Delivery.DOnumber, Sum([Delivery].[Charges]+([D_Details].[Quantity]*[D_Details].[SalePrice])) AS Total " & _
        "FROM Delivery INNER JOIN D_Details ON Delivery.DOnumber = D_Details.DOnumber " & _
        "WHERE ((Delivery.CustomerID='" & strCust & "') AND ((Delivery.Status)<>'PAID' And (Delivery.Status)<>'DELIVERING')) " & _
        "GROUP BY Delivery.DOnumber;"
RSOpen invRS, invSQL, dbOpenDynaset
If Not invRS.EOF Then
    lvDO.ListItems.Clear
    While Not invRS.EOF
        lvDO.ListItems.add , , invRS("DOnumber")
        lvDO.ListItems(lvDO.ListItems.Count).SubItems(1) = Format$(invRS("Total"), "#,##0.00")
        invRS.MoveNext
    Wend
End If
End Sub

Private Sub lvDO_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        owing.Caption = Format$(CDbl(.SubItems(1)) - totalDOPaid(.Text, cmbCustomer.Tag), "#,##0.00")
    End If
End With
End Sub

Private Sub txtAmt_GotFocus()
SelText txtAmt
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum txtAmt
End If
End Sub

Private Sub txtAmt_LostFocus()
If txtAmt.Text = "" Then
    txtAmt.Text = "0"
End If
txtAmt.Text = Format$(txtAmt.Text, "#,##0.00")
End Sub

Private Sub txtChq_GotFocus()
SelText txtChq
End Sub

Private Sub txtChq_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub
