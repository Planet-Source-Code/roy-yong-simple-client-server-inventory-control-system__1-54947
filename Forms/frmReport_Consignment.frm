VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport_Consignment 
   Caption         =   "Consignment Transactions"
   ClientHeight    =   5505
   ClientLeft      =   -120
   ClientTop       =   450
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport_Consignment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6240
   Begin VB.ComboBox customers 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select the customer with a consignment contract."
      Top             =   360
      Width           =   4455
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
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Click here to close this window."
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details:"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      Begin VB.TextBox endDate 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox startDate 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox id 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Contract ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvCon 
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
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
   Begin VB.Label Label1 
      Caption         =   "Consignment Contract:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmReport_Consignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub getContractInfo(ByVal strCustomers As String)
Dim conRS As Recordset, custID As String
custID = ""
On Error Resume Next
RSOpen conRS, "SELECT Contracts.ContractNo, Contracts.StartDate, Contracts.ExpireDate " & _
            "FROM Contracts INNER JOIN Customers ON Contracts.CustomerID = Customers.CustomerID " & _
            "WHERE (((Customers.Name)='" & strCustomers & "'));", dbOpenSnapshot
If Not conRS.EOF Then
    custID = conRS("ContractNo")
    id.Text = conRS("ContractNo")
    startDate.Text = conRS("StartDate")
    endDate.Text = IIf(IsNull(conRS("ExpireDate")), "", conRS("ExpireDate"))
Else
    custID = ""
    id.Text = ""
    startDate.Text = ""
    endDate.Text = ""
End If
If Not custID = "" Then
    RSOpen conRS, "SELECT External_Transaction.Date, External_Transaction.ProductID, Products.Description, External_Transaction.qty, External_Transaction.isIn " & _
                "FROM (Contracts INNER JOIN External_Transaction ON Contracts.ContractNo = External_Transaction.ContractNo) INNER JOIN Products ON External_Transaction.ProductID = Products.ProductID " & _
                "WHERE (((External_Transaction.ContractNo)='" & custID & "'));", dbOpenSnapshot
    While Not conRS.EOF
        With lvCon.ListItems
            .add , , conRS("Date")
            .Item(.Count).SubItems(1) = conRS("ProductID")
            .Item(.Count).SubItems(2) = conRS("Description")
            If conRS("isIn") = True Then
                .Item(.Count).SubItems(3) = conRS("qty")
            Else
                .Item(.Count).SubItems(4) = conRS("qty")
            End If
            conRS.MoveNext
        End With
    Wend
End If
conRS.Close
Set conRS = Nothing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub customers_Click()
If customers.Text <> "" Then
    getContractInfo customers.Text
End If
End Sub

Private Sub Form_Load()
FillCombo customers, "SELECT Name FROM Customers;", "Name"
With lvCon
    .View = lvwReport
    .ColumnHeaders.add , , "Date", 960
    .ColumnHeaders.add , , "Product ID", 980
    .ColumnHeaders.add , , "Description", 2000
    .ColumnHeaders.add , , "In", 700
    .ColumnHeaders.add , , "Out", 700
End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
lvCon.width = Me.ScaleWidth - lvCon.Left * 2
lvCon.height = Me.ScaleHeight - (lvCon.Top + 115)
Frame1.width = lvCon.width - cmdClose.width - Frame1.Left
cmdClose.Left = Frame1.Left * 2 + Frame1.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmReport_Consignment = Nothing
End Sub
