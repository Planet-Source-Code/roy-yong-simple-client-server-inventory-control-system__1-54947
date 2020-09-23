VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport_Inventory 
   Caption         =   "Inventory Transactions"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport_Inventory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9615
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
      Left            =   8280
      TabIndex        =   5
      ToolTipText     =   "Click here to close this window."
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "&Filter"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvHis 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
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
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   0
      ItemData        =   "frmReport_Inventory.frx":08CA
      Left            =   1320
      List            =   "frmReport_Inventory.frx":092B
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   1
      ItemData        =   "frmReport_Inventory.frx":09AB
      Left            =   2040
      List            =   "frmReport_Inventory.frx":09D3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Month"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   2
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a date:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isLevel As Boolean
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFilter_Click()
If (cmbDate(0).Text = "") Or (cmbDate(1).Text = "") Or (cmbDate(2).Text = "") Then
    ValidMsg "Please select the date.", "Invalid date"
Else
    If isDateValid(CByte(cmbDate(0).Text), CByte(cmbDate(1).Text), CInt(cmbDate(2).Text)) = True Then
        getTransactions
    Else
        ValidMsg "The selected date is invalid. Please try again.", "Invalid date"
    End If
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 5
    cmbDate(2).addItem Format$(Year(Now()) - i, "0000"), i
Next i
cmbDate(0).Text = Format$(Day(Now()), "00")
cmbDate(1).Text = Format$(Month(Now()), "00")
cmbDate(2).Text = Format$(Year(Now()), "0000")
End Sub

Private Sub Form_Resize()
lvHis.width = Me.ScaleWidth - (lvHis.Left * 2)
lvHis.height = Me.ScaleHeight - (lvHis.Top * 2)
cmdClose.Top = Me.ScaleHeight - (cmdClose.height + 115)
cmdClose.Left = Me.ScaleWidth - (cmdClose.width + 115)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmReport_Inventory = Nothing
End Sub

Public Sub getLevel()
'Set the form layout
Me.Caption = "Inventory - Products under-stock"
Dim i As Integer
For i = 0 To 2
    cmbDate(i).Visible = False
Next i
cmdFilter.Visible = False
lvHis.Top = 120
lvHis.height = lvHis.height + 480
isLevel = True
With lvHis.ColumnHeaders
    .add , , "Product ID", 950
    .add , , "Description", 3000
    .add , , "Brand", 1500
    .add , , "Category ID", 1500
    .add , , "Quantity", 900
    .add , , "Min Level", 950
    .add , , "Reorder Level", 980
    .add , , "Location", 1200
    If getSettings("allowPrice") = "TRUE" Then
        .add , , "Unit Price"
    End If
End With
Dim tSQL As String
tSQL = "SELECT * FROM Products WHERE Quantity < MinLevel;"
Dim tRS As Recordset
RSOpen tRS, tSQL, dbOpenSnapshot
With lvHis.ListItems
    While Not tRS.EOF
        .add , , tRS("ProductID")
        .Item(.Count).SubItems(1) = tRS("Description")
        .Item(.Count).SubItems(2) = tRS("Brand")
        .Item(.Count).SubItems(3) = tRS("CategoryID")
        .Item(.Count).SubItems(4) = tRS("Quantity")
        .Item(.Count).SubItems(5) = IIf(IsNull(tRS("MinLevel")), "0", tRS("MinLevel"))
        .Item(.Count).SubItems(6) = IIf(IsNull(tRS("ReorderLevel")), "0", tRS("ReorderLevel"))
        .Item(.Count).SubItems(7) = IIf(IsNull(tRS("Location")), "", tRS("Location"))
        If getSettings("allowPrice") = "TRUE" Then
            .Item(.Count).SubItems(8) = Format$(tRS("UnitPrice"), "#,##0.00")
        End If
        tRS.MoveNext
    Wend
    tRS.Close
    Set tRS = Nothing
End With
End Sub

Public Sub getTransactions()
'Set the form layout
Me.Caption = "Inventory - Transactions"
lvHis.Top = 600
isLevel = False
Dim tRS As Recordset
Dim tSQL As String
With lvHis.ColumnHeaders
    .Clear
    .add , , "Product ID"
    .add , , "Description", 4000
    .add , , "In", 500
    .add , , "Out", 500
End With
tSQL = "SELECT Products.Description, Internal_Transaction.ProductID, Internal_Transaction.qty, Internal_Transaction.isIn FROM Products INNER JOIN Internal_Transaction ON Products.ProductID=Internal_Transaction.ProductID WHERE Date ='" & cmbDate(0).Text & "/" & cmbDate(1).Text & "/" & cmbDate(2).Text & "';"
With lvHis.ListItems
    .Clear
    RSOpen tRS, tSQL, dbOpenSnapshot
    While Not tRS.EOF
        .add , , tRS("ProductID")
        .Item(.Count).SubItems(1) = tRS("Description")
        If tRS("isIn") = True Then
            .Item(.Count).SubItems(2) = tRS("qty")
        Else
            .Item(.Count).SubItems(3) = tRS("qty")
        End If
        tRS.MoveNext
    Wend
    tRS.Close
    Set tRS = Nothing
End With
End Sub

