VERSION 5.00
Begin VB.Form frmProduct_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Product"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtBrand 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
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
      Left            =   3480
      TabIndex        =   8
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
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
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Click here to add the new product now."
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtReorder 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtMinimum 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cmbCategoryID 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmProduct_New.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Red labels indicate required fields."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Reorder Level:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Minimum Level:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Unit Price:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Category ID:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Brand:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmProduct_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CheckFields()
If (Len(txtDescription.Text) = 0) Or (Len(cmbCategoryID.Text) = 0) Or _
(Len(txtMinimum.Text) = 0) Or (Len(txtUnitPrice.Text) = 0) Then
    cmdAdd.Enabled = False
Else
    cmdAdd.Enabled = True
End If
End Sub

Private Sub cmbCategoryID_Change()
CheckFields
End Sub

Private Sub cmdAdd_Click()
If (Val(txtUnitPrice.Text) < 0) Then
    ValidMsg "Please enter a price of more than $0.", "Invalid value"
    txtUnitPrice.SetFocus
ElseIf (Val(txtMinimum.Text) < 0) Then
    ValidMsg "Please enter a value of more than 0.", "Invalid value"
    txtMinimum.SetFocus
ElseIf (Val(txtReorder.Text) < 0) Then
    ValidMsg "Please enter a value of more than 0.", "Invalid value"
    txtReorder.SetFocus
Else
    Dim tempSQL As String, tmpProductID As String
    Dim saveRS As Recordset, newProdID As Recordset
    
    On Error GoTo ErrHandler
    'get latest key
    Screen.MousePointer = 11
    tempSQL = "SELECT DataValue FROM Misc WHERE DataType='PRODUCT'"
    RSOpen newProdID, tempSQL, dbOpenDynaset
    If Not newProdID.EOF Then
        tmpProductID = newProdID("DataValue")
    End If
    tempSQL = "SELECT * FROM Products;"
    RSOpen saveRS, tempSQL, dbOpenDynaset
    saveRS.AddNew
    saveRS("ProductID") = tmpProductID
    saveRS("Description") = txtDescription.Text
    saveRS("Brand") = IIf(IsNull(txtBrand.Text), "UNKNOWN", txtBrand.Text)
    saveRS("CategoryID") = cmbCategoryID.Text
    saveRS("UnitPrice") = CSng(txtUnitPrice.Text)
    saveRS("MinLevel") = txtMinimum.Text
    saveRS("ReorderLevel") = txtReorder.Text
    saveRS("Location") = IIf(IsNull(txtLocation.Text), "", txtLocation.Text)
    saveRS.Update
    
    'update latest key
    newProdID.Edit
    newProdID("DataValue") = CStr(CLng(tmpProductID) + 1)
    newProdID.Update
    
    newProdID.Close
    saveRS.Close
    Set newProdID = Nothing
    Set saveRS = Nothing
    Screen.MousePointer = 0
    tempSQL = "Product ID: " & tmpProductID & vbCrLf & "The new product has been successfully added."
    insertLog tempSQL
    InfoMsg tempSQL, "Record saved"
    Unload Me
End If
ErrHandler:
If Err.Number <> 0 Then
    If Err.Number = 3022 Then
        ErrorNotifier Err.Number, "The new preset key exist as a primary key in another record. The new preset key has to be changed before adding a new product."
    Else
        ErrorNotifier Err.Number, Err.description
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
FillCombo cmbCategoryID, "SELECT CategoryID FROM Categories;", "CategoryID"
FillCombo txtBrand, "SELECT DISTINCT Brand FROM Products ORDER BY Brand ASC;", "Brand"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmProduct_New = Nothing
End Sub

Private Sub txtBrand_Click()
CheckFields
End Sub

Private Sub txtBrand_GotFocus()
SelText txtBrand
End Sub

Private Sub txtBrand_LostFocus()
CapCon txtBrand
End Sub

Private Sub txtDescription_Change()
CheckFields
End Sub

Private Sub txtDescription_GotFocus()
SelText txtDescription
End Sub

Private Sub txtDescription_LostFocus()
CapCon txtDescription
End Sub

Private Sub txtLocation_Change()
CheckFields
End Sub

Private Sub txtLocation_GotFocus()
SelText txtLocation
End Sub

Private Sub txtMinimum_Change()
CheckFields
End Sub

Private Sub txtMinimum_GotFocus()
SelText txtMinimum
End Sub

Private Sub txtMinimum_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtReorder_Change()
CheckFields
End Sub

Private Sub txtReorder_GotFocus()
SelText txtReorder
End Sub

Private Sub txtReorder_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtUnitPrice_Change()
CheckFields
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
If Len(txtUnitPrice.Text) > 0 Then
    txtUnitPrice.Text = Format$(txtUnitPrice.Text, "#,##0.00")
End If
End Sub
