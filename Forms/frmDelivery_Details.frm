VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDelivery_Details 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delivery Order Details"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7425
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
   ScaleHeight     =   4170
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox desc 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2400
      Width           =   6015
   End
   Begin VB.TextBox cust 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox qty 
      Height          =   285
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox price 
      Height          =   285
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox label 
      Height          =   315
      ItemData        =   "frmDelivery_Details.frx":0000
      Left            =   1200
      List            =   "frmDelivery_Details.frx":0016
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
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
      Left            =   6240
      TabIndex        =   10
      ToolTipText     =   "Click here to edit the selected item."
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox id 
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
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
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "Click here to close this window."
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvDet 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "Click here to save any changes."
      Top             =   3240
      Width           =   1095
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
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Click here to cancel editing."
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Cust Ref:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Unit Label:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Unit Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Product ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblhidden 
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmDelivery_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub getDetails(ByVal strDOnumber As String)
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
        End With
        gRS.MoveNext
    Wend
    gRS.Close
    Set gRS = Nothing
End If
End Sub

Private Sub setFormMode(ByVal strStatus As ModeStatus)
If strStatus = Editing Then
    id.Enabled = True
    desc.Enabled = True
    cust.Enabled = True
    qty.Enabled = True
    label.Enabled = True
    price.Enabled = True
    lvDet.Enabled = False
    cmdEdit.Visible = False
    cmdClose.Visible = False
    cmdSave.Visible = True
    cmdCancel.Visible = True
Else
    id.Enabled = False
    desc.Enabled = False
    cust.Enabled = False
    qty.Enabled = False
    label.Enabled = False
    price.Enabled = False
    lvDet.Enabled = True
    cmdEdit.Visible = True
    cmdClose.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End If
End Sub

Private Sub cmdCancel_Click()
id.Text = id.Tag
desc.Text = desc.Tag
cust.Text = cust.Tag
qty.Text = qty.Tag
label.Text = label.Tag
price.Text = price.Tag
setFormMode Viewing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
If id.Text = "" Then
    InfoMsg "Please select an item first.", "Missing selection"
Else
    id.Tag = id.Text
    desc.Tag = desc.Text
    cust.Tag = cust.Text
    qty.Tag = qty.Text
    label.Tag = label.Text
    price.Tag = price.Text
    lblhidden.Caption = id.Text
    setFormMode Editing
End If
End Sub

Private Sub cmdSave_Click()
If id.Text = "" Then
    ValidMsg "Please enter a product ID here. Be careful as this ID may affect the accuracy of the system.", "Missing product ID"
    id.SetFocus
ElseIf desc.Text = "" Then
    ValidMsg "Please enter a description.", "Missing description"
    desc.SetFocus
ElseIf ((qty.Text = "") Or (Val(qty.Text) > 30000) Or (Val(qty.Text) < 1)) Then
    ValidMsg "Please enter a quantity between 0 and a maximum of 30000.", "Invalid quantity"
    qty.SetFocus
ElseIf Val(price.Text) < 0 Then
    ValidMsg "Please enter a price of $0 or more.", "Invalid price"
    price.SetFocus
Else
    Dim tmpRS As Recordset
    On Error GoTo ErrHandler
    RSOpen tmpRS, "SELECT * FROM D_Details WHERE DOnumber='" & Me.Tag & "' AND ProductID='" & lblhidden.Caption & "';", dbOpenDynaset
    If Not tmpRS.EOF Then
        tmpRS.Edit
        tmpRS("ProductID") = id.Text
        tmpRS("Description") = desc.Text
        tmpRS("CustRef") = cust.Text
        tmpRS("Quantity") = qty.Text
        tmpRS("UnitLabel") = label.Text
        tmpRS("SalePrice") = price.Text
        tmpRS.Update
        InfoMsg "Record has been successfully updated.", "Record saved"
        setFormMode Viewing
    End If
    tmpRS.Close
    Set tmpRS = Nothing
End If
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "An error has occured when trying to update the record. No changes have been made and please try again." & _
    " If you see this message again, click 'OK' and contact your system administrator.", "Error found"
    Exit Sub
End If
End Sub

Private Sub cust_GotFocus()
SelText cust
End Sub

Private Sub desc_GotFocus()
SelText desc
End Sub

Private Sub Form_Load()
With lvDet
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "Product ID", 975
    .ColumnHeaders.add , , "Description"
    .ColumnHeaders.add , , "Cust Ref"
    .ColumnHeaders.add , , "Quantity", 880
    .ColumnHeaders.add , , "Unit Label", 950
    .ColumnHeaders.add , , "Unit Price", 950
End With
setFormMode Viewing
id.ToolTipText = "Please be careful when changing this Product ID as any changes may not reflect logically in the inventory." & vbNewLine & _
"It is best to leave it as it is."
End Sub

Private Sub label_GotFocus()
SelText label
End Sub

Private Sub lvDet_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        id.Text = .Text
        desc.Text = .SubItems(1)
        cust.Text = .SubItems(2)
        qty.Text = .SubItems(3)
        label.Text = .SubItems(4)
        price.Text = .SubItems(5)
    End If
End With
End Sub

Private Sub price_GotFocus()
SelText price
End Sub

Private Sub price_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub price_LostFocus()
If price.Text <> "" Then
    price.Text = Format$(price.Text, "#,##0.00")
End If
End Sub

Private Sub qty_GotFocus()
SelText qty
End Sub

Private Sub qty_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub
