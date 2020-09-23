VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomer_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Customer"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
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
   Icon            =   "frmCustomer_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCustomer_New.frx":000C
   ScaleHeight     =   3930
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   "[PLEASE SELECT ONE]"
      ToolTipText     =   "Country of address"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   6720
      TabIndex        =   32
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtTerm 
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3360
      Width           =   480
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   6840
      TabIndex        =   16
      ToolTipText     =   "Click here to save the new customer details."
      Top             =   3480
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
      Left            =   8040
      TabIndex        =   17
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   0
      Top             =   840
      Width           =   7695
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   1
      Top             =   1200
      Width           =   7695
   End
   Begin VB.ComboBox cmbCity 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   2280
      Width           =   3735
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "[PLEASE SELECT ONE]"
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "00000"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtPhone1 
      Height          =   285
      Index           =   0
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "000"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtPhone2 
      Height          =   285
      Index           =   0
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   10
      Text            =   "000"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtFax1 
      Height          =   285
      Index           =   0
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "000"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtFax2 
      Height          =   285
      Index           =   0
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   14
      Text            =   "000"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtPhone1 
      Height          =   285
      Index           =   1
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "00000000"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtPhone2 
      Height          =   285
      Index           =   1
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "00000000"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtFax1 
      Height          =   285
      Index           =   1
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "00000000"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtFax2 
      Height          =   285
      Index           =   1
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   15
      Text            =   "00000000"
      Top             =   2640
      Width           =   1815
   End
   Begin MSComCtl2.UpDown udTerm 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   30
      BuddyControl    =   "txtTerm"
      BuddyDispid     =   196612
      OrigLeft        =   1920
      OrigTop         =   3360
      OrigRight       =   2160
      OrigBottom      =   3645
      Increment       =   30
      Max             =   360
      Min             =   30
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "days"
      Height          =   255
      Left            =   2280
      TabIndex        =   34
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   33
      Top             =   120
      Width           =   8415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmCustomer_New.frx":08D6
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
      Width           =   9255
   End
   Begin VB.Label Label2 
      Caption         =   "Email:"
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Limit:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Term:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "State:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Country:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phone (1):"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   22
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phone (2):"
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   21
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax (1):"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax (2):"
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   7200
      X2              =   7320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   7200
      X2              =   7320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   7200
      X2              =   7320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      X1              =   7200
      X2              =   7320
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frmCustomer_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CheckCustomerFields()
If (Len(txtName.Text) = 0) Or (Len(txtAddress.Text) = 0) Or (Len(cmbCountry.Text) = 0) Or _
(Len(cmbState.Text) = 0) Or (Len(cmbCity.Text) = 0) Or (Len(txtPhone1(0).Text) = 0) Or (Len(txtPhone1(1).Text) = 0) Or _
(Len(txtFax1(0).Text) = 0) Or (Len(txtFax1(1).Text) = 0) Or (Len(txtLimit.Text) = 0) Or (Len(txtTerm.Text) = 0) Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If
End Sub

Private Sub cmbCity_Change()
CheckCustomerFields
End Sub

Private Sub cmbCity_GotFocus()
If cmbCity.Text = "[PLEASE SELECT ONE]" Then
    cmbCity.Text = ""
End If
SelText cmbCity
End Sub

Private Sub cmbCity_LostFocus()
CapCon cmbCity
End Sub

Private Sub cmbCountry_Change()
CheckCustomerFields
End Sub

Private Sub cmbCountry_Click()
FillComboState cmbState, cmbCountry.Text
End Sub

Private Sub cmbCountry_GotFocus()
If cmbCountry.Text = "[PLEASE SELECT ONE]" Then
    cmbCountry.Text = ""
End If
SelText cmbCountry
End Sub

Private Sub cmbCountry_LostFocus()
CapCon cmbCountry
End Sub

Private Sub cmbState_Change()
CheckCustomerFields
End Sub

Private Sub cmbState_Click()
FillComboCity cmbCity, cmbState.Text
End Sub

Private Sub cmbState_GotFocus()
If cmbState.Text = "[PLEASE SELECT ONE]" Then
    cmbState.Text = ""
End If
SelText cmbState
End Sub

Private Sub cmbState_LostFocus()
CapCon cmbState
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If CCur(txtLimit.Text) <= 0 Then
    ValidMsg "Credit limit has to been more than $0.", "Invalid value"
    txtLimit.SetFocus
Else
    Dim saveSQL As String, tempChar As String
    Dim newID As Long
    'Get unique ID
    tempChar = Left$(txtName.Text, 1)
    If IsNumeric(tempChar) = True Then
        tempChar = "#"
    End If
    'define query
    saveSQL = "SELECT Ref_Number FROM Customer_IDs WHERE InitialID='" & tempChar & "';"
    On Error GoTo ErrHandler
    'execute query
    Dim SaveRS As Recordset, tempRS As Recordset
    RSOpen tempRS, saveSQL, dbOpenDynaset
    'Obtain the reference number needed to form unique ID
    newID = tempRS("Ref_Number")
    
    saveSQL = "SELECT * FROM Customers;"
    RSOpen SaveRS, saveSQL, dbOpenDynaset
    SaveRS.AddNew
    SaveRS("CustomerID") = tempChar & Format$(newID, "0000")
    SaveRS("Name") = txtName.Text
    SaveRS("Address") = txtAddress.Text
    SaveRS("City") = cmbCity.Text
    SaveRS("State") = cmbState.Text
    SaveRS("Zip") = txtZip.Text
    SaveRS("Country") = cmbCountry.Text
    SaveRS("ACPhone1") = txtPhone1(0).Text
    SaveRS("ACPhone2") = txtPhone2(0).Text
    SaveRS("Phone1") = txtPhone1(1).Text
    SaveRS("Phone2") = txtPhone2(1).Text
    SaveRS("ACFax1") = txtFax1(0).Text
    SaveRS("ACFax2") = txtFax2(0).Text
    SaveRS("Fax1") = txtFax1(1).Text
    SaveRS("Fax2") = txtFax2(1).Text
    SaveRS("Email") = txtEmail.Text
    SaveRS("CreditLimit") = txtLimit.Text
    SaveRS("CreditTerm") = txtTerm.Text
    SaveRS("CurrentBalance") = "0.00"
    SaveRS.Update

    tempRS.Edit
    tempRS("Ref_Number") = newID + 1
    tempRS.Update

    tempRS.Close
    SaveRS.Close
    Set tempRS = Nothing
    Set SaveRS = Nothing
    'Insert into systems log
    insertLog "Customer ID: " & newID & " account has been created."
    InfoMsg "Customer ID: " & tempChar & Format$(newID, "0000") & vbCrLf & "New customer record has been successfully created.", "Record saved"
    
    Unload Me
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub Form_Load()
isOpen = False
FillComboCountry cmbCountry
lblHeader.Caption = "Please enter the details of the new customer accurately." & vbNewLine & _
"Entries will be converted to upper-case automatically."
DisableClose Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCustomer_New = Nothing
End Sub

Private Sub txtAddress_Change()
CheckCustomerFields
End Sub

Private Sub txtAddress_GotFocus()
SelText txtAddress
End Sub

Private Sub txtAddress_LostFocus()
CapCon txtAddress
End Sub

Private Sub txtEmail_Change()
CheckCustomerFields
End Sub

Private Sub txtEmail_GotFocus()
SelText txtEmail
End Sub

Private Sub txtFax1_Change(Index As Integer)
CheckCustomerFields
End Sub

Private Sub txtFax1_GotFocus(Index As Integer)
SelText txtFax1(Index)
End Sub

Private Sub txtFax1_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtFax1_LostFocus(Index As Integer)
If Index = 0 Then
    If (IsNull(txtFax1(Index).Text) = False) Or (txtFax1(Index).Text <> "") Then
        txtFax1(Index).Text = Format$(txtFax1(Index).Text, "000")
    End If
End If

End Sub

Private Sub txtFax2_Change(Index As Integer)
CheckCustomerFields
End Sub

Private Sub txtFax2_GotFocus(Index As Integer)
SelText txtFax2(Index)
End Sub

Private Sub txtFax2_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtFax2_LostFocus(Index As Integer)
If Index = 0 Then
    If (IsNull(txtFax2(Index).Text) = False) Or (txtFax2(Index).Text <> "") Then
        txtFax2(Index).Text = Format$(txtFax2(Index).Text, "000")
    End If
End If
End Sub

Private Sub txtLimit_Change()
CheckCustomerFields
End Sub

Private Sub txtLimit_GotFocus()
SelText txtLimit
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtName_Change()
CheckCustomerFields
End Sub

Private Sub txtName_GotFocus()
SelText txtName
End Sub

Private Sub txtName_LostFocus()
CapCon txtName
End Sub

Private Sub txtPhone1_Change(Index As Integer)
CheckCustomerFields
End Sub

Private Sub txtPhone1_GotFocus(Index As Integer)
SelText txtPhone1(Index)
End Sub

Private Sub txtPhone1_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtPhone1_LostFocus(Index As Integer)
If Index = 0 Then
    If (IsNull(txtPhone1(Index).Text) = False) Or (txtPhone1(Index).Text <> "") Then
        txtPhone1(Index).Text = Format$(txtPhone1(Index).Text, "000")
    End If
End If
End Sub

Private Sub txtPhone2_Change(Index As Integer)
CheckCustomerFields
End Sub

Private Sub txtPhone2_GotFocus(Index As Integer)
SelText txtPhone2(Index)
End Sub

Private Sub txtPhone2_KeyPress(Index As Integer, KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtPhone2_LostFocus(Index As Integer)
If Index = 0 Then
    If (IsNull(txtPhone2(Index).Text) = False) Or (txtPhone2(Index).Text <> "") Then
        txtPhone2(Index).Text = Format$(txtPhone2(Index).Text, "000")
    End If
End If

End Sub

Private Sub txtTerm_Change()
CheckCustomerFields
End Sub

Private Sub txtTerm_GotFocus()
SelText txtTerm
End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtZip_Change()
CheckCustomerFields
End Sub

Private Sub txtZip_GotFocus()
SelText txtZip
End Sub

Private Sub txtZip_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub udTerm_Change()
CheckCustomerFields
End Sub
