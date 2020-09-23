VERSION 5.00
Begin VB.Form frmSelected 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selected Product"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3600
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
   ScaleHeight     =   2760
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Click to add item into order list."
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selected item:"
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtRef 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   2
         ToolTipText     =   "Enter customer reference here if any."
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "0.00"
         ToolTipText     =   "Enter the agreed sale price."
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Enter the quantity ordered."
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Enter the description here."
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtProductID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "This is the product ID. [Fixed]"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   315
         ItemData        =   "frmSelected.frx":0000
         Left            =   1800
         List            =   "frmSelected.frx":000D
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Product ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Customer Ref:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Sale Price:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSelected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private qtyAvailable As Integer, targetDO As String

Private Sub cmbUnit_GotFocus()
SelText cmbUnit
End Sub

Private Sub cmbUnit_LostFocus()
If (Len(txtQuantity.Text) > 0) And (Val(txtQuantity.Text) > 1) Then
    If cmbUnit.Text <> "" Then
        Dim strQ As String
        strQ = UCase(Right$(cmbUnit.Text, 1))
        Select Case strQ
        Case "X"
            cmbUnit.Text = cmbUnit.Text & "ES"
        Case Else
            cmbUnit.Text = cmbUnit.Text & "S"
        End Select
    End If
End If

End Sub

Private Sub cmdAdd_Click()
If Len(txtProductID.Text) < 1 Then
    ValidMsg "Please enter a product ID of the product.", "Missing entry"
    txtProductID.SetFocus
ElseIf Len(txtDescription.Text) < 1 Then
    ValidMsg "Please enter a description of the product.", "Missing entry"
    txtDescription.SetFocus
ElseIf Len(txtPrice.Text) < 1 Then
    ValidMsg "Please enter a price for the product.", "Missing entry"
    txtPrice.SetFocus
ElseIf Len(txtQuantity.Text) < 1 Then
    ValidMsg "Please enter a quantity for the product.", "Missing entry"
    txtQuantity.SetFocus
ElseIf Val(txtPrice.Text) < 0 Then
    'Allows zero dollar because some items are only priced after being supplied
    ValidMsg "Please enter a price of more than $0.", "Invalid value"
    txtPrice.SetFocus
ElseIf ((Val(txtQuantity.Text) < 1) Or (Val(txtQuantity.Text) > qtyAvailable)) Then
    ValidMsg "Please enter a quantity of more than 0.", "Invalid value"
    txtQuantity.SetFocus
Else
    'return pass to designated delivery order
    Dim frm As Form
    
    For Each frm In Forms
        If StrComp(frm.Tag, targetDO, vbTextCompare) = 0 Then
            frm.addItem txtProductID.Text, txtDescription.Text, txtRef.Text, CInt(txtQuantity.Text), cmbUnit.Text, CSng(txtPrice.Text)
            frmProduct_Browse.setStatus txtQuantity.Text & " " & cmbUnit.Text & " " & IIf((Val(txtQuantity.Text) > 1), " have ", " has ") & "been added to " & targetDO
            Exit For
        End If
    Next
    Unload Me
End If
End Sub

Private Sub txtDescription_GotFocus()
SelText txtDescription
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
OnlyAlpha KeyAscii
End Sub

Private Sub txtPrice_GotFocus()
SelText txtPrice
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtPrice_LostFocus()
If txtPrice.Text <> "" Then
    txtPrice.Text = Format$(txtPrice.Text, "#,##0.00")
End If
End Sub

Private Sub txtProductID_GotFocus()
SelText txtProductID
End Sub

Private Sub txtQuantity_GotFocus()
SelText txtQuantity
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub txtRef_GotFocus()
SelText txtRef
End Sub

Public Sub add(ByRef strListitem As ListItem, ByVal frmDONumber As String)
targetDO = frmDONumber
With strListitem
    txtProductID.Text = .Text
    txtDescription.Text = .SubItems(3) & Space(1) & .SubItems(2) & Space(1) & .SubItems(1)
    qtyAvailable = .SubItems(4)
End With
Me.Show vbModal
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
tickerKeys KeyAscii
End Sub
