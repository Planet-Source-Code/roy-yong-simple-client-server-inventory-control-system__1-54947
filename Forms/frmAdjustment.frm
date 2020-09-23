VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Adjustments"
   ClientHeight    =   2505
   ClientLeft      =   -135
   ClientTop       =   435
   ClientWidth     =   7875
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
   ScaleHeight     =   2505
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker cmbDate 
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   19660803
      CurrentDate     =   38157
      MinDate         =   30317
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
      Left            =   6720
      TabIndex        =   5
      ToolTipText     =   "Click here to save the adjustments."
      Top             =   1560
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
      Left            =   6720
      TabIndex        =   6
      ToolTipText     =   "Click here to close this window without saving any changes."
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton credit 
      Caption         =   "Charges, etc"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton debit 
      Caption         =   "Payments, etc"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox amount 
      Height          =   285
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox description 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Date:"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Amount:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAdjustment.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNotes"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   7
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
Attribute VB_Name = "frmAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim accType As String

Public Sub setAccType(ByVal strAccType As String)
accType = strAccType
With description
    .Clear
    If strAccType = "Customer" Then
        .addItem "Payment for DO - "
        .addItem "Interest charged on overdue amount"
    Else
        .addItem "Invoiced for PO - "
        .addItem "Payment for PO - "
    End If
End With
End Sub
Private Sub amount_GotFocus()
SelText amount
End Sub

Private Sub amount_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub amount_LostFocus()
If amount.Text <> "" Then
    amount.Text = Format$(amount.Text, "#,##0.00")
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If ((debit.Value = False) And (credit.Value = False)) Then
    ValidMsg "Please select the type of adjustments.", "Missing selection"
    debit.SetFocus
ElseIf Len(description.Text) > 100 Then
    ValidMsg "The description is too long. Limited length of 100 characters.", "Invalid description"
    description.SetFocus
ElseIf Val(amount.Text) <= 0 Then
    ValidMsg "Please enter an amount of more than 0.", "Invalid amount"
    amount.SetFocus
Else
    'Completed validation
    'Start saving
    If accType = "Supplier" Then
        processSuppTransaction Format$(cmbDate.Value, "dd/mm/yyyy"), description.Text, IIf((debit.Value = True), debit, credit), Me.Tag, CSng(amount.Text)
    Else
        processCustTransaction Format$(cmbDate.Value, "dd/mm/yyyy"), description.Text, IIf((debit.Value = True), credit, debit), Me.Tag, CSng(amount.Text)
    End If
    InfoMsg "The account " & Me.Tag & " has been successfully adjusted.", "Record saved"
    'For tracking purposes
    insertLog "Account " & Me.Tag & " adjusted. Amount: " & amount.Text & " on " & Format$(cmbDate.Day, "00") & "/" & Format$(cmbDate.Month, "00") & "/" & cmbDate.Year & "."
    If accType = "Supplier" Then
        
    Else
        'frmCustomers
    End If
    Unload Me
End If

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
End If
End Sub

Private Sub description_GotFocus()
SelText description
End Sub

Private Sub Form_Load()
'Insert notes here
lblNotes.Caption = "Adjust the customer or supplier account by selecting the type of adjustments followed by the description and the amount." & vbCrLf & _
"Administrator privilege required to do so."
cmbDate.Value = Now()
End Sub

Private Sub Form_Resize()
Shape1.width = Me.width
lblNotes.width = Me.ScaleWidth - (lblNotes.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAdjustment = Nothing
End Sub
