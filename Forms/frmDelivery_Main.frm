VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDelivery_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Orders Management"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDelivery_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11970
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   10680
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox notes 
      Height          =   285
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   16
      Top             =   7920
      Width           =   5055
   End
   Begin VB.TextBox charges 
      Height          =   285
      Left            =   5520
      MaxLength       =   12
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   7440
      Width           =   1095
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
      Left            =   10680
      TabIndex        =   20
      ToolTipText     =   "Click here to close this window."
      Top             =   7800
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
      Left            =   10680
      TabIndex        =   18
      ToolTipText     =   "Click here to edit the selected Delivery Order."
      Top             =   6840
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
      Left            =   3480
      TabIndex        =   26
      Top             =   960
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
      Left            =   10680
      TabIndex        =   22
      ToolTipText     =   "Click here to cancel editing."
      Top             =   7800
      Width           =   1215
   End
   Begin VB.ComboBox employee 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.ComboBox employee 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6000
      Width           =   2055
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery_Main.frx":08CA
      Left            =   1200
      List            =   "frmDelivery_Main.frx":092B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6000
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery_Main.frx":09AB
      Left            =   1920
      List            =   "frmDelivery_Main.frx":09D3
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Month"
      Top             =   6000
      Width           =   615
   End
   Begin VB.ComboBox cmbDate 
      Height          =   315
      Index           =   2
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6000
      Width           =   855
   End
   Begin VB.ComboBox customer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6480
      Width           =   3015
   End
   Begin VB.TextBox txtPO 
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   10
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtAttn 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   13
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox txtREM 
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   14
      Top             =   7920
      Width           =   2055
   End
   Begin VB.ComboBox cmbDelDate 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery_Main.frx":0A07
      Left            =   5520
      List            =   "frmDelivery_Main.frx":0A68
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6480
      Width           =   615
   End
   Begin VB.ComboBox cmbDelDate 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery_Main.frx":0AE8
      Left            =   6240
      List            =   "frmDelivery_Main.frx":0B10
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Month"
      Top             =   6480
      Width           =   615
   End
   Begin VB.ComboBox cmbDelDate 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6480
      Width           =   855
   End
   Begin VB.ComboBox cmbDelTime 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery_Main.frx":0B44
      Left            =   5520
      List            =   "frmDelivery_Main.frx":0B90
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6960
      Width           =   615
   End
   Begin VB.ComboBox cmbDelTime 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery_Main.frx":0BF4
      Left            =   6240
      List            =   "frmDelivery_Main.frx":0C04
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox doNum 
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.ComboBox filter 
      Height          =   315
      Index           =   0
      ItemData        =   "frmDelivery_Main.frx":0C18
      Left            =   1200
      List            =   "frmDelivery_Main.frx":0C7C
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox filter 
      Height          =   315
      Index           =   1
      ItemData        =   "frmDelivery_Main.frx":0CFE
      Left            =   1920
      List            =   "frmDelivery_Main.frx":0D29
      Style           =   2  'Dropdown List
      TabIndex        =   24
      ToolTipText     =   "Month"
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox filter 
      Height          =   315
      Index           =   2
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   960
      Width           =   735
   End
   Begin MSComctlLib.ListView lvDO 
      Height          =   3975
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7011
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
      Left            =   10680
      TabIndex        =   21
      ToolTipText     =   "Click here to save any changes."
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblDel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   44
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   43
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblcust 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      TabIndex        =   42
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "Description:"
      Height          =   255
      Left            =   4320
      TabIndex        =   41
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Charges:"
      Height          =   255
      Left            =   4320
      TabIndex        =   40
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Endorsed By:"
      Height          =   255
      Left            =   4320
      TabIndex        =   39
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Delivered By:"
      Height          =   255
      Left            =   4320
      TabIndex        =   38
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Customer ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "PO Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Attention:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Remark:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Delivery Date:"
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Delivery Time:"
      Height          =   255
      Left            =   4320
      TabIndex        =   32
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "DO No:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Filter Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDelivery_Main.frx":0D5F
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "lblNotes"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   28
      Top             =   120
      Width           =   11175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnu_Print_Blank 
         Caption         =   "Print on blank paper"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_Print_pre 
         Caption         =   "Print on pre-printed paper"
      End
   End
End
Attribute VB_Name = "frmDelivery_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub getDeliveries(ByVal strSQL As String)
Dim doRS As Recordset
With lvDO
    
    .ListItems.Clear
    RSOpen doRS, strSQL, dbOpenSnapshot
    'On Error GoTo ErrHandler
    While Not doRS.EOF
        .ListItems.add , , doRS("DOnumber")
        .ListItems(.ListItems.Count).SubItems(1) = doRS("Name")
        .ListItems(.ListItems.Count).SubItems(2) = doRS("PONumber")
        .ListItems(.ListItems.Count).SubItems(3) = doRS("DelDate")
        .ListItems(.ListItems.Count).SubItems(4) = doRS("DelTime")
        .ListItems(.ListItems.Count).SubItems(5) = doRS("Status")
        doRS.MoveNext
    Wend
    doRS.Close
    Set doRS = Nothing
End With
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "The query passed is invalid. Please try again.", "Main Delivery Order Console"
    Exit Sub
End If
End Sub

Private Sub setFormMode(ByVal strStatus As ModeStatus)
Dim i As Integer
If strStatus = Editing Then
    For i = 0 To 2
        filter(i).Enabled = False
        cmbDelDate(i).Enabled = True
        cmbDate(i).Enabled = True
        If i < 2 Then
            cmbDelTime(i).Enabled = True
            employee(i).Enabled = True
        End If
    Next i
    doNum.Enabled = True
    customer.Enabled = True
    txtPO.Enabled = True
    txtAttn.Enabled = True
    txtREM.Enabled = True
    notes.Enabled = True
    charges.Enabled = True
    lvDO.Enabled = False
    cmdFilter.Enabled = False
    cmdNew.Visible = False
    cmdDelete.Visible = False
    cmdEdit.Visible = False
    cmdClose.Visible = False
    cmdSave.Visible = True
    cmdCancel.Visible = True
Else
    For i = 0 To 2
        filter(i).Enabled = True
        cmbDelDate(i).Enabled = False
        cmbDate(i).Enabled = False
        If i < 2 Then
            cmbDelTime(i).Enabled = False
            employee(i).Enabled = False
        End If
    Next i
    doNum.Enabled = False
    customer.Enabled = False
    txtPO.Enabled = False
    txtAttn.Enabled = False
    txtREM.Enabled = False
    notes.Enabled = False
    charges.Enabled = False
    lvDO.Enabled = True
    cmdFilter.Enabled = True
    cmdNew.Visible = True
    cmdDelete.Visible = True
    cmdEdit.Visible = True
    cmdClose.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End If
End Sub

Public Sub prePrintDO(ByVal strDO As String)
'Delivery order is designed based on requirement
'Uncommented codes when printing on non-preprinted papers.
Dim printRS As Recordset, tmpX As Long, tmpY As Long, grandtotal As Single
cdlg.DialogTitle = "Print Delivery Order"
cdlg.CancelError = True
On Error Resume Next
cdlg.ShowPrinter
If Err Then
    Exit Sub
Else
    RSOpen printRS, "SELECT Delivery.DOnumber, Delivery.Date, Delivery.IssuedBy, Delivery.EmployeeID, Delivery.PONumber, Delivery.Remark, Delivery.Attn, Delivery.Description, Delivery.Charges, Delivery.DelTime, Delivery.DelDate, Customers.Name, Customers.Address, Customers.City, Customers.State, Customers.Zip, Customers.Country, Customers.CreditTerm " & _
    "FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE Delivery.DOnumber='" & strDO & "';", dbOpenSnapshot
    If Not printRS.EOF Then
        With Printer
            .PaperSize = 9 'A4 size
            .FontSize = CLng(GetFromINI("PrintDO", "FontSize", "", pathFileSettings))
            'Begin printing receiver details
            grandtotal = 0
            tmpX = CLng(GetFromINI("PrintDO", "CustMargin", "", pathFileSettings))
            .CurrentX = tmpX
            .CurrentY = CLng(GetFromINI("PrintDO", "CustY", "", pathFileSettings))
            Printer.Print printRS("Name")
            .CurrentX = tmpX
            Printer.Print printRS("Address")
            .CurrentX = tmpX
            Printer.Print printRS("Zip") & " " & printRS("City")
            .CurrentX = tmpX
            Printer.Print printRS("State") & ", " & printRS("Country")
            'Print delivery order details
            tmpX = CLng(GetFromINI("PrintDO", "OrderMargin", "", pathFileSettings))
            .CurrentX = tmpX
            .CurrentY = CLng(GetFromINI("PrintDO", "OrderY", "", pathFileSettings))
            Printer.Print printRS("DOnumber")
            .CurrentX = tmpX
            Printer.Print printRS("Date")
            .CurrentX = tmpX
            Printer.Print printRS("PONumber")
            .CurrentX = tmpX
            Printer.Print printRS("Attn")
            .CurrentX = tmpX
            Printer.Print printRS("CustomerID")
            .CurrentX = tmpX
            Printer.Print printRS("CreditTerm") & " days"
            .CurrentX = tmpX
            Printer.Print .Page
            Dim cnt As Byte, strPrint As String
            RSOpen printRS, "SELECT * FROM D_Details WHERE DOnumber='" & strDO & "';", dbOpenSnapshot
            cnt = 1
            .CurrentY = CLng(GetFromINI("PrintDO", "DetailY", "", pathFileSettings))
            While Not printRS.EOF
                'Loop through the items and print every detail
                tmpX = CLng(GetFromINI("PrintDO", "CntMargin", "", pathFileSettings))
                .CurrentX = tmpX
                tmpY = .CurrentY
                Printer.Print CStr(cnt)
                tmpX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
                strPrint = printRS("Description")
                .CurrentX = tmpX
                .CurrentY = tmpY
                Printer.Print strPrint
                tmpX = CLng(GetFromINI("PrintDO", "QtyMargin", "", pathFileSettings))
                strPrint = printRS("Quantity")
                .CurrentX = tmpX - (25 + .TextWidth(strPrint))
                .CurrentY = tmpY
                Printer.Print strPrint
                strPrint = printRS("UnitLabel")
                .CurrentX = tmpX + 25
                .CurrentY = tmpY
                Printer.Print strPrint
                strPrint = Format$(printRS("SalePrice"), "#,##0.00")
                .CurrentX = CLng(GetFromINI("PrintDO", "SaleMargin", "", pathFileSettings)) - .TextWidth(strPrint)
                .CurrentY = tmpY
                Printer.Print strPrint
                strPrint = Format$(CSng(printRS("SalePrice")) * CInt(printRS("Quantity")), "#,##0.00")
                grandtotal = grandtotal + CSng(strPrint)
                .CurrentX = CLng(GetFromINI("PrintDO", "SubMargin", "", pathFileSettings)) - .TextWidth(strPrint)
                .CurrentY = tmpY
                Printer.Print strPrint
                Printer.Print
                printRS.MoveNext
                cnt = cnt + 1
            Wend
            'Print charges and total
            tmpY = .ScaleHeight - CLng(GetFromINI("PrintDO", "TermsY", "", pathFileSettings))
            
            If printRS("Charges") > 0 Then
                tmpX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
                .CurrentY = tmpY
                .CurrentX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
                Printer.Print printRS("Description")
                .CurrentY = tmpY
                .CurrentX = CLng(GetFromINI("PrintDO", "SubMargin", "", pathFileSettings))
                grandtotal = grandtotal + CSng(printRS("charges"))
                Printer.Print Format$(printRS("Charges"), "#,##0.00")
                tmpY = .CurrentY
            End If
            'Prints the total line
            .CurrentY = tmpY
            .CurrentX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
            Printer.Print "Grand Total:"
            strPrint = Format$(grandtotal, "#,##0.00")
            .CurrentY = tmpY
            .CurrentX = CLng(GetFromINI("PrintDO", "SubMargin", "", pathFileSettings)) - .TextWidth(strPrint)
            Printer.Print strPrint
            Printer.Print
            Printer.Print
            'Prints the total in string
            tmpX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
            .CurrentX = tmpX
            Printer.Print "RINGGIT MALAYSIA " & UCase(NumToString(CCur(grandtotal))) & " ONLY"
            .EndDoc
            InfoMsg "The delivery order has been successfully sent to the printer for printing.", "Document ready"
        End With
    Else
        ValidMsg "An error occured while the system tried to load the delivery order. Please try again.", "Error occured"
    End If
End If
End Sub

Public Sub printDO(ByVal strDO As String)
'Delivery order is designed based on requirement
'Uncommented codes when printing on non-preprinted papers.
Dim printRS As Recordset, tmpX As Long, tmpY As Long, grandtotal As Single
cdlg.DialogTitle = "Print Delivery Order"
cdlg.CancelError = True
On Error Resume Next
cdlg.ShowPrinter
If Err Then
    Exit Sub
Else
    RSOpen printRS, "SELECT Delivery.DOnumber, Delivery.Date, Delivery.IssuedBy, Delivery.EmployeeID, Delivery.PONumber, Delivery.Remark, Delivery.Attn, Delivery.Description, Delivery.Charges, Delivery.DelTime, Delivery.DelDate, Customers.Name, Customers.Address, Customers.City, Customers.State, Customers.Zip, Customers.Country, Customers.CreditTerm " & _
    "FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE Delivery.DOnumber='" & strDO & "';", dbOpenSnapshot
    If Not printRS.EOF Then
        With Printer
            .PaperSize = 9 'A4 size
            .FontSize = CLng(GetFromINI("PrintDO", "FontSize", "", pathFileSettings))
            'Begin printing receiver details
            grandtotal = 0
            tmpX = CLng(GetFromINI("PrintDO", "CustMargin", "", pathFileSettings))
            tmpY = CLng(GetFromINI("PrintDO", "CustY", "", pathFileSettings))
            .CurrentX = tmpX
            .CurrentY = tmpY
            Printer.Print "Delivery To:"
            .CurrentX = tmpX
            Printer.Print printRS("Name")
            .CurrentX = tmpX
            Printer.Print printRS("Address")
            .CurrentX = tmpX
            Printer.Print printRS("Zip") & " " & printRS("City")
            .CurrentX = tmpX
            Printer.Print printRS("State") & ", " & printRS("Country")
            'Print delivery order details
            Dim strLen As Single
            tmpX = CLng(GetFromINI("PrintDO", "OrderMargin", "", pathFileSettings))
            strLen = .TextWidth("PO Number: ")
            tmpX = tmpX - strLen
            .CurrentX = tmpX
            .CurrentY = tmpY
            Printer.Print "DO Number:"
            .CurrentX = tmpX
            Printer.Print "Date:"
            .CurrentX = tmpX
            Printer.Print "PO Number:"
            .CurrentX = tmpX
            Printer.Print "Reference:"
            .CurrentX = tmpX
            Printer.Print "A/C No:"
            .CurrentX = tmpX
            Printer.Print "Terms"
            .CurrentX = tmpX
            Printer.Print "Page No:"
            tmpX = tmpX + strLen
            .CurrentX = tmpX
            .CurrentY = tmpY
            Printer.Print printRS("DOnumber")
            .CurrentX = tmpX
            Printer.Print printRS("Date")
            .CurrentX = tmpX
            Printer.Print printRS("PONumber")
            .CurrentX = tmpX
            Printer.Print printRS("Attn")
            .CurrentX = tmpX
            Printer.Print printRS("CustomerID")
             .CurrentX = tmpX
            Printer.Print printRS("CreditTerm") & " days"
            .CurrentX = tmpX
            Printer.Print .Page
           
            Dim cnt As Byte, strPrint As String
            RSOpen printRS, "SELECT * FROM D_Details WHERE DOnumber='" & strDO & "';", dbOpenSnapshot
            cnt = 1
            .CurrentY = CLng(GetFromINI("PrintDO", "DetailY", "", pathFileSettings))
            While Not printRS.EOF
                'Loop through the items and print every detail
                tmpX = CLng(GetFromINI("PrintDO", "CntMargin", "", pathFileSettings))
                .CurrentX = tmpX
                tmpY = .CurrentY
                Printer.Print CStr(cnt)
                tmpX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
                strPrint = printRS("Description")
                .CurrentX = tmpX
                .CurrentY = tmpY
                Printer.Print strPrint
                tmpX = CLng(GetFromINI("PrintDO", "QtyMargin", "", pathFileSettings))
                strPrint = printRS("Quantity")
                .CurrentX = tmpX - (25 + .TextWidth(strPrint))
                .CurrentY = tmpY
                Printer.Print strPrint
                strPrint = printRS("UnitLabel")
                .CurrentX = tmpX + 25
                .CurrentY = tmpY
                Printer.Print strPrint
                strPrint = Format$(printRS("SalePrice"), "#,##0.00")
                .CurrentX = CLng(GetFromINI("PrintDO", "SaleMargin", "", pathFileSettings)) - .TextWidth(strPrint)
                .CurrentY = tmpY
                Printer.Print strPrint
                strPrint = Format$(CSng(printRS("SalePrice")) * CInt(printRS("Quantity")), "#,##0.00")
                grandtotal = grandtotal + CSng(strPrint)
                .CurrentX = CLng(GetFromINI("PrintDO", "SubMargin", "", pathFileSettings)) - .TextWidth(strPrint)
                .CurrentY = tmpY
                Printer.Print strPrint
                Printer.Print
                printRS.MoveNext
                cnt = cnt + 1
            Wend
            'Print charges and total
            tmpY = .ScaleHeight - CLng(GetFromINI("PrintDO", "TermsY", "", pathFileSettings))
            
            If printRS("Charges") > 0 Then
                tmpX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
                .CurrentY = tmpY
                .CurrentX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
                Printer.Print printRS("Description")
                strPrint = Format$(printRS("Charges"), "#,##0.00")
                .CurrentY = tmpY
                .CurrentX = CLng(GetFromINI("PrintDO", "SubMargin", "", pathFileSettings)) - .TextWidth(strPrint)
                grandtotal = grandtotal + CSng(printRS("charges"))
                Printer.Print strPrint
                tmpY = .CurrentY
            End If
            'Prints the total line
            .CurrentY = tmpY
            .CurrentX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
            Printer.Print "Grand Total:"
            strPrint = Format$(grandtotal, "#,##0.00")
            .CurrentY = tmpY
            .CurrentX = CLng(GetFromINI("PrintDO", "SubMargin", "", pathFileSettings)) - .TextWidth(strPrint)
            Printer.Print strPrint
            'Prints the total in string
            tmpX = CLng(GetFromINI("PrintDO", "DescMargin", "", pathFileSettings))
            .CurrentX = tmpX
            Printer.Print "RINGGIT MALAYSIA " & UCase(NumToString(CCur(grandtotal))) & " ONLY"
            'Signature and terms & conditions
            strPrint = "Terms and conditions: "
            tmpY = .ScaleHeight - CLng(GetFromINI("PrintDO", "TermsY", "", pathFileSettings))
            tmpX = CLng(GetFromINI("PrintDO", "TermsMargin", "", pathFileSettings))
            .CurrentX = tmpX
            .CurrentY = tmpY
            Printer.Print strPrint
            strPrint = "Issued By:"
            .CurrentX = CLng(GetFromINI("PrintDO", "SignMargin", "", pathFileSettings))
            tmpY = .ScaleHeight - CLng(GetFromINI("PrintDO", "SignY", "", pathFileSettings))
            .CurrentY = tmpY
            Printer.Print strPrint
            Printer.Line (tmpX + .TextWidth(strPrint) + 50, .CurrentY)-Step(3500, 0)
            strPrint = "Received By:"
            tmpX = .ScaleWidth / 2
            .CurrentX = tmpX
            .CurrentY = tmpY
            Printer.Print strPrint
            Printer.Line (tmpX + .TextWidth(strPrint) + 50, .CurrentY)-Step(3500, 0)
            .EndDoc
            InfoMsg "The delivery order has been successfully sent to the printer for printing.", "Document ready"
        End With
    Else
        ValidMsg "An error occured while the system tried to load the delivery order. Please try again.", "Error occured"
    End If
End If
End Sub

Private Sub charges_GotFocus()
SelText charges
End Sub

Private Sub charges_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub charges_LostFocus()
If charges.Text = "" Then
    charges.Text = "0"
End If
charges.Text = Format$(charges.Text, "#,##0.00")
End Sub

Private Sub cmdCancel_Click()
Dim k As Integer
doNum.Text = doNum.Tag
For k = 0 To 2
    cmbDate(k).Text = cmbDate(k).Tag
    cmbDelDate(k).Text = cmbDelDate(k).Tag
    If k < 2 Then
        cmbDelTime(k).Text = cmbDelTime(k).Tag
    End If
Next k
employee(0).Text = lblEnd.Caption
employee(1).Text = lblDel.Caption

notes.Text = notes.Tag
charges.Text = charges.Tag
customer.Text = lblcust.Caption
txtPO.Text = txtPO.Tag
txtAttn.Text = txtAttn.Tag
txtREM.Text = txtREM.Tag
setFormMode Viewing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If lvDO.ListItems.Count < 1 Then
    InfoMsg "There are no deliveries to be deleted.", "No delivery available"
Else
    If doNum.Text <> "" Then
        If MsgBox("Are you sure you want to delete the selected delivery order?", vbYesNoCancel + vbQuestion, "Delete delivery order") = vbYes Then
            MySynonDatabase.Execute "DELETE * FROM Delivery WHERE DONumber='" & lvDO.SelectedItem.Text & "';"
            'Insert into systems log
            insertLog "DO number: " & lvDO.SelectedItem.Text & " has been deleted."
            InfoMsg "The delivery order has been successfully deleted.", "Record deleted"
            Call cmdFilter_Click
            doNum.Text = ""
            notes.Text = ""
            charges.Text = ""
            Dim k As Byte
            For k = 0 To 2
                cmbDate(k).ListIndex = -1
                cmbDelDate(k).ListIndex = -1
                If k < 2 Then
                    cmbDelTime(k).ListIndex = -1
                    employee(k).ListIndex = -1
                End If
            Next k
            customer.ListIndex = -1
            lblDel.Caption = ""
            lblEnd.Caption = ""
            lblcust.Caption = ""
            txtPO.Text = ""
            txtAttn.Text = ""
            txtREM.Text = ""
            frmMain.loadRecDeliveries
        End If
    Else
        InfoMsg "Please select a delivery order to be deleted.", "Missing selection"
    End If
End If
End Sub

Private Sub cmdEdit_Click()
If lvDO.ListItems.Count > 0 Then
    If doNum.Text <> "" Then
        Dim k As Integer
        doNum.Tag = doNum.Text
        notes.Tag = notes.Text
        charges.Tag = charges.Text
        For k = 0 To 2
            cmbDate(k).Tag = cmbDate(k).Text
            cmbDelDate(k).Tag = cmbDelDate(k).Text
            If k < 2 Then
                cmbDelTime(k).Tag = cmbDelTime(k).Text
            End If
        Next k
        
        lblDel.Caption = employee(1).Text
        lblEnd.Caption = employee(0).Text
        lblcust.Caption = customer.Text
        txtPO.Tag = txtPO.Text
        txtAttn.Tag = txtAttn.Text
        txtREM.Tag = txtREM.Text
        setFormMode Editing
    Else
        InfoMsg "Please select a deliver order first.", "No delivery order selected"
    End If
Else
    InfoMsg "There are no delivery orders available.", "No delivery order"
End If
End Sub

Private Sub cmdFilter_Click()
If (filter(0).Text <> "") And (filter(1).Text <> "") And (filter(2).Text <> "") Then
    If isDateValid(CByte(filter(0).Text), CByte(filter(1).Text), CInt(filter(2).Text)) = True Then
        getDeliveries "SELECT Delivery.DOnumber, Customers.Name, Delivery.PONumber, Delivery.DelDate, Delivery.DelTime, Delivery.Status FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE (((Delivery.Date)='" & filter(0).Text & "/" & filter(1).Text & "/" & filter(2).Text & "') AND ((Delivery.Status) <> 'INVOICED'));"
    Else
        ValidMsg "The selected date is invalid. Please try again.", "Invalid date"
    End If
End If
End Sub

Private Sub cmdNew_Click()
Dim d As frmDelivery
Set d = New frmDelivery
Load d
d.Show , frmMain

End Sub

Private Sub cmdSave_Click()
If doNum.Text = "" Then
    ValidMsg "Please enter a delivery order number.", "Missing DO Number"
    doNum.SetFocus
ElseIf isDateValid(CByte(cmbDate(0).Text), CByte(cmbDate(1).Text), CInt(cmbDate(2).Text)) = False Then
    ValidMsg "Please select a valid date for the delivery order.", "Invalid date"
    cmbDate(0).SetFocus
ElseIf isDateValid(CByte(cmbDelDate(0).Text), CByte(cmbDelDate(1).Text), CInt(cmbDelDate(2).Text)) = False Then
    ValidMsg "Please select a valid delivery date for the delivery order.", "Invalid delivery date"
    cmbDelDate(0).SetFocus
ElseIf ((notes.Text = "") And (Val(charges.Text) > 0)) Then
    ValidMsg "Please provide a description for the charges.", "Missing description"
    notes.SetFocus
Else
    Dim dRS As Recordset
    RSOpen dRS, "SELECT * FROM Delivery WHERE DOnumber='" & doNum.Tag & "';", dbOpenDynaset
    If Not dRS.EOF Then
        dRS.Edit
        dRS("DOnumber") = doNum.Text
        dRS("Date") = cmbDate(0) & "/" & cmbDate(1).Text & "/" & cmbDate(2).Text
        dRS("CustomerID") = customer.Tag
        dRS("EmployeeID") = employee(0).Tag
        dRS("IssuedBy") = employee(1).Tag
        dRS("PONumber") = txtPO.Text
        dRS("Attn") = txtAttn.Text
        dRS("Remark") = txtREM.Text
        dRS("DelDate") = cmbDelDate(0).Text & "/" & cmbDelDate(1).Text & "/" & cmbDelDate(2).Text
        dRS("DelTime") = cmbDelTime(0).Text & ":" & cmbDelTime(1).Text
        dRS("Charges") = charges.Text
        dRS("Description") = notes.Text
        dRS.Update
        
        dRS.Close
        Set dRS = Nothing
        InfoMsg "The Delivery Order has been successfully updated.", "Record saved"
        setFormMode Viewing
        Call cmdFilter_Click
    End If
End If
End Sub


Private Sub customer_Click()
If Not customer.Text = "" Then
    Dim tempRS As Recordset
    RSOpen tempRS, "SELECT CustomerID FROM Customers WHERE Name='" & customer.Text & "'", dbOpenSnapshot
    If Not tempRS.EOF Then
        customer.Tag = tempRS("CustomerID")
    End If
    tempRS.Close
    Set tempRS = Nothing
End If
End Sub

Private Sub doNum_GotFocus()
SelText doNum
End Sub

Private Sub doNum_KeyPress(KeyAscii As Integer)
OnlyNum KeyAscii
End Sub

Private Sub employee_Click(Index As Integer)
If Not employee(Index).Text = "" Then
    Dim tempRS As Recordset
    RSOpen tempRS, "SELECT EmployeeID FROM Employees WHERE Name='" & employee(Index).Text & "'", dbOpenSnapshot
    If Not tempRS.EOF Then
        employee(Index).Tag = tempRS("EmployeeID")
    End If
    tempRS.Close
    Set tempRS = Nothing
End If
End Sub

Private Sub Form_Load()
DisableClose frmDelivery_Main, True
lblNotes.Caption = "Welcome to the delivery order management console. Please be careful in changing the details of these orders. " & vbCrLf & _
"Changes upon these documents may not reflect the truth in reality thus may cause undesirable outcomes and fatal errors. Ensure that you are " & _
"fully aware of what you are doing."
Dim i As Integer
'Add years
For i = 0 To 10
    cmbDate(2).addItem CStr(Year(Now()) - 5 + i)
    cmbDelDate(2).addItem CStr(Year(Now()) - 5 + i)
    filter(2).addItem CStr(Year(Now()) - 5 + i)
Next i
For i = 0 To 59
    cmbDelTime(1).addItem Format$(i, "00")
Next i

'Populate combo boxes
FillCombo customer, "SELECT Name FROM Customers", "Name"
FillCombo employee(0), "SELECT Name FROM Employees ORDER BY Name ASC", "Name"
FillCombo employee(1), "SELECT Name FROM Employees ORDER BY Name ASC", "Name"

filter(0).Text = Format$(Day(Now()), "00")
filter(1).Text = Format$(Month(Now()), "00")
filter(2).Text = Format$(Year(Now()), "0000")

'Format the list view properties
With lvDO.ColumnHeaders
    lvDO.View = lvwReport
    .Clear
    .add , , "DO No.", 800
    .add , , "Customer", 5000
    .add , , "PO Number"
    .add , , "Delivery Date"
    .add , , "Delivery Time"
    .add , , "Status"
End With
Call cmdFilter_Click
setFormMode Viewing
'Me.WindowState = vbMaximized

End Sub

Private Sub Form_Resize()
On Error Resume Next
Shape1.width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDelivery_Main = Nothing
End Sub

Private Sub lvDO_DblClick()
With lvDO.SelectedItem
    If lvDO.ListItems.Count > 0 Then
        If .Selected Then
            Load frmDelivery_Details
            frmDelivery_Details.getDetails .Text
            frmDelivery_Details.Show vbModal
            frmDelivery_Details.Left = Me.Left + Me.width
        End If
    End If
End With
End Sub

Private Sub lvDO_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item 'If a delivery order is selected, get its details from the DB and display
    If .Selected Then
        Dim lRS As Recordset, eRS As Recordset
        'On Error GoTo ErrHandler
        RSOpen lRS, "SELECT Delivery.DOnumber, Delivery.Date, Delivery.EmployeeID, Delivery.IssuedBy, Delivery.PONumber, Customers.Name, Delivery.Remark, Delivery.Attn, Delivery.DelDate, Delivery.DelTime, Delivery.Status, Delivery.Charges, Delivery.Description " & _
                    "FROM Customers INNER JOIN Delivery ON Customers.CustomerID = Delivery.CustomerID WHERE Delivery.DOnumber='" & .Text & "';", dbOpenSnapshot
        
        If Not lRS.EOF Then
            doNum.Text = lRS("DOnumber")
            customer.Text = lRS("Name")
            RSOpen eRS, "SELECT Name FROM Employees WHERE EmployeeID='" & lRS("EmployeeID") & "';", dbOpenSnapshot
            employee(0).Text = eRS("Name")
            RSOpen eRS, "SELECT Name FROM Employees WHERE EmployeeID='" & lRS("IssuedBy") & "';", dbOpenSnapshot
            employee(1).Text = eRS("Name")
            cmbDate(0).Text = Left$(lRS("Date"), 2)
            cmbDate(1).Text = Right$(Left$(lRS("Date"), 5), 2)
            cmbDate(2).Text = Right$(lRS("Date"), 4)
            cmbDelDate(0).Text = Left$(lRS("DelDate"), 2)
            cmbDelDate(1).Text = Right$(Left$(lRS("DelDate"), 5), 2)
            cmbDelDate(2).Text = Right$(lRS("DelDate"), 4)
            cmbDelTime(0).Text = Left$(lRS("DelTime"), 2)
            cmbDelTime(1).Text = Right$(lRS("DelTime"), 2)
            txtAttn.Text = IIf(IsNull(lRS("Attn")), "", lRS("Attn"))
            txtPO.Text = IIf(IsNull(lRS("PONumber")), "", lRS("PONumber"))
            txtREM.Text = IIf(IsNull(lRS("Remark")), "", lRS("Remark"))
        End If
        eRS.Close
        lRS.Close
        Set eRS = Nothing
        Set lRS = Nothing
    End If
End With
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Unable to load the delivery orders. Please close this window and try again." & vbCrLf & _
    "If you see this message again, please contact your system administrator.", "Error found"
    Exit Sub
End If
End Sub

Private Sub lvDO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    If lvDO.ListItems.Count > 0 Then
        If doNum.Text <> "" Then
            PopupMenu mnu_Options, vbPopupMenuLeftAlign
        End If
    End If
End If
End Sub

Private Sub mnu_Print_Blank_Click()
If ((lvDO.SelectedItem.Selected = True) And (doNum.Text <> "")) Then
    If MsgBox("Are you sure you want to print the following Delivery Order?" & vbCrLf & "DO Number: " & doNum.Text, vbQuestion + vbYesNoCancel, "Print DO") = vbYes Then
        printDO lvDO.SelectedItem.Text
    End If
End If

End Sub

Private Sub mnu_Print_pre_Click()
If ((lvDO.SelectedItem.Selected = True) And (doNum.Text <> "")) Then
    If MsgBox("Are you sure you want to print the following Delivery Order?" & vbCrLf & "DO Number: " & doNum.Text, vbQuestion + vbYesNoCancel, "Print DO") = vbYes Then
        prePrintDO lvDO.SelectedItem.Text
    End If
End If
End Sub

Private Sub txtAttn_GotFocus()
SelText txtAttn
End Sub

Private Sub txtPO_GotFocus()
SelText txtPO
End Sub

Private Sub txtREM_GotFocus()
SelText txtREM
End Sub
