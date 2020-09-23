VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayroll_Print 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Payroll"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
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
   ScaleHeight     =   4545
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   3360
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stBar2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4290
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker datepk 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MM, yyyy"
      Format          =   50987011
      CurrentDate     =   38157
   End
   Begin VB.Frame Frame1 
      Caption         =   "Payrolls for the month:"
      Enabled         =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
      Begin MSComctlLib.ListView lvEmp 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4260
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
   End
   Begin VB.OptionButton OnlySelected 
      Caption         =   "Print only selected employees."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.OptionButton AllPayroll 
      Caption         =   "Print payroll for all employees."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2775
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
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Month:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPayroll_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AllPayroll_Click()
Frame1.Enabled = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
If lvEmp.ListItems.Count = 0 Then
    ValidMsg "No records available to print. Please try selecting another month.", "No records available"
ElseIf (OnlySelected.Value = True) And (lvEmp.SelectedItem.Selected = False) Then
    ValidMsg "Please select at least a employee payroll record to be printed.", "No record selected"
Else
    'begin print
    With cDlg
        On Error Resume Next
        .CancelError = True
        .ShowPrinter
        If Err Then
            Exit Sub
        Else
            
            If AllPayroll.Value = True Then
                getPrintPayroll "SELECT Payroll.*, Employees.Name, Employees.PositionID, Employees.IC " & _
                            "FROM Employees INNER JOIN Payroll ON Employees.EmployeeID = Payroll.EmployeeID;"
            Else
                Dim i As Integer
                For i = 1 To lvEmp.ListItems.Count
                    If lvEmp.ListItems(i).Selected Then
                        getPrintPayroll "SELECT Payroll.*, Employees.Name, Employees.PositionID, Employees.IC " & _
                                    "FROM Employees INNER JOIN Payroll ON Employees.EmployeeID = Payroll.EmployeeID WHERE Payroll.payrollID=" & lvEmp.ListItems(i).Tag & ";"
                    End If
                Next i
            End If
        End If
    End With
End If
End Sub

Private Sub datepk_Click()
getPayroll Format$(Month(datepk.Value), "00"), Format$(Year(datepk.Value), "0000")
End Sub

Private Sub Form_Load()
With lvEmp
    .View = lvwReport
    .MultiSelect = True
    .ColumnHeaders.add , , "ID", 0
    .ColumnHeaders.add , , "Name", 4000
End With
AllPayroll.Value = True
stBar2.Panels(stBar2.Panels.Count).width = stBar2.width
Call datepk_Click
End Sub
Private Sub getPayroll(ByVal strMonth As String, strYear As String)
Dim payRS As Recordset
lvEmp.ListItems.Clear
RSOpen payRS, "SELECT Payroll.payrollID, Employees.EmployeeID, Employees.Name, Payroll.dateIssued " & _
"FROM Employees INNER JOIN Payroll ON Employees.EmployeeID = Payroll.EmployeeID WHERE Payroll.dateIssued LIKE '##/" & strMonth & "/" & strYear & "';", dbOpenSnapshot
While Not payRS.EOF
    lvEmp.ListItems.add , , payRS("EmployeeID")
    lvEmp.ListItems(lvEmp.ListItems.Count).Tag = payRS("payrollID")
    lvEmp.ListItems(lvEmp.ListItems.Count).SubItems(1) = payRS("Name")
    payRS.MoveNext
Wend
payRS.Close
Set payRS = Nothing
End Sub

Private Sub OnlySelected_Click()
Frame1.Enabled = True
End Sub

Private Sub getPrintPayroll(ByVal strParamSQL As String)
'Payroll is designed according to the requirement
Dim net_sal As Single, gross_sal As Single, strText As String
Dim tmpX As Long, tmpY As Long
Dim width As Long, height As Long, coordY As Long, coordX As Long
'On Error GoTo ErrHandler

'Debug.Print strParamSQL
With Printer
'    .Orientation = vbPRORLandscape
    Dim printRS As Recordset
    '.Orientation = vbPRORLandscape
    .FontSize = CByte(GetFromINI("PrintDO", "FontSize", "", pathFileSettings))
    .PaperSize = 11
    'Add a document to the application
    RSOpen printRS, strParamSQL, dbOpenSnapshot
    While Not printRS.EOF
        net_sal = 0
        displayMsg "Preparing document for printing."
        
        displayMsg "Writing data to the document"
        'Company and payroll details
        strText = getSettings("compName")
        .CurrentX = .ScaleWidth / 2 - .TextWidth(strText) / 2
        Printer.Print strText
        Printer.Print
        strText = "Employee Payroll Slip"
        .CurrentX = .ScaleWidth / 2 - .TextWidth(strText) / 2
        Printer.Print strText
        Printer.Print
        strText = "Month: " & MonthName(datepk.Month, False) & " " & CStr(datepk.Year)
        .CurrentX = .ScaleWidth / 2 - .TextWidth(strText) / 2
        Printer.Print strText
        Printer.Print
        'Headings on the left hand side include employee details, performance, etc.
        tmpX = CLng(GetFromINI("PrtPayroll", "HeaderLeft", "", pathFileSettings))
        tmpY = .CurrentY
        strText = "Name: "
        .CurrentX = tmpX
        Printer.Print strText
        strText = "IC No: "
        .CurrentX = tmpX
        Printer.Print strText
        strText = "Position: "
        .CurrentX = tmpX
        Printer.Print strText
        'Format$(printRS("payrollID"), "000000")
        strText = "Total working days: "
        .CurrentX = tmpX
        Printer.Print strText
        'printRS ("EmployeeID")
        'Monthly performance details
        strText = "Total attendants: "
        .CurrentX = tmpX
        Printer.Print strText
        'printRS ("otHours")
        strText = "Paid/annual leaves: "
        .CurrentX = tmpX
        Printer.Print strText
        strText = "Sick leaves: "
        .CurrentX = tmpX
        Printer.Print strText
        strText = "Others: "
        .CurrentX = tmpX
        Printer.Print strText
        
        'Print details for the left hand side
        .CurrentY = tmpY
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("Name")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("IC")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("PositionID")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("workingDays")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("dayWorked")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("annualLeaves")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("sickLeaves")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        tmpX = .ScaleWidth / 2 - 100
        strText = printRS("otherLeaves")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        
        'headings on the right hand side
        .CurrentY = tmpY
        .CurrentX = .ScaleWidth / 2
        tmpX = .CurrentX
        strText = "Gross Salary: "
        Printer.Print strText
        .CurrentX = tmpX
        strText = "Less: unpaid leaves"
        Printer.Print strText
        .CurrentX = tmpX
        strText = "Salary: "
        Printer.Print strText
        .CurrentX = tmpX
        strText = "Less: EPF"
        Printer.Print strText
        .CurrentX = tmpX + .TextWidth("Less: ")
        strText = "SOCSO"
        Printer.Print strText
        .CurrentX = tmpX + .TextWidth("Less: ")
        strText = "Income Tax"
        Printer.Print strText
        .CurrentX = tmpX + .TextWidth("Less: ")
        strText = "Salary Advanced"
        Printer.Print strText
        .CurrentX = tmpX
        strText = "Balance: "
        Printer.Print strText
        .CurrentX = tmpX
        strText = "Add: Incentive"
        Printer.Print strText
        .CurrentX = tmpX
        strText = "Net Salary: "
        Printer.Print strText
        
        'Deductions
        tmpX = .ScaleWidth / 2 + .ScaleWidth / 4 + 1000
        .CurrentY = tmpY
        strText = Format$(printRS("initialSalary"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        gross_sal = (printRS("unpaidLeaves") * printRS("initialSalary") / printRS("workingDays"))
        strText = Format$(gross_sal, "#,##0.00")
        gross_sal = printRS("initialSalary") - gross_sal
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        strText = Format$(printRS("initialSalary"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        strText = Format$(printRS("epfEmployee"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        strText = Format$(printRS("socsoAmount"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        strText = Format$(printRS("incomeTax"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        strText = Format$(printRS("salaryAdvance"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        gross_sal = gross_sal - (printRS("epfEmployee") + printRS("socsoAmount") + printRS("incomeTax") + printRS("salaryAdvance"))
        strText = Format$(gross_sal, "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        strText = Format$(printRS("otAmount"), "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        net_sal = gross_sal + printRS("otAmount")
        strText = Format$(net_sal, "#,##0.00")
        .CurrentX = tmpX - .TextWidth(strText)
        Printer.Print strText
        
        'Signature
        tmpX = CLng(GetFromINI("PrtPayroll", "HeaderLeft", "", pathFileSettings))
        tmpY = CLng(GetFromINI("PrtPayroll", "PaymentLeft", "", pathFileSettings))
        .CurrentX = tmpX
        .CurrentY = tmpY
        strText = "Payment method: " & printRS("paymentMethod")
        Printer.Print strText
        .CurrentX = tmpX
        tmpY = CLng(GetFromINI("PrtPayroll", "SignY", "", pathFileSettings))
        .CurrentY = tmpY
        strText = "Issued by:"
        Printer.Print strText
        .CurrentX = .ScaleWidth / 2
        .CurrentY = tmpY
        strText = "Received by:"
        Printer.Print strText
        
        'Draw the box
        width = CLng(GetFromINI("PrtPayroll", "BoxW", "", pathFileSettings))
        height = CLng(GetFromINI("PrtPayroll", "BoxH", "", pathFileSettings))
        coordX = CLng(GetFromINI("PrtPayroll", "BoxX", "", pathFileSettings))
        coordY = CLng(GetFromINI("PrtPayroll", "BoxY", "", pathFileSettings))
        Printer.Line (coordX, coordY)-(width, height), vbBlack, B
        Printer.Line ((.ScaleWidth / 2) - 25, 1050)-((.ScaleWidth / 2) - 25, height), vbBlack
        .EndDoc
        printRS.MoveNext
        displayMsg "Document saved and ready for printing"
    Wend
    printRS.Close
    Set printRS = Nothing
End With
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "Unable to print the payroll due to a critical error. Please close the application and try again.", "Error"
    Exit Sub
End If
End Sub

Private Sub displayMsg(ByVal strMsg As String)
stBar2.Panels(stBar2.Panels.Count).Text = strMsg & "..."
End Sub

