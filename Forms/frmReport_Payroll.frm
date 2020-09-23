VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport_Payroll 
   Caption         =   "Listing of Payroll"
   ClientHeight    =   5505
   ClientLeft      =   -120
   ClientTop       =   450
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
   Icon            =   "frmReport_Payroll.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdResult 
      Caption         =   "&Get Report..."
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
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbyear 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvPayroll 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Selected year:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Selected month:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport_Payroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdResult_Click()
If (cmbyear.Text <> "") And (cmbMonth.Text <> "") Then
    getPayrolls Format$(cmbMonth.ListIndex + 1, "00"), cmbyear.Text
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 12
    cmbMonth.addItem MonthName(i, False), i - 1
    cmbyear.addItem Format$(Year(Now()) + 11 - i, "0000"), i - 1
Next i
With lvPayroll
    .View = lvwReport
    .FullRowSelect = True
    .BorderStyle = ccNone
    .LabelEdit = lvwManual
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.add , , "Name", 4000
    .ColumnHeaders.add , , "Date Issued"
    .ColumnHeaders.add , , "Initial Salary", 0
    .ColumnHeaders.add , , "Unpaid Leaves", 0
    .ColumnHeaders.add , , "Gross Salary"
    .ColumnHeaders.add , , "EPF Contribution"
    .ColumnHeaders.add , , "EPF Employer", 0
    .ColumnHeaders.add , , "SOCSO"
    .ColumnHeaders.add , , "Income Tax"
    .ColumnHeaders.add , , "Net Salary"
    .ColumnHeaders.add , , "Salary Advance"
    .ColumnHeaders.add , , "Balance"
    .ColumnHeaders.add , , "OT Allowance"
    .ColumnHeaders.add , , "Balance(Bank)"
    
    '.ColumnHeaders.add , , "Gross Salary"
End With
End Sub

Private Sub Form_Resize()
lvPayroll.width = Me.ScaleWidth - (lvPayroll.Left * 2)
lvPayroll.height = Me.ScaleHeight - (lvPayroll.Top + 115)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmReport_Payroll = Nothing
End Sub

Private Sub getPayrolls(ByVal strMonth As String, ByVal strYear As String)
Dim reportRS As Recordset, tmpSQL As String
Dim totalInit, totalUnpaid, totalGross, sumEPF_work, sumEPF_boss, totalSocso, totalInctax, totalNet, totalAdv, totalBal, totalOT, totalBank As Single
'Initialise variables
totalOT = 0
totalInit = 0
sumEPF_work = 0
sumEPF_boss = 0
totalSocso = 0
totalInctax = 0
totalNet = 0
totalGross = 0
totalUnpaid = 0
totalAdv = 0
totalBal = 0
totalBank = 0

'Setup query
tmpSQL = "SELECT Payroll.*, Employees.Name " & _
"FROM Employees INNER JOIN Payroll ON Employees.EmployeeID = Payroll.EmployeeID WHERE (((Payroll.dateIssued) Like '##/" & strMonth & "/" & strYear & "'));"
RSOpen reportRS, tmpSQL, dbOpenSnapshot

With lvPayroll.ListItems
    .Clear
    While Not reportRS.EOF
        .add , , reportRS("Name")
        .Item(.Count).SubItems(1) = reportRS("dateIssued")
        .Item(.Count).SubItems(2) = Format$(reportRS("initialSalary"), "#,##0.00")
        totalInit = totalInit + reportRS("initialSalary")
        .Item(.Count).SubItems(3) = reportRS("unpaidLeaves") * reportRS("initialSalary") / reportRS("workingDays")
        totalUnpaid = totalUnpaid + CSng(.Item(.Count).SubItems(3))
        .Item(.Count).SubItems(4) = CSng(.Item(.Count).SubItems(2)) - CSng(.Item(.Count).SubItems(3))
        totalGross = totalGross + CSng(.Item(.Count).SubItems(2)) - CSng(.Item(.Count).SubItems(3))
        .Item(.Count).SubItems(5) = Format$(reportRS("epfEmployee"), "#,##0.00")
        sumEPF_work = sumEPF_work + reportRS("epfEmployee")
        .Item(.Count).SubItems(6) = Format$(reportRS("epfEmployer"), "#,##0.00")
        sumEPF_boss = sumEPF_boss + reportRS("epfEmployer")
        .Item(.Count).SubItems(7) = Format$(reportRS("socsoAmount"), "#,##0.00")
        totalSocso = totalSocso + reportRS("socsoAmount")
        .Item(.Count).SubItems(8) = Format$(reportRS("incomeTax"), "#,##0.00")
        totalInctax = totalInctax + reportRS("incomeTax")
        .Item(.Count).SubItems(9) = Format$(CSng(.Item(.Count).SubItems(4)) - (CSng(.Item(.Count).SubItems(5)) + CSng(.Item(.Count).SubItems(7)) + CSng(.Item(.Count).SubItems(8))), "#,##0.00")
        totalNet = totalNet + CSng(.Item(.Count).SubItems(9))
        .Item(.Count).SubItems(10) = Format$(reportRS("salaryAdvance"), "#,##0.00")
        totalAdv = totalAdv + reportRS("salaryAdvance")
        
        .Item(.Count).SubItems(11) = Format$(CSng(.Item(.Count).SubItems(9)) - CSng(.Item(.Count).SubItems(10)), "#,##0.00")
        totalBal = totalBal + CSng(.Item(.Count).SubItems(11))
        .Item(.Count).SubItems(12) = Format$(reportRS("otAmount"), "#,##0.00")
        totalOT = totalOT + reportRS("otAmount")
        .Item(.Count).SubItems(13) = CSng(.Item(.Count).SubItems(11)) - CSng(.Item(.Count).SubItems(12))
        totalBank = totalBank + CSng(.Item(.Count).SubItems(13))
        '.Item(.Count).SubItems (13)
        reportRS.MoveNext
    Wend
    'List the total of rows
    .add , , "Total"
    .Item(.Count).SubItems(2) = Format$(totalInit, "#,##0.00")
    .Item(.Count).SubItems(3) = Format$(totalUnpaid, "#,##0.00")
    .Item(.Count).SubItems(4) = Format$(totalGross, "#,##0.00")
    .Item(.Count).SubItems(5) = Format$(sumEPF_work, "#,##0.00")
    .Item(.Count).SubItems(6) = Format$(sumEPF_boss, "#,##0.00")
    .Item(.Count).SubItems(7) = Format$(totalSocso, "#,##0.00")
    .Item(.Count).SubItems(8) = Format$(totalInctax, "#,##0.00")
    .Item(.Count).SubItems(9) = Format$(totalNet, "#,##0.00")
    .Item(.Count).SubItems(10) = Format$(totalAdv, "#,##0.00")
    .Item(.Count).SubItems(11) = Format$(totalBal, "#,##0.00")
    .Item(.Count).SubItems(12) = Format$(totalOT, "#,##0.00")
    .Item(.Count).SubItems(13) = Format$(totalBank, "#,##0.00")
End With
reportRS.Close
Set reportRS = Nothing
End Sub
