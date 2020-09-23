VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayroll_New 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Payroll"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate Payroll..."
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
      Left            =   8400
      TabIndex        =   29
      ToolTipText     =   "Click here to automate calculations"
      Top             =   6840
      Width           =   1695
   End
   Begin VB.OptionButton pay1 
      Caption         =   "Bank"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.OptionButton pay2 
      Caption         =   "Cash"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.OptionButton pay3 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
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
      Left            =   10200
      TabIndex        =   30
      ToolTipText     =   "Click here to save the current payroll for the selected employee."
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   8400
      ScaleHeight     =   4335
      ScaleWidth      =   3015
      TabIndex        =   17
      Top             =   2400
      Width           =   3015
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   10
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   8
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   7
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   6
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   0
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   3
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   4
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         Index           =   5
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Advance:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "OT Allowance:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Income tax:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Salary:"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Unpaid Leaves"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "EPF:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "SOCSO:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Salary:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Salary:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Selected Employee:"
      Height          =   1695
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   5055
      Begin VB.Label lblSalary 
         Height          =   255
         Left            =   3360
         TabIndex        =   56
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Salary:"
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblChild 
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "No. of Children:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblMaritial 
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Maritial Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblName 
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label16 
         Caption         =   "Employee ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker datepk 
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   503
      _Version        =   393216
      Format          =   19529728
      CurrentDate     =   38152
      MinDate         =   31048
   End
   Begin VB.Frame Frame3 
      Caption         =   "Payment Method:"
      Height          =   1815
      Left            =   5280
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
      Begin VB.TextBox txtCheque 
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Cheque Ref:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Required Details:"
      Height          =   2775
      Left            =   5280
      TabIndex        =   7
      Top             =   3600
      Width           =   3015
      Begin VB.TextBox txtUnpaid 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtRate 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "0.00"
         ToolTipText     =   "This rate is taken from the salary divided by the number of days for the month."
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtHours 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtOTHours 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtAnnual 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSick 
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Unpaid Leave(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Estimated daily rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Total Days Worked:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "OT hour(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Annual/Paid Leave(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Sick Leave(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lvEmployees 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7223
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
   Begin VB.Frame Frame1 
      Caption         =   "EPF Contribution:"
      Height          =   975
      Left            =   8400
      TabIndex        =   14
      Top             =   1320
      Width           =   3015
      Begin VB.TextBox txtEmployer 
         Height          =   285
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtEmployee 
         Height          =   285
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Employer:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Employee:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label instructions 
      Height          =   1095
      Left            =   120
      TabIndex        =   57
      Top             =   120
      Width           =   11295
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "&Print..."
      End
      Begin VB.Menu dash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmPayroll_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows users to insert a payroll record under
'an employee's account.
Option Explicit
Dim hour_work, ot_hour, ann_leave, sick_leave, unpaid_leave As Byte

Private Sub cmdCalc_Click()
autoCalc
End Sub

Private Sub cmdSave_Click()
If lblID.Caption = "" Then
    ValidMsg "Please select an employee.", "Missing employee"
    lvEmployees.SetFocus
ElseIf ((Val(txtHours.Text) < 0) Or (Val(txtHours.Text) > 552)) Then
    ValidMsg "It is impossible to work more than 552 hours a month. Please try again.", "Invalid number of hours"
    txtHours.SetFocus
ElseIf (pay1.Value = False) And (pay2.Value = False) And (pay3.Value = False) Then
    ValidMsg "Please select the type of payment for this payroll.", "Missing payment method"
    pay1.SetFocus
ElseIf (pay3.Value = True) And (txtCheque.Text = "") Then
    ValidMsg "Please enter the cheque details.", "Missing cheque details"
    txtCheque.SetFocus
Else
    savePayroll
End If
End Sub

Private Sub Form_Load()
datepk.Value = Now()
getListofEmployees lvEmployees
instructions.Caption = "To create a payroll, select employee from the list." & vbCrLf & _
"Details including leaves taken during the month of the current date would be retrieved." & vbCrLf & _
"Estimated hourly rate is based on the salary divided by the number of days for the month of the selected date." & vbCrLf & _
"Ensure that payment method has been selected and all other fields have been carefully filled with data." & vbCrLf & _
"Feel free to click on 'Calculate Payroll...' to automate the calculations. When you are done, just click on 'Save'."
End Sub

Private Sub getListofEmployees(ByRef strList As ListView)
Dim tempSQL As String
Dim listRS As Recordset
With strList
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.add , , "ID", 0
    .ColumnHeaders.add , , "Name", .width - 100
    'clear existing items
    .ListItems.Clear
    tempSQL = "SELECT Employees.EmployeeID, Employees.Name FROM Employees;"
    On Error GoTo ErrHandler
    RSOpen listRS, tempSQL, dbOpenSnapshot
    While Not listRS.EOF
        .ListItems.add , , listRS("EmployeeID")
        .ListItems(.ListItems.Count).SubItems(1) = listRS("Name")
        listRS.MoveNext
    Wend
End With
ErrHandler:
If Err.Number <> 0 Then
    strList.ListItems.Clear
    strList.ListItems.add , , "UNABLE TO LOAD LIST OF EMPLOYEES."
End If
End Sub

Private Sub getEmployeeInfo(ByVal strEmployeeID As String)
Dim tempSQL As String
Dim tempRS As Recordset
' Gets the salary of the employee
tempSQL = "SELECT Name, Maritial, Children, Salary FROM Employees WHERE EmployeeID='" & strEmployeeID & "';"
On Error GoTo ErrHandler
RSOpen tempRS, tempSQL, dbOpenSnapshot
If Not tempRS.EOF Then
    lblSalary.Caption = Format$(tempRS("Salary"), "#,##0.00")
    lblName.Caption = tempRS("Name")
    lblID.Caption = strEmployeeID
    lblMaritial.Caption = IIf((CBool(tempRS("Maritial")) = True), "Married", "Single")
    lblChild.Caption = tempRS("Children")
End If
'Gets the leaves taken by the employee
Dim thisMonth, thisYear, numAnnual, numSick, numPaid, numUnpaid As Integer
thisMonth = Month(Now())
thisYear = Year(Now())
'Initialise variables
numAnnual = 0
numSick = 0
numPaid = 0
numUnpaid = 0

tempSQL = "SELECT Emp_Leaves.type, (DateDiff('d',Format([beginDate],'dd/mm/yyyy'),Format([endDate],'dd/mm/yyyy'))) AS numDays " & _
"From Emp_Leaves " & _
"WHERE (((Emp_Leaves.date) Like '##/" & Format$(thisMonth, "00") & "/" & Format$(thisYear, "0000") & "') AND ((Emp_Leaves.EmployeeID)='" & strEmployeeID & "'));"
RSOpen tempRS, tempSQL, dbOpenSnapshot
While Not tempRS.EOF
    If tempRS("type") = "ANNUAL" Then
        numAnnual = numAnnual + tempRS("numDays")
    ElseIf tempRS("type") = "SICK" Then
        numSick = numSick + tempRS("numDays")
    ElseIf tempRS("type") = "UNPAID" Then
        numUnpaid = numUnpaid + tempRS("numDays")
    ElseIf tempRS("type") = "PAID" Then
        numPaid = numPaid + tempRS("numDays")
    End If
    tempRS.MoveNext
Wend
txtAnnual.Text = numAnnual
txtSick.Text = numSick

'Calculate estimated hourly rate
txtRate.Text = CSng(lblSalary.Caption) / CByte(getSettings("numDays"))

ErrHandler:
If Err.Number <> 0 Then
    ErrorNotifier Err.Number, Err.description
    txtCalc(0).SetFocus
End If
End Sub

Private Sub savePayroll()
Dim payRS As Recordset
RSOpen payRS, "SELECT * FROM Payroll", dbOpenDynaset
With payRS
    .AddNew
    .Fields("EmployeeID") = lblID.Caption
    If pay1.Value = True Then
        .Fields("paymentMethod") = pay1.Caption
    ElseIf pay2.Value = True Then
        .Fields("paymentMethod") = pay2.Caption
    Else
        .Fields("paymentMethod") = pay3.Caption
        .Fields("chequeNum") = txtCheque.Text
    End If
    .Fields("dateIssued") = Format$(datepk.Value, "dd/mm/yyyy")
    .Fields("hrsWorked") = txtHours.Text
    .Fields("otHours") = txtOTHours.Text
    .Fields("annualLeaves") = txtAnnual.Text
    .Fields("sickLeaves") = txtSick.Text
    .Fields("unpaidLeaves") = txtUnpaid.Text
    .Fields("epfEmployee") = txtEmployee.Text
    .Fields("epfEmployer") = txtEmployer.Text
    .Fields("otAmount") = txtCalc(1).Text
    .Fields("incomeTax") = txtCalc(5).Text
    .Fields("initialSalary") = txtCalc(0).Text
    .Fields("hourlyRate") = txtRate.Text
    .Fields("socsoAmount") = txtCalc(4).Text
    .Fields("salaryAdvance") = txtCalc(7).Text
    .Fields("incentive") = txtCalc(9).Text
    .Update
End With
payRS.Close
Set payRS = Nothing
InfoMsg "The payroll for the selected employee has been successfully created.", "Record saved"

End Sub

Private Sub autoCalc()
Dim init_sal, ot_bonus, epf_work, epf_emp As Single
Dim socso, inc_tax, net_sal, gross_sal As Single
Dim unpaid_amount As Single
Dim hrRate As Single
'Assign values
hrRate = CSng(txtRate.Text)
hour_work = CSng(txtHours.Text)
ot_hour = CSng(txtOTHours.Text)
ann_leave = CSng(txtAnnual.Text)
sick_leave = CSng(txtSick.Text)

unpaid_leave = CSng(txtUnpaid.Text)
init_sal = CSng(lblSalary.Caption)
'formula

epf_work = getSettings("EPFWorkRate")
epf_emp = getSettings("EPFEmpRate")
unpaid_amount = CSng(hrRate * unpaid_leave)
socso = getEmployeeSocso(init_sal)
gross_sal = init_sal - unpaid_amount
inc_tax = estimateTax(gross_sal)
net_sal = gross_sal - (CSng(txtEmployee.Text) + socso + inc_tax)

'show result
txtEmployee.Text = Format$(CInt(epf_work / 100 * init_sal), "#,##0.00")
txtEmployer.Text = Format$(CInt(epf_emp / 100 * init_sal), "#,##0.00")
txtCalc(0).Text = Format$(init_sal, "#,##0.00")
txtCalc(1).Text = Format$(unpaid_amount, "#,##0.00")
txtCalc(2).Text = Format$(gross_sal, "#,##0.00")
txtCalc(3).Text = txtEmployee.Text
txtCalc(4).Text = Format$(socso, "#,##0.00")
txtCalc(5).Text = Format$(inc_tax, "#,##0.00")
txtCalc(6).Text = Format$(net_sal, "#,##0.00")
'txtCalc(7).Text = Format$(, "#,##0.00") 'Salary advance
txtCalc(8).Text = Format$(net_sal - CSng(txtCalc(7).Text), "#,##0.00")
'txtCalc(9).Text = IIf(IsNull(txtCalc(9).Text), "0", txtCalc(9).Text) '+incentives
txtCalc(10).Text = Format$(CSng(txtCalc(8).Text) + CSng(txtCalc(9).Text), "#,##0.00")
End Sub

Private Function getEmployeeSocso(ByVal initialSalary As Single) As Single
Dim socsoRS As Recordset
RSOpen socsoRS, "SELECT employee FROM socso WHERE amount <= " & initialSalary & ";", dbOpenSnapshot
If Not socsoRS.EOF Then
    socsoRS.MoveLast
    getEmployeeSocso = socsoRS("employee")
Else
    getEmployeeSocso = 0
End If
socsoRS.Close
Set socsoRS = Nothing
End Function

Private Function estimateTax(ByVal ChargeableIncome As Single) As Single
'Based on Malaysia Income Tax 2003 system
'ChargeableIncome is referring to monthly income
On Error GoTo ErrHandler
Dim ci As Single
Dim tax As Integer
ci = ChargeableIncome

If ci <= 2500 Then
    tax = 0
ElseIf ci <= 5000 Then
    tax = 1
ElseIf ci <= 20000 Then
    tax = 3
ElseIf ci <= 35000 Then
    tax = 7
ElseIf ci <= 50000 Then
    tax = 13
ElseIf ci <= 70000 Then
    tax = 19
ElseIf ci <= 100000 Then
    tax = 24
ElseIf ci <= 250000 Then
    tax = 27
Else
    tax = 28
End If

estimateTax = ci * tax / 100
ErrHandler:
If Err.Number <> 0 Then
    CriticalMsg "An error has occurred during the calculation for income tax. Please ensure only valid data is entered.", "Critical error"
    estimateTax = 0
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frmPayroll_New = Nothing
End Sub

Private Sub lvEmployees_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    If .Selected Then
        getEmployeeInfo .Text
    End If
End With
End Sub

Private Sub mnu_Close_Click()
Unload Me
End Sub

Private Sub mnu_Print_Click()
frmPayroll_Print.Show vbModal
End Sub

Private Sub mnu_Save_Click()
Call cmdSave_Click
End Sub

Private Sub pay1_Click()
txtCheque.Enabled = False
End Sub

Private Sub pay2_Click()
txtCheque.Enabled = False
End Sub

Private Sub pay3_Click()
txtCheque.Enabled = True
End Sub

Private Sub txtAnnual_GotFocus()
SelText txtAnnual
End Sub

Private Sub txtAnnual_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtCalc_Change(Index As Integer)
If Len(txtCalc(9).Text) > 0 Then
    autoCalc
End If
End Sub

Private Sub txtCalc_GotFocus(Index As Integer)
SelText txtCalc(Index)
End Sub

Private Sub txtCalc_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtCalc_LostFocus(Index As Integer)
txtCalc(Index).Text = Format$(IIf((txtCalc(Index).Text = ""), "0", txtCalc(Index).Text), "#,##0.00")
End Sub

Private Sub txtCheque_GotFocus()
SelText txtCheque
End Sub

Private Sub txtEmployee_Change()
txtCalc(2).Text = txtEmployee.Text
End Sub

Private Sub txtEmployee_GotFocus()
SelText txtEmployee
End Sub

Private Sub txtEmployee_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtEmployee_LostFocus()
txtEmployee.Text = Format$(IIf((txtEmployee.Text = ""), "0", txtEmployee.Text), "#,##0.00")
End Sub

Private Sub txtEmployer_GotFocus()
SelText txtEmployer
End Sub

Private Sub txtEmployer_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtEmployer_LostFocus()
txtEmployer.Text = Format$(IIf((txtEmployer.Text = ""), "0", txtEmployer.Text), "#,##0.00")
End Sub

Private Sub txtHours_GotFocus()
SelText txtHours
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtOTHours_GotFocus()
SelText txtOTHours
End Sub

Private Sub txtOTHours_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtRate_GotFocus()
SelText txtRate
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtSick_GotFocus()
SelText txtSick
End Sub

Private Sub txtSick_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub

Private Sub txtUnpaid_GotFocus()
SelText txtUnpaid
End Sub

Private Sub txtUnpaid_KeyPress(KeyAscii As Integer)
If KeyAscii <> Asc(".") Then
    OnlyNum KeyAscii
End If
End Sub
