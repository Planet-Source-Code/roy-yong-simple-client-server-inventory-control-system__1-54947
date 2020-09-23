VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmReport_Sales 
   Caption         =   "Sales Drill-Down Report"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport_Sales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   10800
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   3360
      ScaleHeight     =   795
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   4680
      Width           =   7335
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Max             =   14
         Orientation     =   1179649
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAvg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtMin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtMax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtVar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Average"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Minimum"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Maximum"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Variance"
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "St. Deviation"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSChart20Lib.MSChart chart 
      Height          =   3855
      Left            =   3360
      OleObjectBlob   =   "frmReport_Sales.frx":08CA
      TabIndex        =   3
      Top             =   840
      Width           =   7335
   End
   Begin MSComctlLib.Toolbar cb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvReport 
      Height          =   1455
      Left            =   3360
      TabIndex        =   1
      Top             =   5520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
   Begin MSComctlLib.TreeView tvReport 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10821
      _Version        =   393217
      LabelEdit       =   1
      Style           =   4
      SingleSel       =   -1  'True
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
   End
End
Attribute VB_Name = "frmReport_Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CHART_TITLE = "Annual Sales"

Private Sub cb_ButtonClick(ByVal Button As MSComctlLib.Button)
With Button
    chart.Plot.DataSeriesInRow = False
    Select Case Button.Key
    Case "LineChart"
        chart.chartType = VtChChartType2dLine
    Case "BarChart"
        chart.chartType = VtChChartType2dBar
    Case "PieChart"
        chart.chartType = VtChChartType2dPie
        chart.Plot.DataSeriesInRow = True
    End Select
End With
End Sub

Private Sub FlatScrollBar1_Change()
If FlatScrollBar1.Value = 0 Then
  txtTotal.Left = 120
Else
    txtTotal.Left = 0 - FlatScrollBar1.Value * 400
End If
Label1.Left = txtTotal.Left
txtAvg.Left = txtTotal.Left + txtTotal.width + 105
Label2.Left = txtAvg.Left
txtMin.Left = txtAvg.Left + txtAvg.width + 105
Label3.Left = txtMin.Left
txtMax.Left = txtMin.Left + txtMin.width + 105
Label4.Left = txtMax.Left
txtVar.Left = txtMax.Left + txtMax.width + 105
Label5.Left = txtVar.Left
txtSt.Left = txtVar.Left + txtVar.width + 105
Label6.Left = txtSt.Left
End Sub

Private Sub Form_Load()
With tvReport
    .Style = tvwTreelinesText
    .ImageList = frmMain.img16
    .Nodes.add , , "Root", "Annual Report"
End With
With lvReport
    .View = lvwReport
End With
addChild "Root", "SELECT DISTINCT Year(Format([date],'mm/dd/yyyy')) AS [Financial Year] From cust_transactions " & _
"GROUP BY Year(Format([date],'mm/dd/yyyy'));", "[Financial Year]"
With chart
    'Default type
    .chartType = VtChChartType2dBar
    'Establish the number of items in the group
    
    'Turn off the background grids
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull
    .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleNull
    .Plot.Axis(VtChAxisIdY2).AxisGrid.MajorPen.Style = VtPenStyleNull
    .Plot.Wall.Pen.Style = VtPenStyleNull
    'Define the background color to white
    .Backdrop.Fill.Brush.FillColor.Set 255, 255, 255
    .Backdrop.Fill.Style = VtFillStyleBrush
    
    .ShowLegend = True
    '.SeriesColumn = 1
    'Set the title
    .title.Text = CHART_TITLE
    With chart.title.VtFont
        .Name = "Helvetica"
        .Style = VtFontStyleBold
        '.Effect = VtFontEffectUnderline
        .Size = 14
        .VtColor.Set 0, 0, 0
    End With
    .Visible = False
End With

'Set the graphics for the buttons
With cb.Buttons
    .Clear
    cb.ImageList = frmMain.img32
    .add , "PieChart", "Pie Chart", , "pie"
    .add , "BarChart", "Bar Chart", , "bar"
    .add , "LineChart", "Line Chart", , "line"
    .Item(1).ToolTipText = "Pie Chart View"
    .Item(2).ToolTipText = "Bar Chart View"
    .Item(3).ToolTipText = "Line Chart View"
End With
FlatScrollBar1.Top = Picture1.height - (FlatScrollBar1.height + 50)
End Sub

Private Sub Form_Resize()
On Error Resume Next
chart.width = Me.ScaleWidth - (chart.Left + tvReport.Left)
chart.height = Me.ScaleHeight - (lvReport.Top / 2) - (cb.height + Picture1.height)
Picture1.Top = chart.Top + chart.height
lvReport.Top = chart.Top + chart.height + Picture1.height
lvReport.width = Me.ScaleWidth - (lvReport.Left + tvReport.Left)
Picture1.width = lvReport.width
lvReport.height = Me.ScaleHeight - (lvReport.Top + tvReport.Left)
tvReport.height = Me.ScaleHeight - (tvReport.Top + tvReport.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmReport_Sales = Nothing
End Sub

Private Sub Picture1_Resize()
If Picture1.width < 7335 Then
    FlatScrollBar1.width = Picture1.width
    FlatScrollBar1.Visible = True
Else
    FlatScrollBar1.Visible = False
End If
End Sub

Private Sub tvReport_Click()
'With tvReport
'    If .SelectedItem.Children = 0 Then
'    End If
'End With
End Sub

Private Sub tvReport_NodeClick(ByVal Node As MSComctlLib.Node)
If Not Node.Root.Selected Then
    If Node.Parent.Key = "Root" Then
        'year selected
        getChartData Node.Text
    ElseIf Node.Children = 0 Then
        'Months selected
        'convert text to month in digits
        getData Right$(Node.Key, 2), Node.Parent.Text
        getCategorySummary Right$(Node.Key, 2), Node.Parent.Text
    End If
End If
End Sub

Private Sub getChartData(ByVal strYear As String)
'Gets the total sales for the year
With chart
    If .Visible = False Then
        .Visible = True
    End If
    .title.Text = CHART_TITLE & " For The Year " & CStr(strYear)
    Dim chartRS As Recordset
    Dim chartSQL As String
    Dim i As Integer, nSeries As Integer
    Dim arrayData(1 To 12, 1 To 2)
    For i = 1 To 12
        chartSQL = "SELECT Sum(cust_transactions.debit) AS TotalInvoiced " & _
                    "FROM cust_transactions " & _
                    "WHERE (((cust_transactions.date) Like '##/" & Format$(i, "00") & "/" & strYear & "'));"
            
        RSOpen chartRS, chartSQL, dbOpenSnapshot
            arrayData(i, 1) = MonthName(i, False)
            If Not chartRS.EOF Then
                arrayData(i, 2) = CDbl(Format$(IIf(IsNull(chartRS("TotalInvoiced")), "0", chartRS("TotalInvoiced")), "#,##0.00"))
            Else
                arrayData(i, 2) = CDbl(Format$(0, "#,##0.00"))
            End If
    Next i
    chartRS.Close
    Set chartRS = Nothing
    .ChartData = arrayData
    .ColumnCount = 1
    .SeriesColumn = 1
    .ColumnLabel = "Monthly Sales"
    nSeries = .Plot.SeriesCollection.Count
    
    'Add black border to each series
    For i = 1 To nSeries
        .Plot.SeriesCollection(i).DataPoints(-1).EdgePen.VtColor.Set 0, 0, 0
    Next i
    .Refresh
    chartSQL = "SELECT Sum(cust_transactions.debit) AS Total, Avg(cust_transactions.debit) AS Average_Invoice, Min(cust_transactions.debit) AS Minimum, Max(cust_transactions.debit) AS Maximum, StDev(cust_transactions.debit) AS St_Deviation, Var(cust_transactions.debit) AS Var_Invoice " & _
            "FROM cust_transactions " & _
            "WHERE (((cust_transactions.date) Like '##/##/" & strYear & "'));"
    getData "##", strYear 'Data for the year instead of monthly sales shown
End With
End Sub

Private Sub addChild(ByVal strParent As String, ByVal strCondition As String, ByVal strTargetField As String)
Dim treeRS As Recordset, monthRS As Recordset
Dim i As Integer
With tvReport
    RSOpen treeRS, strCondition, dbOpenSnapshot
    While Not treeRS.EOF
        .Nodes.add "Root", tvwChild, "Root" & treeRS(strTargetField), CStr(treeRS(strTargetField))
        'Populate with Months
        For i = 1 To 12
            .Nodes.add CStr("Root" & treeRS(strTargetField)), tvwChild, "Child" & Format(i, "00"), MonthName(i, False)
        Next i
        treeRS.MoveNext
    Wend
    treeRS.Close
    Set treeRS = Nothing
End With
End Sub

Private Sub getData(ByVal strMonth As String, strYear As String)
Dim dataSQL As String
dataSQL = "SELECT Sum(cust_transactions.debit) AS Total, Avg(cust_transactions.debit) AS Average_Invoice, Min(cust_transactions.debit) AS Minimum, Max(cust_transactions.debit) AS Maximum, StDev(cust_transactions.debit) AS St_Deviation, Var(cust_transactions.debit) AS Var_Invoice " & _
          "FROM cust_transactions " & _
          "WHERE (((cust_transactions.date) Like '##/" & strMonth & "/" & strYear & "'));"
With lvReport.ListItems
    'Clear the items and column headers
    .Clear
    lvReport.ColumnHeaders.Clear
    Dim dataRS As Recordset, i As Integer
    'Load recordset based on query
    RSOpen dataRS, dataSQL, dbOpenSnapshot
    If Not dataRS.EOF Then
        txtTotal.Text = Format$(IIf(IsNull(dataRS("Total")), "0", dataRS("Total")), "#,##0.00")
        txtAvg.Text = Format$(IIf(IsNull(dataRS("Average_Invoice")), "0", dataRS("Average_Invoice")), "#,##0.00")
        txtMax.Text = Format$(IIf(IsNull(dataRS("Maximum")), "0", dataRS("Maximum")), "#,##0.00")
        txtMin.Text = Format$(IIf(IsNull(dataRS("Minimum")), "0", dataRS("Minimum")), "#,##0.00")
        txtVar.Text = Format$(IIf(IsNull(dataRS("Var_Invoice")), "0", dataRS("Var_Invoice")), "#,##0.00")
        txtSt.Text = Format$(IIf(IsNull(dataRS("St_Deviation")), "0", dataRS("var_invoice")), "#,##0.00")
    Else
        txtTotal.Text = Format$(0, "#,##0.00")
        txtAvg.Text = Format$(0, "#,##0.00")
        txtMax.Text = Format$(0, "#,##0.00")
        txtMin.Text = Format$(0, "#,##0.00")
        txtVar.Text = Format$(0, "#,##0.00")
        txtSt.Text = Format$(0, "#,##0.00")
    End If
    'Add column headers
    'For i = 0 To dataRS.Fields.Count - 1
    '    lvReport.ColumnHeaders.add , , dataRS.Fields(i).Name
    'Next i
    'Populate with data
    'While Not dataRS.EOF
    '    For i = 0 To dataRS.Fields.Count - 1
    '        If i = 0 Then
    '            .add , , Format$(IIf(IsNull(dataRS.Fields(i).Value), "0.00", dataRS.Fields(i).Value), "#,##0.00")
    '        Else
    '            .Item(.Count).SubItems(i) = Format$(IIf(IsNull(dataRS.Fields(i).Value), "0.00", dataRS.Fields(i).Value), "#,##0.00")
    '        End If
    '    Next i
    '    dataRS.MoveNext
    'Wend
    dataRS.Close
    Set dataRS = Nothing
End With
End Sub

Private Sub getCategorySummary(ByVal strMonth As String, ByVal strYear As String)
Dim catSQL As String
Dim CatRS As Recordset
catSQL = "SELECT Categories.CategoryID, Sum(D_Details.SalePrice) AS SumOfSalePrice " & _
"FROM Delivery INNER JOIN ((Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID) INNER JOIN D_Details ON Products.ProductID = D_Details.ProductID) ON Delivery.DOnumber = D_Details.DOnumber " & _
"WHERE (((D_Details.isInvoiced)=True) AND ((Delivery.Date) Like '##/" & strMonth & "/" & strYear & "')) " & _
"GROUP BY Categories.CategoryID;"
With lvReport
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.add , , "Category ID", 1500
    .ColumnHeaders.add , , "Total Sales for the month of " & MonthName(CLng(strMonth), False), 3000
    RSOpen CatRS, catSQL, dbOpenSnapshot
    While Not CatRS.EOF
        .ListItems.add , , CatRS("CategoryID")
        .ListItems(.ListItems.Count).SubItems(1) = Format$(IIf(IsNull(CatRS("SumOfSalePrice")), "0", CatRS("SumOfSalePrice")), "#,##0.00")
        CatRS.MoveNext
    Wend
    CatRS.Close
    Set CatRS = Nothing
End With
End Sub

