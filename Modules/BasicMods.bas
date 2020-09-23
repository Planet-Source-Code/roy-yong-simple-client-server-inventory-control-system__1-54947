Attribute VB_Name = "BasicModules"
Option Explicit
'Created by: Roy Yong
'Basic modules where procedures and functions are stored and called for
'Global Variables
Global CurrentUser As SynonUser
Global dbWorkspace As Workspace
Global MySynonDatabase As Database
Global isOpen As Boolean 'Use to determine if database is connected
Global pathFileSettings As String
Global DBLocation As String
Global NumDOForm As Byte
Global NumPOForm As Byte

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Boolean) As Long
   
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Public Const ICC_USEREX_CLASSES = &H200
'Public Variables

'Custom Types
Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Type SynonUser
    employeeID As String
    strUsername As String
    strPassword As String
    lastPassword As Date
    mustChange As Boolean
    isLocked As Boolean
    isDisabled As Boolean
    prvlgAdmin As Boolean
    prvlgAPS As Boolean
    prvlgARS As Boolean
    prvlgDOS As Boolean
    prvlgHRS As Boolean
    prvlgReport As Boolean
End Type
'Custom enumerated type
Public Enum ModeStatus
    Editing
    Viewing
End Enum

Public Enum eTrans
    credit
    debit
    transfer
End Enum

Public Enum form_Condition
    force_Change 'Requirement
    mild_Change 'Optional
End Enum

' Functions
Public Function GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
    Dim sBuffer As String, lRet As Long
    ' Fill String with 255 spaces
    sBuffer = String$(255, 0)
    ' Call DLL
    lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        ' DLL failed, save default
        If sDefault <> "" Then AddToINI sSection, sKey, sDefault, sIniFile
        GetFromINI = sDefault
    Else
        ' DLL successful
        ' return string
        GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
End Function

' Returns True if successful. If section does not
' exist it creates it.
Public Function AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    ' Call DLL
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    AddToINI = (lRet)
End Function

Public Function getSettings(ByVal strSubject As String) As Variant
'gets settings from database which is shared among all users
If isOpen = True Then
    'Make sure database connected
    Dim settingRS As Recordset, settingSQL As String
    settingSQL = "SELECT * FROM Pub_settings WHERE Pub_settings.Subject='" & strSubject & "';"
    RSOpen settingRS, settingSQL, dbOpenSnapshot
    If Not settingRS.EOF Then
        getSettings = settingRS("Value")
    Else
        getSettings = ""
        ErrorNotifier 1011, "Unable to retrieve the primary settings due to invalid parameter passed. As a result, this program would not run as normal."
    End If
    settingRS.Close
    Set settingRS = Nothing
End If
End Function

Public Function getNextKeys(ByVal strSubject As String) As Variant
If isOpen = True Then
    'make sure database is connected
    Dim sRS As Recordset, sSQL As String
    sSQL = "SELECT DataValue FROM Misc WHERE DataType='" & strSubject & "';"
    RSOpen sRS, sSQL, dbOpenSnapshot
    If Not sRS.EOF Then
        getNextKeys = sRS("DataValue")
    Else
        getNextKeys = ""
        ErrorNotifier 1011, "Unable to retrieve the primary settings due to invalid parameter passed. As a result, this program would not run as normal."
    End If
    sRS.Close
    Set sRS = Nothing
End If
End Function

Sub Main()
    'This section of comments and some related codes taken from isExplorerBar project
        ' we need to call InitCommonControls before we
        ' can use XP visual styles.  Here I'm using
        ' InitCommonControlsEx, which is the extended
        ' version provided in v4.72 upwards (you need
        ' v6.00 or higher to get XP styles)
        On Error Resume Next
        ' this will fail if Comctl not available
        '  - unlikely now though!
        Dim iccex As tagInitCommonControlsEx
        With iccex
            .lngSize = LenB(iccex)
            .lngICC = ICC_USEREX_CLASSES
        End With
        InitCommonControlsEx iccex

        ' now start the application
    
        On Error GoTo 0
        'Loads the very first form
    'End of copied code
    frmLogin.Show
    pathFileSettings = App.Path & "\settings.ini"
End Sub

'Public procedures
Public Sub DisableClose(frm As Form, Optional _
  Disable As Boolean = True)
    'Setting Disable to False disables the 'X',
     'otherwise, it's reset

    Dim hMenu As Long
    Dim nCount As Long
    
    If Disable Then
        hMenu = GetSystemMenu(frm.hwnd, False)
        nCount = GetMenuItemCount(hMenu)
        
        Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or _
            MF_BYPOSITION)
        Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or _
            MF_BYPOSITION)
    
        DrawMenuBar frm.hwnd
    Else
        GetSystemMenu frm.hwnd, True
        
        DrawMenuBar frm.hwnd
    End If
End Sub

Public Sub ValidMsg(ByVal strMssg As String, strTitle As String)
MsgBox strMssg, vbExclamation + vbOKOnly, strTitle
End Sub

Public Sub CriticalMsg(ByVal strMssg As String, strTitle As String)
MsgBox strMssg, vbCritical + vbOKOnly, strTitle
End Sub

Public Sub InfoMsg(ByVal strMssg As String, strTitle As String)
MsgBox strMssg, vbInformation + vbOKOnly, strTitle
End Sub

Public Sub SelText(ByRef strControlName As Control)
'Highlights the string within the control
On Error Resume Next
strControlName.SelStart = 0
strControlName.SelLength = Len(strControlName.Text)
End Sub

Public Sub CapCon(ByRef strControlName As Control)
'Converts the string passed in to upper case
On Error Resume Next
strControlName.Text = UCase(strControlName.Text)
End Sub
Public Sub OnlyNum(ByRef strAscii As Integer)
'Strict limitations. Non-digital keys converted to 0
Select Case strAscii
Case Asc("0") To Asc("9")
    'Do nothing
Case vbKeyBack
    'do nothing
Case Else
    strAscii = 0
End Select
End Sub

Public Sub onlyPassword(ByRef strAscii As Integer)
Select Case strAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc("0") To Asc("9"), Asc("_"), vbKeySpace, vbKeyBack
    'Do nothing
    Case Else
        strAscii = 0
End Select
End Sub

Public Sub ErrorNotifier(ByVal ErrNumber As Long, ByVal ErrDescription As String)
Dim errorRS As Recordset, errorSQL As String
Screen.MousePointer = 0
CriticalMsg "Error No: " & ErrNumber & vbCrLf & "Description: " & ErrDescription & vbCrLf & "Please contact system administrator.", "Critical Error"
If isOpen = True Then
    errorSQL = "INSERT INTO Error_Logging VALUES ('" & Format$(Now(), "dd/mm/yyyy") & "','" & CStr(ErrNumber) & "','" & ErrDescription & "');"
    On Error GoTo ErrHandler
    MySynonDatabase.Execute errorSQL
End If
ErrHandler:
If Err.Number <> 0 Then
    Exit Sub
End If
End Sub

Public Sub OnlyAlpha(ByRef strAscii As Integer)
'Strict limitations. Non-alphabet keys converted to 0
Select Case strAscii
Case Asc("A") To Asc("Z")
    'Do nothing
Case Asc("a") To Asc("z")
    'Do nothing
Case vbKeyBack, vbKeySpace
    'Do nothing
Case Else
    strAscii = 0
End Select
End Sub

Public Sub tickerKeys(ByRef strAscii As Integer)
Select Case strAscii
    Case Asc("."), Asc("!"), Asc(","), Asc("("), Asc(")"), Asc(":"), Asc(";"), Asc("?"), Asc("/"), Asc("0") To Asc("9")
        'Do nothing
    Case Else
        OnlyAlpha strAscii
End Select
End Sub

Public Sub FillComboCountry(ByRef strComboBox As ComboBox)
Dim cmbRS As Recordset

On Error GoTo ErrHandler
RSOpen cmbRS, "SELECT Countries.CountryID FROM Countries;", dbOpenSnapshot
While Not cmbRS.EOF
    strComboBox.addItem cmbRS("CountryID")
    cmbRS.MoveNext
Wend

cmbRS.Close
Set cmbRS = Nothing
ErrHandler:
If Err.Number <> 0 Then
    Exit Sub
End If
End Sub

Public Sub FillComboState(ByRef strComboBox As ComboBox, ByVal strCountryID As String)
Dim comboRS As Recordset, tmpString As String
tmpString = strComboBox.Text
strComboBox.Clear
On Error GoTo ErrHandler
RSOpen comboRS, "SELECT States.StateID FROM States WHERE States.CountryID='" & strCountryID & "';", dbOpenSnapshot
While Not comboRS.EOF
    strComboBox.addItem comboRS("StateID")
    comboRS.MoveNext
Wend
comboRS.Close
Set comboRS = Nothing
strComboBox.Text = tmpString
ErrHandler:
If Err.Number <> 0 Then
    Exit Sub
End If
End Sub

Public Sub FillComboCity(ByRef strComboBox As ComboBox, ByVal strStateID As String)
Dim comboRS As Recordset, tmpString As String
tmpString = strComboBox.Text
strComboBox.Clear
On Error GoTo ErrHandler
RSOpen comboRS, "SELECT Cities.City FROM Cities WHERE Cities.StateID='" & strStateID & "';", dbOpenSnapshot
While Not comboRS.EOF
    strComboBox.addItem comboRS("City")
    comboRS.MoveNext
Wend
comboRS.Close
Set comboRS = Nothing
strComboBox.Text = tmpString
ErrHandler:
If Err.Number <> 0 Then
    Exit Sub
End If
End Sub

Public Sub FillCombo(ByRef strComboBox As ComboBox, ByVal strSQL As String, ByVal strItem As String)
Dim comboRS As Recordset

'On Error GoTo ErrHandler
RSOpen comboRS, strSQL, dbOpenSnapshot
While Not comboRS.EOF
    If IsNull(comboRS(strItem)) = False Then
        strComboBox.addItem comboRS(strItem)
    End If
    comboRS.MoveNext
Wend
comboRS.Close
Set comboRS = Nothing

ErrHandler:
    Exit Sub
End Sub

Public Sub ConnDB()
On Error GoTo ErrHandler
DBLocation = GetFromINI("Database", "Path", "", pathFileSettings)
Set dbWorkspace = DBEngine.Workspaces(0)
Set MySynonDatabase = dbWorkspace.OpenDatabase(DBLocation)

isOpen = True

ErrHandler:
If Err.Number <> 0 Then
    If Err.Number = 3044 Then 'Cannot find database
        isOpen = False
        Set dbWorkspace = Nothing
        Set MySynonDatabase = Nothing
        ErrorNotifier Err.Number, "Unable to locate database given the file path. Check your settings under Options."
        DoEvents
    End If
End If
End Sub

Public Sub RSOpen(ByRef strRecordset As Recordset, ByVal strQuery As String, Optional ByVal connLocking As RecordsetTypeEnum)
If isOpen = False Then
    Call ConnDB
End If
If IsMissing(connLocking) = True Then
    connLocking = dbOpenDynaset
End If
Set strRecordset = MySynonDatabase.OpenRecordset(strQuery, connLocking)
End Sub

Public Sub closeDB()
MySynonDatabase.Close
Set MySynonDatabase = Nothing
End Sub

Public Function isDateValid(ByVal bDay As Byte, ByVal bMonth As Byte, ByVal iYear As Integer) As Boolean
'Attempts to verify if a date is valid or not. Values are passed by parameter.
isDateValid = True
If (bDay < 0) Or (bMonth < 0) Or (iYear < 0) Then
    isDateValid = False
Else
    Select Case bMonth
        Case 1, 3, 5, 7, 8, 10, 12
            If bDay > 31 Then
                isDateValid = False
            End If
        Case 4, 6, 9, 11
            If bDay > 30 Then
                isDateValid = False
            End If
        Case 2
            If iYear Mod 2 = 0 Then
                If bDay > 29 Then
                    isDateValid = False
                End If
            Else
                If bDay > 28 Then
                    isDateValid = False
                End If
            End If
        Case Else
            isDateValid = False
    End Select
End If
End Function
Public Function processCustTransaction(ByVal transDate As String, ByVal description As String, ByVal transType As eTrans, ByVal accountID As String, ByVal amount As Single) As Boolean
Dim proSQL As String
Dim proRS As Recordset
proSQL = "SELECT * FROM cust_transactions"
Set proRS = MySynonDatabase.OpenRecordset(proSQL, dbOpenDynaset, dbAppendOnly)
On Error GoTo ErrHandler
proRS.AddNew
proRS("date") = transDate
If transType = credit Then
    proRS("credit") = amount
ElseIf transType = debit Then
    proRS("debit") = amount
End If
proRS("notes") = description
proRS("CustomerID") = accountID
proRS.Update

'Fulfilling double entry rule
proSQL = "SELECT CurrentBalance FROM Customers WHERE CustomerID='" & accountID & "';"
Set proRS = MySynonDatabase.OpenRecordset(proSQL, dbOpenDynaset)
proRS.Edit
If transType = credit Then
    proRS("CurrentBalance") = proRS("CurrentBalance") - amount
Else
    proRS("CurrentBalance") = proRS("CurrentBalance") + amount
End If
proRS.Update

proRS.Close
Set proRS = Nothing
processCustTransaction = True

ErrHandler:
If Err.Number <> 0 Then
    processCustTransaction = False
End If
End Function

Public Function processSuppTransaction(ByVal transDate As String, ByVal description As String, ByVal transType As eTrans, ByVal accountID As String, ByVal amount As Single) As Boolean
Dim proSQL As String
Dim proRS As Recordset
proSQL = "SELECT * FROM supp_transactions"
Set proRS = MySynonDatabase.OpenRecordset(proSQL, dbOpenDynaset, dbAppendOnly)
On Error GoTo ErrHandler
proRS.AddNew
proRS("date") = transDate
If transType = credit Then
    proRS("debit") = amount
ElseIf transType = debit Then
    proRS("credit") = amount
End If
proRS("notes") = description
proRS("SupplierID") = accountID
proRS.Update

'Fulfilling double entry rule
proSQL = "SELECT CurrentBalance FROM Suppliers WHERE SupplierID='" & accountID & "';"
Set proRS = MySynonDatabase.OpenRecordset(proSQL, dbOpenDynaset)
proRS.Edit
If transType = credit Then
    proRS("CurrentBalance") = proRS("CurrentBalance") + amount
Else
    proRS("CurrentBalance") = proRS("CurrentBalance") - amount
End If
proRS.Update

proRS.Close
Set proRS = Nothing
processSuppTransaction = True

ErrHandler:
If Err.Number <> 0 Then
    processSuppTransaction = False
End If
End Function

Public Sub insertLog(ByVal strLog As String)
If isOpen Then
    'Insert into systems log
    MySynonDatabase.Execute "INSERT INTO Logging VALUES ('" & CurrentUser.strUsername & "','" & strLog & "','" & FormatDateTime(Now(), vbLongTime) & "','" & Format$(Now(), "dd/mm/yyyy") & "');"
End If
End Sub

Public Function NumToString(ByVal nNumber As Currency) As String
'Written by Scott Seligman. Copied from www.freevbcode.com/ShowCode.asp?ID=343
Dim bNegative As Boolean
Dim bHundred As Boolean

If nNumber < 0 Then
    bNegative = True
End If

nNumber = Abs(Int(nNumber))

If nNumber < 1000 Then
    If nNumber \ 100 > 0 Then
        NumToString = NumToString & _
             NumToString(nNumber \ 100) & " hundred"
        bHundred = True
    End If
    nNumber = nNumber - ((nNumber \ 100) * 100)
    Dim bNoFirstDigit As Boolean
    bNoFirstDigit = False
    Select Case nNumber \ 10
        Case 0
            Select Case nNumber Mod 10
                Case 0
                    If Not bHundred Then
                        NumToString = NumToString & " zero"
                    End If
                Case 1: NumToString = NumToString & " one"
                Case 2: NumToString = NumToString & " two"
                Case 3: NumToString = NumToString & " three"
                Case 4: NumToString = NumToString & " four"
                Case 5: NumToString = NumToString & " five"
                Case 6: NumToString = NumToString & " six"
                Case 7: NumToString = NumToString & " seven"
                Case 8: NumToString = NumToString & " eight"
                Case 9: NumToString = NumToString & " nine"
            End Select
            bNoFirstDigit = True
        Case 1
            Select Case nNumber Mod 10
                Case 0: NumToString = NumToString & " ten"
                Case 1: NumToString = NumToString & " eleven"
                Case 2: NumToString = NumToString & " twelve"
                Case 3: NumToString = NumToString & " thirteen"
                Case 4: NumToString = NumToString & " fourteen"
                Case 5: NumToString = NumToString & " fifteen"
                Case 6: NumToString = NumToString & " sixteen"
                Case 7: NumToString = NumToString & " seventeen"
                Case 8: NumToString = NumToString & " eighteen"
                Case 9: NumToString = NumToString & " nineteen"
            End Select
            bNoFirstDigit = True
        Case 2: NumToString = NumToString & " twenty"
        Case 3: NumToString = NumToString & " thirty"
        Case 4: NumToString = NumToString & " forty"
        Case 5: NumToString = NumToString & " fifty"
        Case 6: NumToString = NumToString & " sixty"
        Case 7: NumToString = NumToString & " seventy"
        Case 8: NumToString = NumToString & " eighty"
        Case 9: NumToString = NumToString & " ninety"
    End Select
    If Not bNoFirstDigit Then
        If nNumber Mod 10 <> 0 Then
            NumToString = NumToString & "-" & _
                          Mid(NumToString(nNumber Mod 10), 2)
        End If
    End If
Else
    Dim nTemp As Currency
    nTemp = 10 ^ 12 'trillion
    Do While nTemp >= 1
        If nNumber >= nTemp Then
            NumToString = NumToString & _
                          NumToString(Int(nNumber / nTemp))
            Select Case Int(Log(nTemp) / Log(10) + 0.5)
                Case 12: NumToString = NumToString & " trillion"
                Case 9: NumToString = NumToString & " billion"
                Case 6: NumToString = NumToString & " million"
                Case 3: NumToString = NumToString & " thousand"
            End Select
           
            nNumber = nNumber - (Int(nNumber / nTemp) * nTemp)
        End If
        nTemp = nTemp / 1000
    Loop
End If

If bNegative Then
    NumToString = " negative" & NumToString
End If
    
End Function

Public Function DollarToString(ByVal nAmount As Currency) As String
'Written by Scott Seligman. Copied from www.freevbcode.com/ShowCode.asp?ID=343

    Dim nDollar As Currency
    Dim nCent As Currency
    
    nDollar = Int(nAmount)
    nCent = (Abs(nAmount) * 100) Mod 100
    
    DollarToString = NumToString(nDollar) & " dollar"
    
    If Abs(nDollar) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
    DollarToString = DollarToString & " and" & _
                     NumToString(nCent) & " cent"
                     
    If Abs(nCent) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
End Function


