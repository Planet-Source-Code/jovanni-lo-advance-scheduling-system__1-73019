Attribute VB_Name = "modFunc"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public PathToDoc As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hWnd As Long, ByVal lpOperation As String, _
                ByVal lpFile As String, ByVal lpParameters As String, _
                ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ExecLogs()
If DbMgt = True Then Exit Sub
Screen.MousePointer = 11
mdiMain.cmdWarning.Caption = Warnings & " W&arnings"
mdiMain.cmdReminders.Caption = Notifications & " N&otification"
Screen.MousePointer = 0
End Sub

Public Function NoRcrd(ListView As Variant, Optional Prompt As String) As Boolean
If ListView.ListItems.Count = 0 Then
    Screen.MousePointer = 0
    If Prompt <> Empty Then
        MsgBox Prompt, vbExclamation
    End If
    NoRcrd = True
Else
    NoRcrd = False
End If
End Function

Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As String) As String
Dim RcrdNo As Integer
Dim owat As Integer
RunSql "Select * from " & Table & " order by " & FldNo & " ASC"
With Rs
    If Rs.EOF = False Then
        .MoveLast
        If Identifier = Empty Then
            RcrdNo = Val(.Fields(FldNo)) + 1
        Else
            RcrdNo = Val((Right(.Fields(FldNo), Val(Len(.Fields(FldNo))) - Val(Len(Identifier))))) + 1
        End If
    Else
        RcrdNo = 1
    End If
    If Identifier <> Empty Then
        RcrdId = Identifier & RcrdNo
    Else
        RcrdId = RcrdNo
    End If
End With
End Function

Public Function ValBox(Prompt As String, Icon As Image, Optional Title As String, _
                        Optional Default As Double, _
                        Optional Header As String = "Value Box") As Double
With frmValue
    If Title <> Empty Then
        .Caption = Title
    Else
        .Caption = App.Title
    End If
    .lblHeader.Caption = StrConv(Header, vbUpperCase)
    .imgIcon.Picture = Icon.Picture
    .lblPrompt.Caption = Prompt
    .Default Val(Default)
    .Show 1
    ValBox = Val(.txtValue.Text)
    Unload frmValue
End With
End Function

Public Function StrBox(Prompt As String, Icon As Image, Optional Title As String, _
                        Optional Default As String, _
                        Optional Header As String = "Text Box", _
                        Optional Style As Integer = 1, _
                        Optional Table As String, _
                        Optional RcrdFld As String, _
                        Optional RcrdStr As String) As String
With frmString
    If Title <> Empty Then
        .Caption = Title
    Else
        .Caption = App.Title
    End If
    .lblHeader.Caption = StrConv(Header, vbUpperCase)
    .imgIcon.Picture = Icon.Picture
    .lblPrompt.Caption = Prompt
    Select Case Style
        Case 1
            Set x = .txtStr
            x.Text = Default
        Case 2
            Set x = .cboStr
            If Table <> Empty Then
                LoadCbo Table, .cboStr, RcrdFld, Default, 0, , RcrdFld, RcrdStr
            Else
                x.Text = Default
            End If
        Case 3
            Set x = .cboStr2
            If Table <> Empty Then
                LoadCbo Table, .cboStr2, RcrdFld, Default, 0, , RcrdFld, RcrdStr
            Else
                x.Text = Default
            End If
    End Select
    x.TabIndex = 0
    x.Visible = True
    .Show 1
    StrBox = x.Text
    Unload frmString
End With
End Function

Public Sub LoadCbo(Table As String, _
                    Control As ComboBox, _
                    FldStr As String, _
                    Optional Default As String, _
                    Optional DefaultInt As Integer, _
                    Optional Locked As Boolean = False, _
                    Optional RcrdFld As String, _
                    Optional RcrdStr As String = "%")
If RcrdFld = Empty Then
    RunSql "select distinct " & FldStr & " from " & Table
Else
    RunSql "Select distinct " & FldStr & " from " & Table & " where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
End If
With Rs
    Control.Clear
    If Default <> Empty Then
        Control.AddItem Default
    End If
    While Not .EOF = True
        Control.AddItem (.Fields(FldStr))
        .MoveNext
    Wend
End With
Control.Locked = Locked
If Control.ListCount <> 0 Then
    Control.ListIndex = DefaultInt
End If
End Sub

Public Sub LoadCboFld(Table, Fields, Control As ComboBox, Optional Default As Integer = 0)
RunSql "Select " & Fields & " from " & Table
With Rs
    Control.Clear
    For i = 0 To (.Fields.Count - 1)
        Control.AddItem (.Fields(i).Name)
    Next i
End With
Control.ListIndex = Default
End Sub

Public Function CboEmp(ByRef ComboBox As ComboBox, _
                Optional TabObject As ComctlLib.TabStrip, _
                Optional TabIndex As Integer) _
                As Boolean
'function for combobox value is default (Select)
If ComboBox.ListIndex = 0 Then
    CboEmp = True
    MsgBox "Please select a record from the list.", vbExclamation
    If TabIndex <> Empty Then
        TabObject.SelectedItem = TabObject.Tabs(TabIndex)
    End If
    ComboBox.SetFocus
Else
    CboEmp = False
End If
End Function

Public Sub DtpValue(dtpObject As DTPicker)
dtpObject.Value = Now
End Sub

Public Sub SelAll(ByRef stext As Variant)
'highlight textbox on focus
With stext
    .SelStart = 0
    .SelLength = Len(stext)
End With
End Sub
Public Function TxtEmp(ByRef stext As Variant, _
                        Optional TabObject As ComctlLib.TabStrip, _
                        Optional TabIndex As Integer) _
                        As Boolean
Screen.MousePointer = 0
'if the textbox is empty then TxtEmp = true
If Trim(stext) = Empty Or stext.Text = "  /  /    " Then
    TxtEmp = True
    MsgBox "Please fill in all required fields.", vbExclamation
    If TabIndex <> Empty Then
        TabObject.SelectedItem = TabObject.Tabs(TabIndex)
    End If
    stext.SetFocus
Else
    TxtEmp = False
End If
End Function

Public Function txtNum(ByRef stext As Variant) As Boolean
'if the input is not a numeric then true
If IsNumeric(stext) = False Then
    txtNum = True
    MsgBox "The field requires a numeric value.", vbExclamation
    stext.SetFocus
    SelAll stext
Else
    txtNum = False
End If
End Function

Public Function UserLimit(ByRef lvl As String, ByRef SysLvl As String) As Boolean
'only the administrator can access some stuffs
If lvl = "Administrator" Then
    UserLimit = False
    Exit Function
End If
If SysLvl <> lvl Then
    If lvl <> Empty And ReadINI("Preferences", "Other Users") = "1" Then
        UserLimit = False
    Else
        Screen.MousePointer = 0
        MsgBox "You dont have the right to access this task. Please Log in as 'Administrator'.", vbExclamation
        UserLimit = True
    End If
Else
    UserLimit = False
End If
End Function

Public Sub FrmShow(lvl As String)
If lvl = "Administrator" Or lvl = "User" Then
    Select Case ReadINI("Preferences", "Default Form")
        Case "Items"
            frmItems.Show
        Case "Transactions"
            frmTransactions.Show 1
        Case "Summary of Transactions"
            frmTransView.Show
        Case Else
            mdiMain.Show
    End Select
Else
    mdiMain.Show
End If
End Sub

Public Sub SetLv(ListView As ListView, Optional FullRow As Boolean, Optional GridLines As Boolean)
If GridLines = True Then
    LvGrid ListView
End If
If FullRow = True Then
    LvFullRow ListView
End If
End Sub

Public Function ReadINI(strKey As String, strName As String) As String
Dim intLen As Integer
Dim strText As String
strText = "                                                                                                    "
intLen = GetPrivateProfileString(strKey, strName, "", strText, Len(strText), App.Path & "\Settings.ini")
If intLen > -1 Then
    strText = Left(strText, intLen)
Else
    MsgBox "Error on reading configuration", vbCritical
    End
End If
ReadINI = strText
End Function

Public Sub WriteINI(strKey As String, strName As String, strText As String)
Dim intLen As Integer
intLen = WritePrivateProfileString(strKey, strName, strText, App.Path & "\Settings.ini")
End Sub

Public Function Scheduler(BaseDate As Date, _
                        GapVal As Integer, _
                        Optional Interval As String) As Date
                        
Dim Max As Long, LastVal As Integer
Dim IntM As Integer, IntD As Integer, IntY As Integer
Dim IntH As Integer, IntN As Integer
Dim strPmAm As String
IntH = Format(BaseDate, "hh")
IntN = Format(BaseDate, "nn")
strPmAm = Format(BaseDate, "ampm")
IntM = Format(BaseDate, "mm")
IntD = Format(BaseDate, "dd")
IntY = Format(BaseDate, "yyyy")
Select Case Interval
    Case "Minute"
        Scheduler = Format(DateAdd("n", GapVal, BaseDate), "mm/dd/yyyy h:n ampm")
        Exit Function
    Case "Hour"
        Scheduler = Format(DateAdd("h", GapVal, BaseDate), "mm/dd/yyyy h:n ampm")
        Exit Function
    Case "Day"
        Max = Val(ReadINI("Month Max", MonthName(IntM, True)))
        IntD = IntD + GapVal
        For i = 1 To IntD
            If i = Max Then
                LastVal = Max
                IntM = IntM + 1
                If IntM > 12 Then IntM = 1: IntY = IntY + 1
                Max = Max + Val(ReadINI("Month Max", MonthName(IntM, True)))
            End If
        Next i
        IntD = IntD - LastVal
    Case "Week"
        Dim MaxDays As Integer
        MaxDays = 7 * GapVal
        Max = Val(ReadINI("Month Max", MonthName(IntM, True)))
        IntD = IntD + MaxDays
        For i = 1 To IntD
            If i = Max Then
                LastVal = Max
                IntM = IntM + 1
                If IntM > 12 Then IntM = 1: IntY = IntY + 1
                Max = Max + Val(ReadINI("Month Max", MonthName(IntM, True)))
            End If
        Next i
        IntD = IntD - LastVal
    Case "Month"
        IntM = IntM + GapVal
        If IntM > 12 Then IntM = IntM - 12: IntY = IntY + 1
    Case "Year"
        IntY = IntY + GapVal
    Case Else
        Scheduler = Now
        Exit Function
End Select
Scheduler = DateSerial(IntY, IntM, IntD)
End Function

Public Function Warnings(Optional dType As Integer) As Integer
DetectionType = dType
RunSql "Delete * from tblDetections"

'-------no available qty
RunSql "SELECT reg.*, format((Select sum(stat.qty) from tblItemStatus as stat " & _
        "INNER JOIN tblStatus ON stat.status = tblStatus.description " & _
        "where stat.item_id = reg.item_id and tblStatus.include = 1), '#0') AS Available, " & _
        "format((Select sum(stat.qty) from tblItemStatus as stat " & _
        "INNER JOIN tblStatus ON stat.status = tblStatus.description " & _
        "where stat.item_id = reg.item_id and tblStatus.include = 0), '#0') AS Unavailable " & _
        "FROM tblRegistered AS reg"
With Rs
    While Not .EOF = True
        If .Fields!available = 0 Then
            s = "No available quantity for this item on your item list. Please review transactions for more details." & vbNewLine & vbNewLine & _
                "Item Description: " & .Fields!Description & vbNewLine & vbNewLine & _
                "Total count: " & .Fields!qty & vbNewLine & vbNewLine & _
                "Unavailable: " & .Fields!Unavailable
            SaveDetection .Fields!item_id, "Item Status", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'-----Item not registered
RunSql "Select * from tblItemList"
With Rs
    While Not .EOF = True
        SubSql "Select * from tblRegistered where item_id = '" & .Fields!item_id & "'"
        If SubRs.EOF = True Then
            s = "Item " & .Fields!item_id & " is not yet registered on your Registered List. " & _
                "Register this item to include this on transactions." & vbNewLine & vbNewLine & _
                "Description: " & .Fields!Description
            SaveDetection .Fields!item_id, "Unregistered", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'-----Did not return item.

RunSql "Select b.*, u.*, t.*, r.* from ((tblBorrow as b INNER JOIN tblClientProfile as u ON " & _
        "b.client_no = u.client_no) INNER JOIN tblTransactions as t ON b.trans_no = t.trans_no) " & _
        "INNER JOIN tblRegistered as r ON b.item_id = r.item_id"
With Rs
    While Not .EOF = True
        d = Scheduler(.Fields!trans_date, .Fields!gap_val, .Fields!Interval)
        If DateDiff(ReadINI("Interval", .Fields!Interval), d, Now) > 0 Then
            s = "Client " & .Fields("u.client_no") & " did not return a borrowed item." & vbNewLine & vbNewLine & _
                "Client: " & StrConv(.Fields!fname & " " & Left(.Fields!mname, 1) & ". " & .Fields!lname, vbProperCase) & vbNewLine & _
                "Item: " & .Fields!Description & "; Status: " & .Fields("b.status") & "; qty: " & .Fields("b.qty") & vbNewLine & _
                "Borrowed Date: " & .Fields!trans_date & vbNewLine & _
                "Time Interval: " & .Fields!gap_val & " " & .Fields!Interval & "(s)" & vbNewLine & _
                "Expected Return: " & d & vbNewLine & _
                "Transaction #: " & .Fields("t.trans_no")
            SaveDetection .Fields("u.client_no"), "Unreturned Items", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'-----did not update status
RunSql "Select * from tblRegistered"
With Rs
    While Not .EOF = True
        SubSql "Select * from tblItemStatus where item_id = '" & .Fields!item_id & "'"
        If SubRs.EOF = True Then
            s = "Status of this item is not yet updated. Please update it to include on transactions." & vbNewLine & vbNewLine & _
                "Item Description: " & .Fields!Description & vbNewLine & _
                "Item Count: " & .Fields!qty
            SaveDetection .Fields!item_id, "Item Status", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'-----date exceeded on reserved items
RunSql "Select re.*, u.*, t.*, r.* from ((tblReserve as re INNER JOIN tblClientProfile as u ON " & _
        "re.client_no = u.client_no) INNER JOIN tblTransactions as t ON re.trans_no = t.trans_no) " & _
        "INNER JOIN tblRegistered as r ON re.item_id = r.item_id"
With Rs
    While Not .EOF = True
        d = DateDiff("d", .Fields!reserve_date, Now)
        If d > 0 Then
            s = "Client " & .Fields("u.client_no") & " reservation has expired. Please check transactions for details." & vbNewLine & vbNewLine & _
                "Client: " & StrConv(.Fields!fname & " " & Left(.Fields!mname, 1) & ". " & .Fields!lname, vbProperCase) & vbNewLine & _
                "Item: " & .Fields!Description & "; Status: " & .Fields("re.status") & "; qty: " & .Fields("re.qty") & vbNewLine & _
                "Reserved Date: " & .Fields!reserve_date & vbNewLine & _
                "Transaction Date: " & .Fields!trans_date & vbNewLine & _
                "Transaction #: " & .Fields("t.trans_no")
            SaveDetection .Fields("u.client_no"), "Expired Reservation", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

RunSql "Select * from tblDetections"
With Rs
    Warnings = .RecordCount
End With

End Function

Public Function Notifications(Optional dType As Integer) As Integer
DetectionType = dType
RunSql "Delete * from tblDetections"

'--------schedule task
RunSql "Select s.title as title, s.remarks as remarks, s.sched_date as sched_date, u.lname as lname, u.mname as mname, u.fname as fname " & _
        "from tblSchedules as s INNER JOIN tblAccountProfile as u ON s.user_id = u.id"
With Rs
    While Not .EOF = True
        If Format(.Fields!sched_date, "mm/dd/yyyy") = Format(Date, "mm/dd/yyyy") Then
            s = "Scheduled task is today." & vbNewLine & vbNewLine & _
                "Remarks: " & .Fields!remarks & vbNewLine & vbNewLine & _
                "Date: " & Format(.Fields!sched_date, "mm/dd/yyyy") & vbNewLine & vbNewLine & _
                "User: " & .Fields!lname & ", " & .Fields!fname & " " & StrConv(Left(.Fields!mname, 1), vbUpperCase) & "."
            SaveDetection Format(.Fields!Title, "mm/dd/yyyy"), "Scheduled Task", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'---------return date
RunSql "Select b.*, u.*, t.*, r.* from ((tblBorrow as b INNER JOIN tblClientProfile as u ON " & _
        "b.client_no = u.client_no) INNER JOIN tblTransactions as t ON b.trans_no = t.trans_no) " & _
        "INNER JOIN tblRegistered as r ON b.item_id = r.item_id"
With Rs
    While Not .EOF = True
        d = Scheduler(.Fields!trans_date, .Fields!gap_val, .Fields!Interval)
        If DateDiff(ReadINI("Interval", .Fields!Interval), d, Now) = 0 Then
            s = "Client " & .Fields("u.client_no") & " will return the item today." & vbNewLine & vbNewLine & _
                "Client: " & StrConv(.Fields!fname & " " & Left(.Fields!mname, 1) & ". " & .Fields!lname, vbProperCase) & vbNewLine & _
                "Item: " & .Fields!Description & "; Status: " & .Fields("b.status") & "; qty: " & .Fields("b.qty") & vbNewLine & _
                "Borrowed Date: " & .Fields!trans_date & vbNewLine & _
                "Time Interval: " & .Fields!gap_val & " " & .Fields!Interval & "(s)" & vbNewLine & _
                "Expected Return: " & d & vbNewLine & _
                "Transaction #: " & .Fields("t.trans_no")
            SaveDetection .Fields("u.client_no"), "Return Items", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'---------reserved date
RunSql "Select re.*, u.*, t.*, r.* from ((tblReserve as re INNER JOIN tblClientProfile as u ON " & _
        "re.client_no = u.client_no) INNER JOIN tblTransactions as t ON re.trans_no = t.trans_no) " & _
        "INNER JOIN tblRegistered as r ON re.item_id = r.item_id"
With Rs
    While Not .EOF = True
        d = DateDiff("d", .Fields!reserve_date, Now)
        If d = 0 And DateDiff("h", .Fields!reserve_date, Now) <= 0 Then
            s = "Client " & .Fields("u.client_no") & " reservation date is today. Check transactions for details." & vbNewLine & vbNewLine & _
                "Client: " & StrConv(.Fields!fname & " " & Left(.Fields!mname, 1) & ". " & .Fields!lname, vbProperCase) & vbNewLine & _
                "Item: " & .Fields!Description & "; Status: " & .Fields("re.status") & "; qty: " & .Fields("re.qty") & vbNewLine & _
                "Reserved Date: " & .Fields!reserve_date & vbNewLine & _
                "Transaction Date: " & .Fields!trans_date & vbNewLine & _
                "Transaction #: " & .Fields("t.trans_no")
            SaveDetection .Fields("u.client_no"), "Reservations", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

'-------low item count
RunSql "Select * from tblRegistered"
With Rs
    If .RecordCount = 0 Or .RecordCount < 5 Then
        s = "You have a minimun items registered on the list. " & _
            "Please add and register new items." & vbNewLine & vbNewLine & _
            "Registered Items: " & .RecordCount
        SaveDetection "Items", "Add Items", s, "tblDetections"
    End If
End With

RunSql "Select * from tblDetections"
With Rs
    Notifications = .RecordCount
End With

End Function

Private Sub SaveDetection(Reference As String, Title As String, Description As String, Table As String)
SubSql "Select * from " & Table
With SubRs
    .AddNew
    .Fields!record_no = Val(RcrdId(Table, , "record_no"))
    .Fields!Reference = Reference
    .Fields!war_type = Title
    .Fields!Description = Description
    .Update
End With
End Sub

