Attribute VB_Name = "modTrans"
Option Explicit

Dim RcrdNo As Integer, Table As String
'--------------Borrow Transaction Modules-----------------------
Public Sub ClearBItems()
With frmTransactions
    .lblBCount.Caption = "---"
    .lblBLocation.Caption = "---"
    .lblBQty.Caption = "---"
    .lblBReturn.Caption = "---"
    .lblBStatus.Caption = "---"
    .txtBTime.Text = Empty
    .cboBGap.ListIndex = 2
    .lblTrans.Caption = RcrdId("tblTransactions", Format(Date, "yymmdd"), "trans_no")
    ViewItemList "tblBorrow", "item_id", "%", frmTransactions.lvwBItems, .lblId.Caption
End With
End Sub

Public Function SaveBTrans() As String
RcrdNo = 0
With frmTransactions
    SaveTrans
    RunSql "Select * from tblBorrow order by record_no ASC"
    If Rs.EOF = False Then
        Rs.MoveLast
        RcrdNo = Rs.Fields!record_no
    End If
    For i = 1 To .lvwBView.ListItems.Count
        RunSql "Select * from tblBorrow"
        Rs.AddNew
        Rs.Fields!record_no = RcrdNo + i
        Rs.Fields!client_no = .lblId.Caption
        Rs.Fields!item_id = .lvwBView.ListItems(i).Text
        Rs.Fields!Status = .lvwBView.ListItems(i).SubItems(4)
        Rs.Fields!qty = .lvwBView.ListItems(i).SubItems(1)
        Rs.Fields!gap_val = .lvwBView.ListItems(i).SubItems(2)
        Rs.Fields!Interval = .lvwBView.ListItems(i).SubItems(3)
        Rs.Fields!trans_no = .lblTrans.Caption
        Rs.Update
        UpdateStat .lvwBView, "Borrowed", .lvwBView.ListItems(i).SubItems(4), i
    Next i
    SaveBTrans = "Transaction of Client " & .lblId.Caption & " has been saved successfully."
End With
End Function
'------------------END Borrow Modules-----------------------------------

'------------------Reserve Transaction Module---------------------------
Public Sub ClearRItems()
With frmTransactions
    .lblRCount.Caption = "---"
    .lblRLocation.Caption = "---"
    .lblRQty.Caption = "---"
    .lblRStatus.Caption = "---"
    .chkTest(0).Value = 0
    .chkTest(1).Value = 0
    .lblTrans.Caption = RcrdId("tblTransactions", Format(Date, "yymmdd"), "trans_no")
    ViewItemList "tblReserve", "item_id", "%", frmTransactions.lvwRItems, .lblId.Caption
End With
End Sub

Public Function SaveRTrans() As String
RcrdNo = 0
With frmTransactions
    SaveTrans
    RunSql "Select * from tblReserve order by record_no ASC"
    If Rs.EOF = False Then
        Rs.MoveLast
        RcrdNo = Rs.Fields!record_no
    End If
    For i = 1 To .lvwRView.ListItems.Count
        RunSql "Select * from tblReserve"
        Rs.AddNew
        Rs.Fields!record_no = RcrdNo + i
        Rs.Fields!client_no = .lblId.Caption
        Rs.Fields!item_id = .lvwRView.ListItems(i).Text
        Rs.Fields!Status = .lvwRView.ListItems(i).SubItems(3)
        Rs.Fields!qty = .lvwRView.ListItems(i).SubItems(1)
        Rs.Fields!reserve_date = .lvwRView.ListItems(i).SubItems(2)
        Rs.Fields!trans_no = .lblTrans.Caption
        Rs.Update
        UpdateStat .lvwRView, "Reserved", .lvwRView.ListItems(i).SubItems(3), i
    Next i
    SaveRTrans = "Transaction of Client " & .lblId.Caption & " has been saved successfully."
End With
End Function
'------------------END Reserve Transactions Modules---------------------

'------------------Return Transaction Modules---------------------------
Public Sub ClearBorrowed()
With frmTransactions
    .lblBDate.Caption = "---"
    .lblRDate.Caption = "---"
    .lblBTrans.Caption = "---"
    .lblQty.Caption = "---"
    .chkSvReturn.Value = Val(ReadINI("Preferences", "Save Trans"))
    .lblBrange.Caption = "---"
    .lblTrans.Caption = RcrdId("tblTransactions", Format(Date, "yymmdd"), "trans_no")
    ViewBorrowed "description", "%", .lvwBorrowed, .lblId.Caption
End With
End Sub
Public Sub ViewBorrowed(RcrdFld As String, RcrdStr As String, lvTrans As ListView, ClientNo As String)
RunSql "SELECT bar.record_no, bar.item_id, reg.description, bar.qty, bar.status " & _
        "FROM tblBorrow as bar INNER JOIN tblRegistered as reg ON bar.item_id = reg.item_id " & _
        "WHERE bar.client_no = " & Val(ClientNo) & " and reg." & RcrdFld & " LIKE '" & RcrdStr & "%'"
With Rs
    lvTrans.ListItems.Clear
    While Not .EOF = True
        Set x = lvTrans.ListItems.Add(, , .Fields(0), , 1)
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Public Function SaveReturn() As String
RcrdNo = 0
With frmTransactions
    SaveTrans
    RunSql "Select * from tblReturn order by record_no ASC"
    If Rs.EOF = False Then
        Rs.MoveLast
        RcrdNo = Rs.Fields!record_no
    End If
    For i = 1 To .lvwReturned.ListItems.Count
        RunSql "Select * from tblReturn"
        Rs.AddNew
        Rs.Fields!record_no = RcrdNo + i
        Rs.Fields!client_no = .lblId.Caption
        Rs.Fields!item_id = .lvwReturned.ListItems(i).SubItems(1)
        Rs.Fields!Status = .lvwReturned.ListItems(i).SubItems(3)
        Rs.Fields!qty = .lvwReturned.ListItems(i).SubItems(2)
        Rs.Fields!trans_no = .lblTrans.Caption
        Rs.Update
        UpdateStat .lvwReturned, "Returned", .lvwReturned.ListItems(i).SubItems(3), i
    Next i
    SaveReturn = "Transaction of Client " & .lblId.Caption & " has been saved successfully."
End With
End Function
'------------------END Return Transaction Modules-----------------------

'------------------Cancel Reservation Module--------------------------------
Public Sub ClearReserved()
With frmTransactions
    .lblTDate.Caption = "---"
    .lblReserveDate.Caption = "---"
    .lblCQty.Caption = "---"
    .txtCRemarks.Text = Empty
    .lblCTrans.Caption = "---"
    .chkSvCancel.Value = Val(ReadINI("Preferences", "Save Trans"))
    .lblTrans.Caption = RcrdId("tblTransactions", Format(Date, "yymmdd"), "trans_no")
    ViewReserved "description", "%", .lvwReserved, .lblId.Caption
End With
End Sub

Public Sub ViewReserved(RcrdFld As String, RcrdStr As String, lvTrans As ListView, ClientNo As String)
RunSql "SELECT res.record_no, res.item_id, reg.description, res.qty, res.status " & _
        "FROM tblReserve as res INNER JOIN tblRegistered as reg ON res.item_id = reg.item_id " & _
        "WHERE res.client_no = " & Val(ClientNo) & " and reg." & RcrdFld & " LIKE '" & RcrdStr & "%'"
With Rs
    lvTrans.ListItems.Clear
    While Not .EOF = True
        Set x = lvTrans.ListItems.Add(, , .Fields(0), , 1)
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Public Function SaveCancel() As String
RcrdNo = 0
With frmTransactions
    SaveTrans
    RunSql "Select * from tblCancel order by record_no ASC"
    If Rs.EOF = False Then
        Rs.MoveLast
        RcrdNo = Rs.Fields!record_no
    End If
    For i = 1 To .lvwCanceled.ListItems.Count
        RunSql "Select * from tblCancel"
        Rs.AddNew
        Rs.Fields!record_no = RcrdNo + i
        Rs.Fields!client_no = .lblId.Caption
        Rs.Fields!item_id = .lvwCanceled.ListItems(i).SubItems(1)
        Rs.Fields!Status = .lvwCanceled.ListItems(i).SubItems(3)
        Rs.Fields!qty = .lvwCanceled.ListItems(i).SubItems(2)
        Rs.Fields!remarks = .lvwCanceled.ListItems(i).SubItems(4)
        Rs.Fields!trans_no = .lblTrans.Caption
        Rs.Update
        UpdateStat .lvwCanceled, "Canceled", .lvwCanceled.ListItems(i).SubItems(3), i
    Next i
    SaveCancel = "Transaction of Client " & .lblId.Caption & " has been saved successfully."
End With
End Function
'------------------------END Reserve Module-----------------------------

'------------------Manage Transactions Module---------------------------
Public Sub ViewTransactions(ShowUnused As Integer, lstControl As ListBox, ClientNo As Long)
Dim lstIndex As Integer
lstControl.Clear
If ShowUnused = 0 Then
    RunSql "Select trans_no from tblTransactions where client_no = " & ClientNo
    With Rs
        While Not .EOF = True
            lstControl.AddItem .Fields(0)
            .MoveNext
        Wend
    End With
Else
    RunSql "Select * from tblTransactions"
    While Not Rs.EOF = True
        Select Case Rs.Fields!Transaction
            Case "Borrow"
                s = "Item Barrowing"
                Table = "tblBorrow"
            Case "Reserve"
                s = "Item Reservation"
                Table = "tblReserve"
            Case "Return"
                s = "Returned Items"
                Table = "tblReturn"
            Case "Cancel"
                s = "Canceled Reservations"
                Table = "tblCancel"
        End Select
        If Table = "tblReturn" Or Table = "tblCancel" Then
            lstControl.AddItem Rs.Fields!trans_no
        Else
            SubSql "SELECT tbl.* " & _
                    "FROM " & Table & " as tbl INNER JOIN tblTransactions as trans ON tbl.trans_no = trans.trans_no " & _
                    "WHERE trans.trans_no = " & Rs.Fields!trans_no
            With SubRs
                If .EOF = True Then
                    lstControl.AddItem Rs.Fields!trans_no
                End If
            End With
        End If
        Rs.MoveNext
    Wend
End If
End Sub

Public Sub ViewTrans(Table As String, TransNo As String, ClientNo As String, lvControl As ListView)
RunSql "Select * from " & Table & _
        " Where client_no = " & ClientNo & " and trans_no = " & TransNo
With Rs
    n = 0
    lvControl.ColumnHeaders.Clear
    For i = 1 To (.Fields.Count)
        lvControl.ColumnHeaders.Add
        If n < .Fields.Count Then
            lvControl.ColumnHeaders(i).Text = .Fields(n).Name
        End If
        n = n + 1
    Next i

    lvControl.ListItems.Clear
    While Not .EOF = True
        Set x = lvControl.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

'----------------------END Manage Transactions Module-------------------


'------------------General Functions------------------------------------

Public Sub ViewItemList(Table As String, RcrdFld As String, _
                        RcrdStr As String, lvTrans As ListView, _
                        ClientNo As String)
Dim IntBullet As Integer

RunSql "SELECT reg.item_id, reg.description, stat.qty, stat.status " & _
        "FROM tblStatus, tblRegistered AS reg INNER JOIN tblItemStatus AS stat ON reg.item_id = stat.item_id " & _
        "WHERE stat.qty <> 0 and tblStatus.Description = stat.Status And tblStatus.include = 1 And reg." & RcrdFld & " Like '" & RcrdStr & "%'" & _
        "ORDER BY reg.description ASC"
With Rs
    lvTrans.ListItems.Clear
    While Not .EOF = True
        SubSql "SELECT * from " & Table & " WHERE item_id = '" & .Fields(0) & "' and client_no = " & Val(ClientNo) & " and status = '" & .Fields!Status & "'"
        If SubRs.EOF = False Then
            Set x = lvTrans.ListItems.Add(, , .Fields(0), , 2)
        Else
            Set x = lvTrans.ListItems.Add(, , .Fields(0))
        End If
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Public Sub SaveTrans()
RunSql "Select * from tblTransactions"
With frmTransactions
    Rs.AddNew
    Rs.Fields!trans_no = .lblTrans.Caption
    Rs.Fields!Transaction = .tabTrans.SelectedItem.Key
    Rs.Fields!client_no = .lblId.Caption
    Rs.Fields!trans_date = Now
    Rs.Update
End With
End Sub

Public Sub UpdateStat(lvView As ListView, Transaction As String, Status As String, Index As Long)
Dim TStatus, TTable As String

If Transaction = "Returned" Then
    TStatus = "Borrowed"
    TTable = "tblBorrow"
ElseIf Transaction = "Canceled" Then
    TStatus = "Reserved"
    TTable = "tblReserve"
End If

If Transaction = "Returned" Or Transaction = "Canceled" Then
    RunSql "SELECT * from tblItemStatus where item_id = '" & lvView.ListItems(Index).SubItems(1) & "' " & _
            "and status = '" & TStatus & "'"
    With Rs
        If .EOF = False Then
            n = .Fields!qty - lvView.ListItems(Index).SubItems(2)
            If n <= 0 Then
                .Delete
            Else
                .Fields!qty = n
                .Update
            End If
        End If
    End With
    RunSql "SELECT * from tblItemStatus where item_id = '" & lvView.ListItems(Index).SubItems(1) & "' " & _
            "and status = '" & Status & "'"
    With Rs
        If .EOF = False Then
            .Fields!qty = .Fields!qty + lvView.ListItems(Index).SubItems(2)
            .Update
        End If
    End With
    RunSql "DELETE * from " & TTable & " where record_no = " & lvView.ListItems(Index).Text
    Exit Sub
End If

RunSql "Select * from tblItemStatus " & _
        "where item_id = '" & lvView.ListItems(Index).Text & "' and " & _
        "status = '" & Status & "'"
If Rs.EOF = False Then
    Rs.Fields!qty = Rs.Fields!qty - lvView.ListItems(Index).SubItems(1)
    Rs.Update
End If
SubSql "SELECT * from tblItemStatus " & _
        "WHERE item_id = '" & lvView.ListItems(Index).Text & "' and status = '" & Transaction & "'"
If SubRs.EOF = True Then
    SubRs.AddNew
    SubRs.Fields!qty = lvView.ListItems(Index).SubItems(1)
Else
    SubRs.Fields!qty = SubRs.Fields!qty + lvView.ListItems(Index).SubItems(1)
End If
SubRs.Fields!record_no = RcrdId("tblItemStatus", , "record_no")
SubRs.Fields!item_id = lvView.ListItems(Index).Text
SubRs.Fields!Status = Transaction
SubRs.Update
End Sub

Public Function ViewHistory(ClientNo As Integer, Transaction As String) As String
Dim TransStr As String

RunSql "Select * from tblTransactions where client_no = " & ClientNo & " and transaction = '" & Transaction & "'"
With Rs
    While Not .EOF = True
        TransStr = Empty
        Select Case Transaction
            Case "Borrow"
                SubSql "SELECT bar.*, reg.*, trans.* " & _
                        "FROM (tblBorrow as bar INNER JOIN tblRegistered as reg ON bar.item_id = reg.item_id) " & _
                        "INNER JOIN tblTransactions as trans ON bar.trans_no = trans.trans_no " & _
                        "WHERE bar.client_no = " & ClientNo & " and trans.trans_no = " & .Fields!trans_no
                If SubRs.EOF = True Then
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - System did not find any transaction matching this transaction number, " & .Fields!trans_no & _
                                    ". You may delete this record on the 'Manage Transactions' tab."
                Else
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - Borrowed the following..."
                End If
                While Not SubRs.EOF = True
                    TransStr = TransStr & vbNewLine & "   " & SubRs.Fields("bar.qty") & " " & SubRs.Fields!Description & "; Status: " & SubRs.Fields!Status & vbNewLine & "   Return Date: " & Scheduler(SubRs.Fields("trans_date"), SubRs.Fields("gap_val"), SubRs.Fields("interval"))
                    SubRs.MoveNext
                Wend
            Case "Reserve"
                SubSql "SELECT res.*, reg.*, trans.* " & _
                        "FROM (tblReserve as res INNER JOIN tblRegistered as reg ON res.item_id = reg.item_id) " & _
                        "INNER JOIN tblTransactions as trans ON res.trans_no = trans.trans_no " & _
                        "WHERE res.client_no = " & ClientNo & " and trans.trans_no = " & .Fields!trans_no
                If SubRs.EOF = True Then
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - System did not find any transaction matching this transaction number, " & .Fields!trans_no & _
                                    ". You may delete this record on the 'Manage Transactions' tab."
                Else
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - Reserved the following..."
                End If
                While Not SubRs.EOF = True
                    TransStr = TransStr & vbNewLine & "   " & SubRs.Fields("res.qty") & " " & SubRs.Fields!Description & "; Status: " & SubRs.Fields!Status
                    SubRs.MoveNext
                Wend
            Case "Return"
                SubSql "SELECT re.*, reg.*, trans.* " & _
                        "FROM (tblReturn as re INNER JOIN tblRegistered as reg ON re.item_id = reg.item_id) " & _
                        "INNER JOIN tblTransactions as trans ON re.trans_no = trans.trans_no " & _
                        "WHERE re.client_no = " & ClientNo & " and trans.trans_no = " & .Fields!trans_no
                If SubRs.EOF = True Then
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - System did not find any transaction matching this transaction number, " & .Fields!trans_no & _
                                    ". You may delete this record on the 'Manage Transactions' tab."
                Else
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - Returned the following items..."
                End If
                While Not SubRs.EOF = True
                    TransStr = TransStr & vbNewLine & "   " & SubRs.Fields("re.qty") & " " & SubRs.Fields!Description & "; Status: " & SubRs.Fields!Status
                    SubRs.MoveNext
                Wend
            Case "Cancel"
                SubSql "SELECT ca.*, reg.*, trans.* " & _
                        "FROM (tblCancel as ca INNER JOIN tblRegistered as reg ON ca.item_id = reg.item_id) " & _
                        "INNER JOIN tblTransactions as trans ON ca.trans_no = trans.trans_no " & _
                        "WHERE ca.client_no = " & ClientNo & " and trans.trans_no = " & .Fields!trans_no
                If SubRs.EOF = True Then
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - System did not find any transaction matching this transaction number, " & .Fields!trans_no & _
                                    ". You may delete this record on the 'Manage Transactions' tab."
                Else
                    ViewHistory = ViewHistory & Format(.Fields!trans_date, "mm/dd/yyyy mm:nn ampm") & " - Canceled the following reservations..."
                End If
                While Not SubRs.EOF = True
                    TransStr = TransStr & vbNewLine & "   " & SubRs.Fields("ca.qty") & " " & SubRs.Fields!Description & "; Status: " & SubRs.Fields!Status & "; Remarks: " & SubRs.Fields!remarks
                    SubRs.MoveNext
                Wend
        End Select
        ViewHistory = ViewHistory & TransStr & vbNewLine & vbNewLine
        .MoveNext
    Wend
End With
End Function
