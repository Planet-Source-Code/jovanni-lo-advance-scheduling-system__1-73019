Attribute VB_Name = "modCon"
Option Explicit

Public Con As New ADODB.Connection
Public Rs As New ADODB.Recordset
Public SubRs As New ADODB.Recordset

Public StrCon As String

Public Sub OpenCon(Path As String)
'open connection
Set Con = New ADODB.Connection
StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
        & Path & ";Persist Security Info=False"
Con.Open StrCon
If DbMgt = False Then
    dtaGroups.conDb.ConnectionString = StrCon
End If
End Sub

Public Function RunSql(Statement As String) As Boolean
On Error GoTo ErrMsg
Set Rs = New ADODB.Recordset
Rs.Open Statement, Con, adOpenKeyset, adLockPessimistic
RunSql = False
Exit Function
ErrMsg:
    RunSql = True
End Function

Public Sub SubSql(Statement As String)
On Error GoTo ErrMsg
Set SubRs = New ADODB.Recordset
SubRs.Open Statement, Con, adOpenKeyset, adLockPessimistic
Exit Sub
ErrMsg:
    MsgBox Statement & vbNewLine & vbNewLine & Err.Description, vbCritical
End Sub

Public Sub CloseRs(dtaRs As Variant)
If dtaRs.State = adStateOpen Then
    dtaRs.Close
End If
End Sub

Public Function CompactDB(pFileName As String) As Boolean
On Error GoTo ErrH
Dim CONN As New JRO.JetEngine
Dim ConnstringSorg As String, ConnstringDest As String

' Ensure file is not read only
SetAttr pFileName, vbNormal
ConnstringSorg = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
pFileName & ";User ID=;Password=;"
ConnstringDest = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
App.Path & "\Temp.mdb" & ";Jet OLEDB:Engine Type=5;"

Screen.MousePointer = vbHourglass
CONN.CompactDatabase ConnstringSorg, ConnstringDest
Screen.MousePointer = vbDefault

'Copy compacted file
Kill pFileName
FileCopy App.Path & "\Temp.mdb", pFileName
Kill App.Path & "\Temp.mdb"

Set CONN = Nothing
CompactDB = True
Exit Function
ErrH:
Screen.MousePointer = vbDefault
Debug.Print Err.Description
End Function



