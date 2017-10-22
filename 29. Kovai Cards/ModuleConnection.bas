Attribute VB_Name = "ModuleConnection"
Global db As New ADODB.Connection
Global rs As New ADODB.Recordset
Global rs1 As New ADODB.Recordset
Global rs2 As New ADODB.Recordset
Global serialno As String

'Global dele As String

Public Sub connect()
If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"
End Sub

Public Sub generate_sno(a As String)
If rs.State = 1 Then rs.Close
rs.Open "select billno from tbl_cashsales where billno like'" & a & "%' order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    
    If Mid(rs.Fields("billno"), 8, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 9, 1)) + 1
        If Val(serialno) > 9 Then
            serialno = a & "00000" & serialno
        Else
            serialno = a & "000000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 7, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 8, 2)) + 1
        If Val(serialno) > 99 Then
            serialno = a & "0000" & serialno
        Else
            serialno = a & "00000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 6, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 7, 3)) + 1
        If Val(serialno) > 999 Then
            serialno = a & "000" & serialno
        Else
            serialno = a & "0000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 5, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 6, 4)) + 1
        If Val(serialno) > 9999 Then
            serialno = a & "00" & serialno
        Else
            serialno = a & "000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 4, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 5, 5)) + 1
        If Val(serialno) > 99999 Then
            serialno = a & "0" & serialno
        Else
            serialno = a & "00" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 3, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 4, 6)) + 1
        If Val(serialno) > 99999 Then
            serialno = a & serialno
        Else
            serialno = a & "0" & serialno
        End If
    End If
Else
    serialno = a & "000001"
End If
rs.Close
End Sub

Public Sub generate_sno1(a As String)
If rs.State = 1 Then rs.Close
rs.Open "select billno from tbl_creditsales where billno like'" & a & "%' order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    
    If Mid(rs.Fields("billno"), 8, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 9, 1)) + 1
        If Val(serialno) > 9 Then
            serialno = a & "0000" & serialno
        Else
            serialno = a & "00000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 7, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 8, 2)) + 1
        If Val(serialno) > 99 Then
            serialno = a & "000" & serialno
        Else
            serialno = a & "0000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 6, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 7, 3)) + 1
        If Val(serialno) > 999 Then
            serialno = a & "00" & serialno
        Else
            serialno = a & "000" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 5, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 6, 4)) + 1
        If Val(serialno) > 9999 Then
            serialno = a & "0" & serialno
        Else
            serialno = a & "00" & serialno
        End If
    ElseIf Mid(rs.Fields("billno"), 4, 1) = "0" Then
        serialno = Val(Mid(rs.Fields("billno"), 5, 5)) + 1
        If Val(serialno) > 99999 Then
            serialno = a & serialno
        Else
            serialno = a & "0" & serialno
        End If
'    ElseIf Mid(rs.Fields("billno"), 3, 1) = "0" Then
'        serialno = Val(Mid(rs.Fields("billno"), 4, 6)) + 1
'        If Val(serialno) > 99999 Then
'            serialno = a & serialno
'        Else
'            serialno = a & "0" & serialno
'        End If
    End If
Else
    serialno = a & "000001"
End If
rs.Close
End Sub
