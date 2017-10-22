Attribute VB_Name = "ModuleConnection"
Global db As New ADODB.Connection
Global rs As New ADODB.Recordset
Global rs1 As New ADODB.Recordset
Global rs2 As New ADODB.Recordset
Global sysuser As String

Public Sub connect()
If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"
sysuser = "Machine-1"
End Sub
