VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "KPS Hospital Pvt. Ltd."
   ClientHeight    =   7920
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10170
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnudoctor 
      Caption         =   "&Doctor Details"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu mnufee 
      Caption         =   "&Fees Details"
   End
   Begin VB.Menu mnuOP 
      Caption         =   "&Billing"
   End
   Begin VB.Menu mnudaybook 
      Caption         =   "Daybook"
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnubillreport 
         Caption         =   "Bill Report"
      End
      Begin VB.Menu mnudaybookreport 
         Caption         =   "Daybook "
      End
   End
   Begin VB.Menu mnubackup 
      Caption         =   "B&ackup"
   End
   Begin VB.Menu mnucalculator 
      Caption         =   "&Calculator"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<================================ Printer Code ===========================================>
Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
   pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
   pcWritten As Long) As Long
   
Private Sub mnubillreport_Click()
FrmBillPeriodRpt.Show
'Call connect
'a = MsgBox("Do you want to print the Stock", vbYesNo)
'If a = vbYes Then
'    stmt = "select itemcode, sum(qty) as qty,purchaserate,salesrate,itemname,expirydate from tbl_stock group by itemcode,itemname,purchaserage,salesrate,expirydate"
'    If rs.State = 1 Then rs.Close
'    rs.Open stmt, db, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then
'        '----------Notepad print------------------
'        Open App.Path & "\rptcurrentstock.txt" For Output As #1
'
'        Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
'        Print #1, Space(22) & "Ramakrishna Mission Vidyalaya"
'        Print #1, Space(28) & "Coimbatore - 641020"
'        Print #1, ""
'        Print #1, "Current Stock as on " & Format(Date, "DD/MM/YYYY") & Space(24) & "          Time:" & Time()
'        Print #1, "--------------------------------------------------------------------------------"
'        Print #1, "I. Code" & Space(2) & "Item Name " & Space(20) & Space(2) & "Exp. Dt " & Space(2) & "  Qty" & Space(2) & " Pur. Rate" & Space(2) & "Sales Rate"
'        Print #1, "--------------------------------------------------------------------------------"
'        tqty = 0
'        tprate = 0
'        tsrate = 0
'        While Not rs.EOF
'            tpr = Val(rs.Fields("qty")) * Val(rs.Fields("purchaserate"))
'            tsr = Val(rs.Fields("qty")) * Val(rs.Fields("salesrate"))
'            tqty = Val(tqty) + Val(rs.Fields("qty"))
'            tprate = Val(tprate) + Val(tpr)
'            tsrate = Val(tsrate) + Val(tsr)
'
'            icode = 7 - Len(rs.Fields("itemcode"))
'            iname = 30 - Len(Mid(rs.Fields("itemname"), 1, 30))
'            iexp = 8 - Len(IIf(IsNull(rs.Fields("expiredate")), "0", rs.Fields("expiredate")))
'            iqty = 5 - Len(Round(rs.Fields("qty")))
'            iprate = 10 - Len(Format(tpr, "0.00"))
'            israte = 10 - Len(Format(tsr, "0.00"))
'
'            If Not rs.Fields("qty") = 0 Then
'                Print #1, UCase(rs.Fields("itemcode")) & Space(icode) & Space(2) & UCase(Mid(rs.Fields("itemname"), 1, 30)) & Space(iname) & Space(2) & rs.Fields("expiredate") & Space(iexp) & Space(2) & Space(iqty) & rs.Fields("qty") & Space(2) & Space(iprate) & Format(tpr, "0.00") & Space(2) & Space(israte) & Format(tsr, "0.00")
'            End If
'            rs.MoveNext
'        Wend
'        Print #1, "--------------------------------------------------------------------------------"
'        Print #1, Space(46) & Space(iqty) & tqty & Space(1) & Space(10 - Len(Format(tprate, "0.00"))) & Format(tprate, "0.00") & Space(1) & Space(10 - Len(Format(tsrate, "0.00"))) & Format(tsrate, "0.00")
'        Print #1, ""
'        Print #1, "Current Profit is " & Format(tsrate, "0.00") & " - " & Format(tprate, "0.00") & " = " & Val(Format(tsrate, "0.00")) - Val(Format(tprate, "0.00"))
'        Close #1
'        retval = Shell("notepad.exe rptcurrentstock.txt", vbMaximizedFocus)
'    End If
'
''Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
''Print #1, "TYPE rptcurrentstock.txt>PRN"
''Print #1, "EXIT"
''Close #1
''retval = Shell(App.Path & "\print.bat", vbHide)
'
'    '<==================== Printing Code ========================>
'    Dim lhPrinter As Long
'    Dim lReturn As Long
'    Dim lpcWritten As Long
'    Dim lDoc As Long
'    Dim sWrittenData As String
'    Dim MyDocInfo As DOCINFO
'    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
'    If lReturn = 0 Then
'        MsgBox "The Printer Name you typed wasn't recognized."
'        Exit Sub
'    End If
'    MyDocInfo.pDocName = "AAAAAA"
'    MyDocInfo.pOutputFile = vbNullString
'    MyDocInfo.pDatatype = vbNullString
'    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
'    Call StartPagePrinter(lhPrinter)
'
'    Dim var1 As String
'    Open App.Path & "\rptcurrentstock.txt" For Input As #1
'    var1 = Input(LOF(1), #1)
'    Close #1
'    sWrittenData = var1 '& vbFormFeed
'
'    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
'    Len(sWrittenData), lpcWritten)
'    lReturn = EndPagePrinter(lhPrinter)
'    lReturn = EndDocPrinter(lhPrinter)
'    lReturn = ClosePrinter(lhPrinter)
'    '<==================== Printing Code ========================>
'End If

End Sub

Private Sub mnudaybook_Click()
DBdaybookfrm.Show
End Sub

Private Sub mnudoctor_Click()
FrmDoctor.Show
End Sub

Private Sub mnufee_Click()
FrmFee.Show
End Sub

Private Sub mnubackup_Click()
'Open App.Path & "\bkp.bat" For Output As #1 '//Creating Batch file
'Print #1, "z:"
'Print #1, "cd RMV Medical"
'Print #1, "copy master.mdb d:\backup\"
'Close #1
'Call Shell(App.Path & "\bkp.bat", vbHide)
'MsgBox "Database Backuped Successfully"
End Sub

Private Sub mnuOP_Click()
FrmOP.Show
End Sub

Private Sub mnucalculator_Click()
Call Shell("calc.exe", vbNormalFocus)
End Sub

Private Sub mnuexit_Click()
End
End Sub
