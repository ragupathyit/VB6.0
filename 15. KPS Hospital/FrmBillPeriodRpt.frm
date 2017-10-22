VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBillPeriodRpt 
   BackColor       =   &H00FF0000&
   Caption         =   "Sales Period Wise Report"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   8400
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msgrid 
      Height          =   3375
      Left            =   480
      TabIndex        =   15
      Top             =   4920
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   $"FrmBillPeriodRpt.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdexcel 
      Caption         =   "&EXCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   14
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton CmdReport1 
      Caption         =   "&BILL WISE REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   12
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtfrom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cmbdoctor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      TabIndex        =   8
      Text            =   "Select the Doctor"
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   6
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "&DATE WISE REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94633985
      CurrentDate     =   40537
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94633985
      CurrentDate     =   40537
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No. To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   7320
      TabIndex        =   10
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No. From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   7320
      TabIndex        =   9
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Doctor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   2610
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES PERIOD WISE REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the To Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   2760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the From Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   3165
   End
End
Attribute VB_Name = "FrmBillPeriodRpt"
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
   
Private Sub cmdexcel_Click()
Dim xlapp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
Set xlapp = CreateObject("excel.application")
Set xlbook = xlapp.Workbooks.Add
Set xlsheet = xlbook.Worksheets(1)

xlsheet.Range("A1").EntireColumn.ColumnWidth = 8
xlsheet.Range("B1").EntireColumn.ColumnWidth = 12
xlsheet.Range("C1").EntireColumn.ColumnWidth = 25
xlsheet.Range("D1").EntireColumn.ColumnWidth = 25
xlsheet.Range("E1").EntireColumn.ColumnWidth = 10

xlsheet.Range("A1:E1").Font.Bold = True
xlsheet.Range("A1:E1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

Set i = Nothing
Set j = Nothing
For i = 0 To MSGrid.Rows - 1
    For j = 0 To MSGrid.Cols - 1
        xlsheet.Cells(i + 1, j + 1).Value = MSGrid.TextMatrix(i, j)
        
        If Trim(xlsheet.Cells(i + 1, j + 1).Value) = "Total" Then
            xlsheet.Range("A" & i + 1 & ":E" & i + 1).Font.Bold = True
        End If
        
        '--------------------Border---------------------------------------------------------
        xlsheet.Range("A" & i + 1 & ":E" & i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":E" & i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":E" & i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":E" & i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":E" & i + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------Border---------------------------------------------------------
        '--------------------Number Format 0.00-------------------------------------------
        xlsheet.Range("E" & i + 1).NumberFormat = "0.00"
        '--------------------Number Format 0.00-------------------------------------------
    Next j
Next i

xlapp.Application.Visible = True
End Sub

Private Sub CmdReport_Click()
If cmbdoctor.Text = "Select the Doctor" Then
    MsgBox "Select the doctor name properly...", vbInformation, "KPS Hospital"
    cmbdoctor.SetFocus
Else
    If cmbdoctor.Text = "All Doctors" Then
        If rs.State = 1 Then rs.Close
        rs.Open "Select * from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# ", db, adOpenDynamic, adLockOptimistic
        If rs.EOF Then
            MsgBox "No Records Check the Date", vbInformation, "KPS Hospital"
            Exit Sub
        End If
        
        If rs.State = 1 Then rs.Close
        
        rs.Open "Select distinct billno, payamt from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N'", db, adOpenDynamic, adLockOptimistic
        pamt = 0
        While Not rs.EOF
            pamt = Val(pamt) + Val(rs.Fields("payamt"))
            rs.MoveNext
        Wend
    
        stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N' order by doctorname,billno"
        If rs.State = 1 Then rs.Close
        rs.Open stmt, db, adOpenDynamic, adLockOptimistic
    Else    '================== Particular selected doctor report ==============================
        If rs.State = 1 Then rs.Close
        rs.Open "Select * from tbl_op where doctorname='" & Trim(cmbdoctor.Text) & "' and opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
        If rs.EOF Then
            MsgBox "No Records Check the Date", vbInformation, "KPS Hospital"
            Exit Sub
        End If
        
        If rs.State = 1 Then rs.Close
        
        rs.Open "Select distinct billno, payamt from tbl_op where doctorname='" & Trim(cmbdoctor.Text) & "' and opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N'", db, adOpenDynamic, adLockOptimistic
        pamt = 0
        While Not rs.EOF
            pamt = Val(pamt) + Val(rs.Fields("payamt"))
            rs.MoveNext
        Wend
    
        stmt = "select distinct billno, opdate, doctorname, patientname, totamt,cancel1 from tbl_op where doctorname='" & Trim(cmbdoctor.Text) & "' and opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N' order by billno"
        If rs.State = 1 Then rs.Close
        rs.Open stmt, db, adOpenDynamic, adLockOptimistic
    End If
    
    If Not rs.EOF Then
        '----------Notepad print------------------
        Open App.Path & "\rptopbillperiod.txt" For Output As #1
        
        Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
        Print #1, Space(13) & "K.P.S HOSPITALS (P) LTD"
        Print #1, Space(9) & "ANNUR ROAD, METTUPALAYAM - 641 301"
        Print #1, Space(16) & "04254-224314, 224315"
        Print #1, ""
        Print #1, Space(9) & "OP Bill Period Wise Report Details"
        Print #1, ""
        Print #1, "Report for the date from " & Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy")
        Print #1, "-------------------------------------------------------------"      ' 61
        Print #1, "B. No" & Space(2) & "Bill Date   " & Space(2) & "Patient Name" & Space(16) & Space(2) & "  Bill Amt"
        Print #1, "-------------------------------------------------------------"
        While Not rs.EOF
            ibilno = 5 - Len(rs.Fields("billno"))
            isadate = 12 - Len(Format(rs.Fields("opdate"), "DD/MM/YYYY"))
            ipname = 28 - Len(Trim(Mid(rs.Fields("patientname"), 1, 28)))   'Mid(rs.Fields("itemname"), 1, 30)
            itamt = 10 - Len(Format(rs.Fields("totamt"), "0.00"))
            
            Print #1, rs.Fields("billno") & Space(ibilno) & Space(2) & Format(rs.Fields("opdate"), "DD/MM/YYYY") & Space(isadate) & Space(2) & UCase(Trim(Mid(rs.Fields("patientname"), 1, 28))) & Space(ipname) & Space(2) & Space(itamt) & Format(rs.Fields("totamt"), "0.00")
            
            rs.MoveNext
        Wend
        Print #1, "-------------------------------------------------------------"
        
        ipamt = 8 - Len(Format(pamt, "0.00"))
        
        Print #1, Space(34) & "Total Amount (Rs): " & Space(ipamt) & Format(Val(pamt), "0.00")
        Print #1, ""
        '--------------------------------Canceled Bills-----------------------------------------
        stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='Y' order by doctorname,billno"
        If rs1.State = 1 Then rs1.Close
        rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            
            'MSGrid.AddItem "Canceled Bills"
            
            Print #1, "Canceled Bills"
            Print #1, "~~~~~~~~ ~~~~~"
            While Not rs1.EOF
                ibilno = 5 - Len(rs1.Fields("billno"))
                isadate = 12 - Len(Format(rs1.Fields("opdate"), "DD/MM/YYYY"))
                ipname = 28 - Len(Trim(Mid(rs1.Fields("patientname"), 1, 28)))   'Mid(rs.Fields("itemname"), 1, 30)
                itamt = 10 - Len(Format(rs1.Fields("totamt"), "0.00"))
                
                Print #1, rs1.Fields("billno") & Space(ibilno) & Space(2) & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & Space(isadate) & Space(2) & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & Space(ipname) & Space(2) & Space(itamt) & Format(rs1.Fields("totamt"), "0.00")
                'MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
                
                rs1.MoveNext
            Wend
        End If
        '--------------------------------Canceled Bills-----------------------------------------
        Close #1
        retval = Shell("notepad.exe rptopbillperiod.txt", vbMaximizedFocus)
        '----------Notepad print------------------
    End If
        
    a = MsgBox("Do you want to print the OP Bill Report", vbYesNo)
    If a = vbYes Then
        '<==================== Printing Code ========================>
        Dim lhPrinter As Long
        Dim lReturn As Long
        Dim lpcWritten As Long
        Dim lDoc As Long
        Dim sWrittenData As String
        Dim MyDocInfo As DOCINFO
        lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
        If lReturn = 0 Then
            MsgBox "The Printer Name you typed wasn't recognized."
            Exit Sub
        End If
        MyDocInfo.pDocName = "AAAAAA"
        MyDocInfo.pOutputFile = vbNullString
        MyDocInfo.pDatatype = vbNullString
        lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
        Call StartPagePrinter(lhPrinter)
    
        Dim var1 As String
        Open App.Path & "\rptopbillperiod.txt" For Input As #1
        var1 = Input(LOF(1), #1)
        Close #1
        sWrittenData = var1 '& vbFormFeed
    
        lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
        Len(sWrittenData), lpcWritten)
        lReturn = EndPagePrinter(lhPrinter)
        lReturn = EndDocPrinter(lhPrinter)
        lReturn = ClosePrinter(lhPrinter)
        '<==================== Printing Code ========================>
    End If
    
    MSGrid.Rows = 1
    If cmbdoctor.Text = "All Doctors" Then
        If rs.State = 1 Then rs.Close
        rs.Open "Select doctorname, sum(charges) as tamt from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N' group by doctorname", db, adOpenDynamic, adLockOptimistic
        While Not rs.EOF
            If rs1.State = 1 Then rs1.Close
            stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where doctorname='" & Trim(rs.Fields("doctorname")) & "' and opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N' order by doctorname,billno"
            rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
            While Not rs1.EOF
                MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
                rs1.MoveNext
            Wend
            MSGrid.AddItem "                                                    Total" & vbTab & vbTab & vbTab & vbTab & Format(rs.Fields("tamt"), "0.00")
            rs.MoveNext
        Wend
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "Select sum(charges) as tamt from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N' and doctorname='" & Trim(cmbdoctor.Text) & "'", db, adOpenDynamic, adLockOptimistic
        While Not rs.EOF
            If rs1.State = 1 Then rs1.Close
            stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where doctorname='" & Trim(cmbdoctor.Text) & "' and opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='N' order by doctorname,billno"
            rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
            While Not rs1.EOF
                MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
                rs1.MoveNext
            Wend
            MSGrid.AddItem "                                                    Total" & vbTab & vbTab & vbTab & vbTab & Format(rs.Fields("tamt"), "0.00")
            rs.MoveNext
        Wend
    End If
    '--------------------------------Canceled Bills-----------------------------------------
    stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='Y' order by doctorname,billno"
    If rs1.State = 1 Then rs1.Close
    rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        
        MSGrid.AddItem "Canceled Bills"
        
        While Not rs1.EOF
            MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
            rs1.MoveNext
        Wend
    End If
    '--------------------------------Canceled Bills-----------------------------------------
End If
End Sub

Private Sub CmdReport1_Click()
If rs.State = 1 Then rs.Close
rs.Open "Select * from tbl_op where billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " order by billno", db, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "No Records Check the Date", vbInformation, "KPS Hospital"
    Exit Sub
End If

If rs.State = 1 Then rs.Close

rs.Open "Select distinct billno, payamt from tbl_op where billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='N' order by billno", db, adOpenDynamic, adLockOptimistic
pamt = 0
While Not rs.EOF
    pamt = Val(pamt) + Val(rs.Fields("payamt"))
    rs.MoveNext
Wend

stmt = "select distinct billno, opdate, patientname, totamt,cancel1 from tbl_op where billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='N' order by billno"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\rptopbillno.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, Space(13) & "K.P.S HOSPITALS (P) LTD"
    Print #1, Space(9) & "ANNUR ROAD, METTUPALAYAM - 641 301"
    Print #1, Space(16) & "04254-224314, 224315"
    Print #1, ""
    Print #1, Space(9) & "OP Bill Period Wise Report Details"
    Print #1, ""
    Print #1, "Report for the date from " & Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy")
    Print #1, "-------------------------------------------------------------"      ' 61
    Print #1, "B. No" & Space(2) & "Bill Date   " & Space(2) & "Patient Name" & Space(16) & Space(2) & "  Bill Amt"
    Print #1, "-------------------------------------------------------------"
    While Not rs.EOF
        ibilno = 5 - Len(rs.Fields("billno"))
        isadate = 12 - Len(Format(rs.Fields("opdate"), "DD/MM/YYYY"))
        ipname = 28 - Len(Trim(Mid(rs.Fields("patientname"), 1, 28)))   'Mid(rs.Fields("itemname"), 1, 30)
        itamt = 10 - Len(Format(rs.Fields("totamt"), "0.00"))
        Print #1, rs.Fields("billno") & Space(ibilno) & Space(2) & Format(rs.Fields("opdate"), "DD/MM/YYYY") & Space(isadate) & Space(2) & UCase(Trim(Mid(rs.Fields("patientname"), 1, 28))) & Space(ipname) & Space(2) & Space(itamt) & Format(rs.Fields("totamt"), "0.00")
        rs.MoveNext
    Wend
    Print #1, "-------------------------------------------------------------"
    
    ipamt = 8 - Len(Format(pamt, "0.00"))
    
    Print #1, Space(34) & "Total Amount (Rs): " & Space(ipamt) & Format(Val(pamt), "0.00")
    Print #1, ""
    '--------------------------------Canceled Bills-----------------------------------------
    stmt = "select distinct billno, opdate, patientname, totamt, cancel1 from tbl_op where billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='Y' order by billno"
    If rs1.State = 1 Then rs1.Close
    rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        Print #1, "Canceled Bills"
        Print #1, "~~~~~~~~ ~~~~~"
        While Not rs1.EOF
            ibilno = 5 - Len(rs1.Fields("billno"))
            isadate = 12 - Len(Format(rs1.Fields("opdate"), "DD/MM/YYYY"))
            ipname = 28 - Len(Trim(Mid(rs1.Fields("patientname"), 1, 28)))   'Mid(rs.Fields("itemname"), 1, 30)
            itamt = 10 - Len(Format(rs1.Fields("totamt"), "0.00"))
            Print #1, rs1.Fields("billno") & Space(ibilno) & Space(2) & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & Space(isadate) & Space(2) & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & Space(ipname) & Space(2) & Space(itamt) & Format(rs1.Fields("totamt"), "0.00")
            rs1.MoveNext
        Wend
    End If
    '--------------------------------Canceled Bills-----------------------------------------
    Close #1
    retval = Shell("notepad.exe rptopbillno.txt", vbMaximizedFocus)
    '----------Notepad print------------------
End If

a = MsgBox("Do you want to print the OP Bill Report", vbYesNo)
If a = vbYes Then
    '<==================== Printing Code ========================>
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    MyDocInfo.pDocName = "AAAAAA"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)

    Dim var1 As String
    Open App.Path & "\rptopbillno.txt" For Input As #1
    var1 = Input(LOF(1), #1)
    Close #1
    sWrittenData = var1 '& vbFormFeed

    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    '<==================== Printing Code ========================>
End If
    
MSGrid.Rows = 1
If cmbdoctor.Text = "All Doctors" Then
    If rs.State = 1 Then rs.Close
    rs.Open "Select doctorname, sum(charges) as tamt from tbl_op where billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='N' group by doctorname", db, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        If rs1.State = 1 Then rs1.Close
        stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where doctorname='" & Trim(rs.Fields("doctorname")) & "' and billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='N' order by doctorname,billno"
        rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
            MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
            rs1.MoveNext
        Wend
        MSGrid.AddItem "                                                    Total" & vbTab & vbTab & vbTab & vbTab & Format(rs.Fields("tamt"), "0.00")
        rs.MoveNext
    Wend
Else
    If rs.State = 1 Then rs.Close
    rs.Open "Select sum(charges) as tamt from tbl_op where billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='N' and doctorname='" & Trim(cmbdoctor.Text) & "'", db, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        If rs1.State = 1 Then rs1.Close
        stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where doctorname='" & Trim(cmbdoctor.Text) & "' and billno>=" & Val(txtfrom.Text) & " and billno<=" & Val(txtto.Text) & " and cancel1='N' order by doctorname,billno"
        rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
            MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
            rs1.MoveNext
        Wend
        MSGrid.AddItem "                                                    Total" & vbTab & vbTab & vbTab & vbTab & Format(rs.Fields("tamt"), "0.00")
        rs.MoveNext
    Wend
End If
'--------------------------------Canceled Bills-----------------------------------------
stmt = "select distinct billno, opdate, patientname, totamt, doctorname, cancel1 from tbl_op where opdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and cancel1='Y' order by doctorname,billno"
If rs1.State = 1 Then rs1.Close
rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    
    MSGrid.AddItem "Canceled Bills"
    
    While Not rs1.EOF
        MSGrid.AddItem rs1.Fields("billno") & vbTab & Format(rs1.Fields("opdate"), "DD/MM/YYYY") & vbTab & UCase(Trim(rs1.Fields("doctorname"))) & vbTab & UCase(Trim(Mid(rs1.Fields("patientname"), 1, 28))) & vbTab & Format(rs1.Fields("totamt"), "0.00")    'Grid Fill Code
        rs1.MoveNext
    Wend
End If
'--------------------------------Canceled Bills-----------------------------------------
    
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call connect

cmbdoctor.AddItem "All Doctors"
If rs.State = 1 Then rs.Close
rs.Open "select doctorname from tbl_doctormaster order by dcode", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        cmbdoctor.AddItem rs.Fields("doctorname")
        rs.MoveNext
    Loop
End If

DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
