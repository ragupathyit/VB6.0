VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesPeriodRpt 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Period Wise Report"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   10290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4080
      TabIndex        =   7
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton CmdReport 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96927745
      CurrentDate     =   40537
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96927745
      CurrentDate     =   40537
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Customer"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   3210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6240
      TabIndex        =   5
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES PERIOD WISE REPORT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   5805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Report From"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   3630
   End
End
Attribute VB_Name = "SalesPeriodRpt"
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
   
Private Sub CmdReport_Click()
If rs.State = 1 Then rs.Close
If Combo1.Text <> "" Then
    rs.Open "Select * from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "Select * from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true", db, adOpenDynamic, adLockOptimistic
End If
If rs.EOF Then
    MsgBox "No Records Check the Date", vbInformation, "Sumathi Stores"
    Exit Sub
End If

If rs.State = 1 Then rs.Close
If Combo1.Text <> "" Then
    rs.Open "Select sum(amount) from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "Select sum(amount) from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true", db, adOpenDynamic, adLockOptimistic
End If
gamt = rs.Fields(0)

If rs.State = 1 Then rs.Close
If Combo1.Text <> "" Then
    rs.Open "Select sum(amount) from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=false", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "Select sum(amount) from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=false", db, adOpenDynamic, adLockOptimistic
End If
gamt_cancel = rs.Fields(0)

If rs.State = 1 Then rs.Close
If Combo1.Text <> "" Then
    rs.Open "Select distinct sum(discount) from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "Select distinct sum(discount) from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true", db, adOpenDynamic, adLockOptimistic
End If
gdisc = rs.Fields(0)

If Combo1.Text <> "" Then
    stmt = "select billno,sno,itemcode,itemname,itemrate,qty,qtytype,amount from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true order by billno,sno"
Else
    stmt = "select billno,sno,itemcode,itemname,itemrate,qty,qtytype,amount from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true order by billno,sno"
End If
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\rptsalesperiod.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, ""
    Print #1, Space(31) & "Sumathi Stores"
'    Print #1, Space(27) & "New Vegitable Market"
'    Print #1, Space(18) & "Vegitable Supplier and Commission Agency"
'    Print #1, Space(27) & "METTUPALAYAM - 641301"
'    Print #1, Space(21) & "CELL NO - 93641 33333, 90034 00000"
    Print #1, ""
    Print #1, Space(20) & "Sales Period Wise Report Details"
    If Combo1.Text <> "" Then
        Print #1, "Report for the Customer: " & Trim(Combo1.Text)
    End If
    Print #1, "Report for the date from " & Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy")
    Print #1, "--------------------------------------------------------------------------------"    ' 80 Characters
    Print #1, "B.No " & Space(1) & " S.No" & Space(1) & "ICode" & Space(1) & "Item Name  " & Space(20) & Space(1) & " Item Rate" & Space(1) & "Quantity" & Space(1) & " Total Amt"
    Print #1, "--------------------------------------------------------------------------------"
    While Not rs.EOF
        ibno = 5 - Len(rs.Fields("billno"))
        isno = 5 - Len(rs.Fields("sno"))
        iicode = 5 - Len(rs.Fields("itemcode"))
        iiname = 31 - Len(rs.Fields("itemname"))
        iirate = 10 - Len(Format(rs.Fields("itemrate"), "0.00"))
        iiqty = 8 - Len(rs.Fields("qty") & " " & rs.Fields("qtytype"))
        iiamt = 10 - Len(Format(rs.Fields("amount"), "0.00"))
        
        Print #1, rs.Fields("billno") & Space(ibno) & Space(1) & Space(isno) & rs.Fields("sno") & Space(1) & UCase(rs.Fields("itemcode")) & Space(iicode) & Space(1) & UCase(rs.Fields("itemname")) & Space(iiname) & Space(1) & Space(iirate) & Format(rs.Fields("itemrate"), "0.00") & Space(1) & rs.Fields("qty") & " " & rs.Fields("qtytype") & Space(iiqty) & Space(1) & Space(iiamt) & Format(rs.Fields("amount"), "0.00")
        rs.MoveNext
    Wend
    Print #1, "--------------------------------------------------------------------------------"
    
    igamt = 10 - Len(Format(gamt, "0.00"))
    igdisc = 10 - Len(Format(gdisc, "0.00"))
    itpa = 10 - Len(Format(Val(gamt) - Val(gdisc), "0.00"))
    
    Print #1, Space(51) & "Total Amount (Rs): " & Space(igamt) & Format(gamt, "0.00")
    Print #1, Space(51) & "Discount     (Rs): " & Space(igdisc) & Format(gdisc, "0.00")
    Print #1, Space(45) & "Total Sales Amount (Rs): " & Space(itpa) & Format(Val(gamt) - Val(gdisc), "0.00")
    Print #1, ""
    Print #1, ""
    rs.Close
    
    If Combo1.Text <> "" Then
        stmt = "select billno,sno,itemcode,itemname,itemrate,qty,qtytype,amount from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=false order by billno,sno"
    Else
        stmt = "select billno,sno,itemcode,itemname,itemrate,qty,qtytype,amount from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=false order by billno,sno"
    End If
    rs.Open stmt, db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        Print #1, "Canceled Bills"
        Print #1, "~~~~~~~~ ~~~~~"
        Print #1, "--------------------------------------------------------------------------------"    ' 80 Characters
        Print #1, "B.No " & Space(1) & " S.No" & Space(1) & "ICode" & Space(1) & "Item Name  " & Space(20) & Space(1) & " Item Rate" & Space(1) & "Quantity" & Space(1) & " Total Amt"
        Print #1, "--------------------------------------------------------------------------------"
        While Not rs.EOF
            ibno = 5 - Len(rs.Fields("billno"))
            isno = 5 - Len(rs.Fields("sno"))
            iicode = 5 - Len(rs.Fields("itemcode"))
            iiname = 31 - Len(rs.Fields("itemname"))
            iiqty = 8 - Len(rs.Fields("qty") & " " & rs.Fields("qtytype"))
            iirate = 10 - Len(Format(rs.Fields("itemrate"), "0.00"))
            iiamt = 10 - Len(Format(rs.Fields("amount"), "0.00"))
        
            Print #1, rs.Fields("billno") & Space(ibno) & Space(1) & Space(isno) & rs.Fields("sno") & Space(1) & UCase(rs.Fields("itemcode")) & Space(iicode) & Space(1) & UCase(rs.Fields("itemname")) & Space(iiname) & Space(1) & Space(iirate) & Format(rs.Fields("itemrate"), "0.00") & Space(1) & rs.Fields("qty") & " " & rs.Fields("qtytype") & Space(iiqty) & Space(1) & Space(iiamt) & Format(rs.Fields("amount"), "0.00")
            rs.MoveNext
        Wend
        Print #1, "--------------------------------------------------------------------------------"
        igamt_cancel = 10 - Len(Format(gamt_cancel, "0.00"))
        Print #1, Space(51) & "Total Amount (Rs): " & Space(igamt_cancel) & Format(gamt_cancel, "0.00")
        Print #1, ""
        Print #1, ""
    End If
    rs.Close
    Close #1
    retval = Shell("notepad.exe rptsalesperiod.txt", vbMaximizedFocus)
End If

'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'Print #1, "TYPE rptsalesperiod.txt>PRN"
'Print #1, "EXIT"
'Close #1
'retval = Shell(App.Path & "\print.bat", vbHide)
    
s = MsgBox("Do You Want Print", vbYesNo)
If s = vbYes Then
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
    Open App.Path & "\rptsalesperiod.txt" For Input As #1
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

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(35, 29, 29)
Call connect

If rs.State = 1 Then rs.Close
rs.Open "select customername from tbl_custmaster order by customername", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    While Not rs.EOF
        Combo1.AddItem rs.Fields("customername")
        rs.MoveNext
    Wend
End If
rs.Close

DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
