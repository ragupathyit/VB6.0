VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesPeriodRpt 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Retail Sales Period Wise Report"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3840
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
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   109379585
      CurrentDate     =   43019
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   109379585
      CurrentDate     =   43019
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   3600
      Width           =   8415
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES PERIOD WISE REPORT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   1680
      TabIndex        =   5
      Top             =   200
      Width           =   5070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the To Date to Take Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   3810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the From Date to Take Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   4110
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   8415
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
rs.Open "Select * from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "No Records Check the Date", vbInformation, "Rotary Club of Mettupalayam"
    Exit Sub
End If

If rs.State = 1 Then rs.Close
'Debug.Print "Select distinct billno,totamt from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
rs.Open "Select distinct billno,totamt from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
gamt = 0
While Not rs.EOF
    gamt = Val(gamt) + Val(rs.Fields("totamt"))
    rs.MoveNext
Wend

If rs.State = 1 Then rs.Close
rs.Open "Select distinct billno,discount from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
dis = 0
While Not rs.EOF
    dis = Val(dis) + Val(rs.Fields("discount"))
    rs.MoveNext
Wend

stmt = "select distinct billno,salesdate,totamt from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\rptrsalesperiod.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, ""
    Print #1, Space(2) & "Rotary Club of Mettupalayam"
    Print #1, ""
    Print #1, Space(0) & "Sales Period Wise Report Details"
    Print #1, "Report from " & Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy")
    Print #1, "-----------------------------------"    ' 35 Characters
    Print #1, "Bill No" & Space(2) & "Sales Date" & Space(2); "  Bill Amt"
    Print #1, "-----------------------------------"
    While Not rs.EOF
        iibillno = 7 - Len(rs.Fields("billno"))
        iisalesdate = 10 - Len(rs.Fields("salesdate"))
        iiamt = 10 - Len(Format(rs.Fields("totamt"), "0.00"))
        
        Print #1, UCase(rs.Fields("billno")) & Space(iibillno) & Space(2) & UCase(rs.Fields("salesdate")) & Space(iisalesdate) & Space(2) & Space(iiamt) & Format(rs.Fields("totamt"), "0.00")
        rs.MoveNext
    Wend
    Print #1, "-----------------------------------"
    
    igamt = 10 - Len(Format(gamt, "0.00"))
    idis = 10 - Len(Format(dis, "0.00"))
    itpa = 10 - Len(Format(Val(gamt) - Val(dis), "0.00"))
    
    Print #1, Space(2) & "Total Amount (Rs): " & Space(igamt) & Format(gamt, "0.00")
    Print #1, Space(2) & "Discount Amt (Rs): " & Space(idis) & Format(dis, "0.00")
    Print #1, Space(2) & "Total Amount (Rs): " & Space(itpa) & Format(Val(gamt) - Val(dis), "0.00")
    Close #1
    retval = Shell("notepad.exe rptrsalesperiod.txt", vbMaximizedFocus)
End If

s = MsgBox("Do You Want Print", vbYesNo)
If s = vbYes Then
    'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
    'Print #1, "TYPE rptrsalesperiod.txt>PRN"
    'Print #1, "EXIT"
    'Close #1
    'retval = Shell(App.Path & "\print.bat", vbHide)
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
    Open App.Path & "\rptrsalesperiod.txt" For Input As #1
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

'Me.BackColor = RGB(255, 204, 203)

Call connect
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
