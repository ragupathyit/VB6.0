VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSalesReport 
   BackColor       =   &H00400040&
   Caption         =   "Reports"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12750
   WindowState     =   2  'Maximized
   Begin Project1.Button BtnOK 
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSalesReport.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSalesReport.frx":001C
      PICN            =   "FrmSalesReport.frx":0038
      PICH            =   "FrmSalesReport.frx":074A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnExcel 
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "SAVE"
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Excel  "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSalesReport.frx":0E5C
      PICN            =   "FrmSalesReport.frx":0E78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Bill No  |Bill Date      |Qty    |Tax Amt         |Bill Amt                "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTP_from 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   97648643
      CurrentDate     =   42428
   End
   Begin Project1.Button BtnReport 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "SAVE"
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Report"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSalesReport.frx":158A
      PICN            =   "FrmSalesReport.frx":15A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox cmb_bnot 
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   1560
      Width           =   855
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1508;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3120
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin MSForms.ComboBox cmb_bnof 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   855
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1508;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "1"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmSalesReport.frx":1CB8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAILY SALES REPORT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3780
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   7560
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "FrmSalesReport"
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
   
Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub BtnExcel_Click()
Dim xlapp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
Set xlapp = CreateObject("excel.application")
Set xlbook = xlapp.Workbooks.Add
Set xlsheet = xlbook.Worksheets(1)

xlsheet.Range("A1").EntireColumn.ColumnWidth = 8
xlsheet.Range("B1").EntireColumn.ColumnWidth = 12
xlsheet.Range("C1").EntireColumn.ColumnWidth = 12
xlsheet.Range("D1").EntireColumn.ColumnWidth = 8
xlsheet.Range("E1").EntireColumn.ColumnWidth = 14

xlsheet.Range("A1:E1").Font.Bold = True
xlsheet.Range("A1:E1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

Set i = Nothing
Set j = Nothing
For i = 0 To MSGrid.Rows - 1
    For j = 0 To MSGrid.Cols - 1
        xlsheet.Cells(i + 1, j + 1).Value = MSGrid.TextMatrix(i, j)
        
        If Trim(xlsheet.Cells(i + 1, j + 1).Value) = "Total" Then
            'xlsheet.Range("A" & i + 1 & ":D" & j + 1).Merge
            'xlsheet.Range("A" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
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
        xlsheet.Range("D" & i + 1).NumberFormat = "0.00"
        xlsheet.Range("E" & i + 1).NumberFormat = "0.00"
        '--------------------Number Format 0.00-------------------------------------------
    Next j
Next i

If Trim(xlsheet.Cells(i, j).Value) = "Total" Then
    xlsheet.Range("A" & i & ":C" & j).Merge
    'xlsheet.Range("A" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
    xlsheet.Range("A" & i & ":D" & i).Font.Bold = True
    xlsheet.Range("A" & i & ":E" & i).Font.Bold = True
End If
        
xlapp.Application.Visible = True
End Sub

Private Sub BtnOK_Click()
'If rs.State = 1 Then rs.Close
'If cmb_bnot.Text = "" Then
'    rs.Open "select distinct vattax, payamt from tbl_order where iscomplete=true and orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "#", db, adOpenDynamic, adLockOptimistic
'Else
'    rs.Open "select distinct vattax, payamt from tbl_order where iscomplete=true and orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "# and billno>=" & Val(Trim(cmb_bnof.Value)) & " and billno<=" & Val(Trim(cmb_bnot.Value)), db, adOpenDynamic, adLockOptimistic
'End If
'
'If Not rs.EOF Then
'    pamt = 0
'    vtax = 0
'    While Not rs.EOF
'        pamt = Val(pamt) + Val(rs.Fields("payamt"))
'        vtax = Val(vtax) + IIf(IsNull(rs.Fields("vattax")), "0", rs.Fields("vattax"))
'        rs.MoveNext
'    Wend
'End If

If rs.State = 1 Then rs.Close
If cmb_bnot.Text = "" Then
    rs.Open "select distinct billno, orderdate, totqty, vattax, payamt from tbl_order where iscomplete=true and orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "# order by billno", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "select distinct billno, orderdate, totqty, vattax, payamt from tbl_order where iscomplete=true and orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "# and billno>=" & Val(Trim(cmb_bnof.Value)) & " and billno<=" & Val(Trim(cmb_bnot.Value)), db, adOpenDynamic, adLockOptimistic
End If

If Not rs.EOF Then
    MSGrid.Rows = 1
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("billno") & vbTab & Format(rs.Fields("orderdate"), "d/m/yyyy") & vbTab & rs.Fields("totqty") & vbTab & Format(rs.Fields("vattax"), "0.00") & vbTab & Format(rs.Fields("payamt"), "0.00")
        rs.MoveNext
    Wend
Else
    MsgBox "No Bills Found. Check the Date", vbInformation, "Sri Saravana Bhavan"
End If

vtax = 0
pamt = 0
For i = 1 To MSGrid.Rows - 1
    'vtax = Val(vtax) + Val(MSGrid.TextMatrix(i, 3))
    pamt = Val(pamt) + Val(MSGrid.TextMatrix(i, 4))
Next i

MSGrid.AddItem "Total" & vbTab & vbTab & vbTab & Format(vtax, "0.00") & vbTab & Format(pamt, "0.00")
End Sub

Private Sub BtnReport_Click()
If rs.State = 1 Then rs.Close
If cmb_bnot.Text = "" Then
    rs.Open "select distinct vattax, payamt from tbl_order where iscomplete=true and orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "#", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "select distinct vattax, payamt from tbl_order where iscomplete=true and orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "# and billno>=" & Val(Trim(cmb_bnof.Value)) & " and billno<=" & Val(Trim(cmb_bnot.Value)), db, adOpenDynamic, adLockOptimistic
End If

If Not rs.EOF Then
    pamt = 0
    vtax = 0
    While Not rs.EOF
        pamt = Val(pamt) + Val(rs.Fields("payamt"))
        'vtax = Val(vtax) + Val(rs.Fields("vattax"))
        vtax = Val(vtax) + IIf(IsNull(rs.Fields("vattax")), "0", rs.Fields("vattax"))
        rs.MoveNext
    Wend
End If

'----------Notepad print------------------
Open App.Path & "\salesreport.txt" For Output As #1
'Print #1, Chr(18); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
Print #1, "             Sales Report"
Print #1, Space(9) & "Sri Saravana Bhavan"
Print #1, ""
Print #1, "Sales Report Date - " & Format(DTP_from.Value, "dd/mm/yyyy")
Print #1, "------------------------------------------"      '42 characters
Print #1, "B. No Bill Date  Qty    Vat Tax   Bill Amt"
Print #1, "------------------------------------------"
For i = 1 To MSGrid.Rows - 2
    ibno = 5 - Len(MSGrid.TextMatrix(i, 0))
    ibdate = 10 - Len(Format(MSGrid.TextMatrix(i, 1), "dd/mm/yyyy"))
    iqty = 4 - Len(MSGrid.TextMatrix(i, 2))
    ivtax = 10 - Len(Format(Val(MSGrid.TextMatrix(i, 3)), "0.00"))
    ibamt = 10 - Len(Format(Val(MSGrid.TextMatrix(i, 4)), "0.00"))
    
    Print #1, Space(ibno) & MSGrid.TextMatrix(i, 0) & Space(1) & Space(ibdate) & Format(MSGrid.TextMatrix(i, 1), "dd/mm/yyyy") & Space(iqty) & MSGrid.TextMatrix(i, 2) & Space(1) & Space(ivtax) & Format(Val(MSGrid.TextMatrix(i, 3)), "0.00") & Space(1) & Space(ibamt) & Format(Val(MSGrid.TextMatrix(i, 4)), "0.00")
Next i
Print #1, "------------------------------------------"
Print #1, "             Total: " & Space(11 - Len(Format(Val(vtax), "0.00"))) & Format(Val(vtax), "0.00") & Space(11 - Len(Format(Val(pamt), "0.00"))) & Format(Val(pamt), "0.00")
Print #1, "                   -----------------------"
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Chr$(&H1B); "m"; Chr$(&HA);
Close #1
retval = Shell("notepad.exe salesreport.txt", vbHide)

If MsgBox("Are you sure to take daily sales report", vbYesNo, "Sri Saravana Bhavan") = vbYes Then
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
    Open App.Path & "\salesreport.txt" For Input As #1
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

Private Sub DTP_from_Change()
'MsgBox DTP_from.Value
cmb_bnof.Clear
cmb_bnot.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "select max(billno) as bno from tbl_order where orderdate=#" & Format(DTP_from.Value, "m/d/yyyy") & "#", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    For i = 1 To Val(rs1.Fields("bno"))
        cmb_bnof.AddItem i
        cmb_bnot.AddItem i
    Next
Else
    cmb_bnot.AddItem "1"
End If
rs1.Close
cmb_bnof.Text = 1
End Sub

Private Sub Form_Load()
Call connect

DTP_from.Value = Date
'DTP_to.Value = Date
cmb_bnof.Clear
cmb_bnot.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "select billno from tbl_tempbill order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    rs1.MoveLast
    For i = 1 To Val(rs1.Fields("billno"))
        cmb_bnof.AddItem i
        cmb_bnot.AddItem i
    Next
Else
    cmb_bnot.AddItem "1"
End If
rs1.Close

cmb_bnof.Text = 1

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i
End Sub
