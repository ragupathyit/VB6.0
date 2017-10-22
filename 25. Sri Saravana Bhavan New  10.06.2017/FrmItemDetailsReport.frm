VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmItemDetailsReport 
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
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   6960
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
      MICON           =   "FrmItemDetailsReport.frx":0000
      PICN            =   "FrmItemDetailsReport.frx":001C
      PICH            =   "FrmItemDetailsReport.frx":072E
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
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "SAVE"
      Top             =   7560
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
      MICON           =   "FrmItemDetailsReport.frx":0E40
      PICN            =   "FrmItemDetailsReport.frx":0E5C
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
      Height          =   6255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "I. Code  |Item Name                            |Price    |Item Type    "
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
   Begin Project1.Button BtnReport 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "SAVE"
      Top             =   7560
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
      MICON           =   "FrmItemDetailsReport.frx":156E
      PICN            =   "FrmItemDetailsReport.frx":158A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmItemDetailsReport.frx":1C9C
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM DETAILS REPORT"
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
      Width           =   3990
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   7320
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "FrmItemDetailsReport"
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

xlsheet.Range("A1").EntireColumn.ColumnWidth = 10
xlsheet.Range("B1").EntireColumn.ColumnWidth = 24
xlsheet.Range("C1").EntireColumn.ColumnWidth = 10
xlsheet.Range("D1").EntireColumn.ColumnWidth = 12

xlsheet.Range("A1:D1").Font.Bold = True
xlsheet.Range("A1:D1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

Set i = Nothing
Set j = Nothing
For i = 0 To MSGrid.Rows - 1
    For j = 0 To MSGrid.Cols - 1
        xlsheet.Cells(i + 1, j + 1).Value = MSGrid.TextMatrix(i, j)
        
        If Trim(xlsheet.Cells(i + 1, j + 1).Value) = "Total" Then
            'xlsheet.Range("A" & i + 1 & ":D" & j + 1).Merge
            'xlsheet.Range("A" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            xlsheet.Range("A" & i + 1 & ":D" & i + 1).Font.Bold = True
        End If
        
        '--------------------Border---------------------------------------------------------
        xlsheet.Range("A" & i + 1 & ":D" & i + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":D" & i + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":D" & i + 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":D" & i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 1 & ":D" & i + 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------Border---------------------------------------------------------
        '--------------------Number Format 0.00-------------------------------------------
        xlsheet.Range("C" & i + 1).NumberFormat = "0.00"
        '--------------------Number Format 0.00-------------------------------------------
    Next j
Next i

'If Trim(xlsheet.Cells(i, j).Value) = "Total" Then
'    xlsheet.Range("A" & i & ":C" & j).Merge
'    'xlsheet.Range("A" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'    xlsheet.Range("A" & i & ":D" & i).Font.Bold = True
'End If
        
xlapp.Application.Visible = True
End Sub

'Private Sub BtnOK_Click()
'If rs.State = 1 Then rs.Close
'rs.Open "select distinct billno, payamt from tbl_order where iscomplete=true and orderdate between #" & Format(DTP_from.Value, "m/d/yyyy") & "# and #" & Format(DTP_to.Value, "m/d/yyyy") & "#", db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    pamt = 0
'    While Not rs.EOF
'        pamt = Val(pamt) + Val(rs.Fields("payamt"))
'        rs.MoveNext
'    Wend
'End If
'
'If rs.State = 1 Then rs.Close
'rs.Open "select distinct billno, orderdate, orderid, totqty, payamt from tbl_order where iscomplete=true and orderdate between #" & Format(DTP_from.Value, "m/d/yyyy") & "# and #" & Format(DTP_to.Value, "m/d/yyyy") & "# order by billno", db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    MSGrid.Rows = 1
'    While Not rs.EOF
'        MSGrid.AddItem rs.Fields("billno") & vbTab & Format(rs.Fields("orderdate"), "dd/mm/yyyy") & vbTab & rs.Fields("orderid") & vbTab & rs.Fields("totqty") & vbTab & Format(rs.Fields("payamt"), "0.00")
'        rs.MoveNext
'    Wend
'    MSGrid.AddItem "Total" & vbTab & vbTab & vbTab & vbTab & Format(pamt, "0.00")
'Else
'    MsgBox "No Bills Found. Check the Date", vbInformation, "Sri Saravana Bhavan"
'End If
'End Sub

Private Sub BtnReport_Click()
'----------Notepad print------------------
Open App.Path & "\itemreport.txt" For Output As #1
'Print #1, Chr(18); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
Print #1, "          Sri Saravana Bhavan"
Print #1, ""
Print #1, "          Item Details Report"
Print #1, ""
Print #1, "------------------------------------------"      '42 characters
Print #1, "ICo. Item Name               Price I.Type "
Print #1, "------------------------------------------"
For i = 1 To MSGrid.Rows - 2
    iicode = 4 - Len(MSGrid.TextMatrix(i, 0))
    iiname = 21 - Len(Mid(MSGrid.TextMatrix(i, 1), 1, 21))
    iprice = 7 - Len(Format(Val(MSGrid.TextMatrix(i, 2)), "0.00"))
    iitype = 7 - Len(MSGrid.TextMatrix(i, 3))
    
    Print #1, Space(iicode) & MSGrid.TextMatrix(i, 0) & Space(1) & Mid(MSGrid.TextMatrix(i, 1), 1, 21) & Space(iiname) & Space(1) & Space(iprice) & Format(Val(MSGrid.TextMatrix(i, 2)), "0.00") & Space(1) & MSGrid.TextMatrix(i, 3) & Space(1) & Space(iitype)
Next i
Print #1, "------------------------------------------"
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Chr$(&H1B); "m"; Chr$(&HA);
Close #1
retval = Shell("notepad.exe itemreport.txt", vbHide)

'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'Print #1, "start DOSPrinter.exe /ESC E /F'Lucida Console' salesreport.txt"
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
Open App.Path & "\itemreport.txt" For Input As #1
var1 = Input(LOF(1), #1)
Close #1

sWrittenData = var1 '& vbFormFeed

lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
Len(sWrittenData), lpcWritten)
lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)
'<==================== Printing Code ========================>
    
End Sub

Private Sub Form_Load()
Call connect

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

MSGrid.Rows = 1

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_itemmaster order by itemcode", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    MSGrid.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & Format(Val(rs.Fields("price")), "0.00") & vbTab & rs.Fields("itemtype")
    rs.MoveNext
Wend
rs.Close

End Sub
