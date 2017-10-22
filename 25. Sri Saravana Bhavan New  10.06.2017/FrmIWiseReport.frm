VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmIWiseReport 
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
      Left            =   5280
      TabIndex        =   6
      Top             =   1800
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
      MICON           =   "FrmIWiseReport.frx":0000
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
      Left            =   7320
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
      MICON           =   "FrmIWiseReport.frx":001C
      PICN            =   "FrmIWiseReport.frx":0038
      PICH            =   "FrmIWiseReport.frx":074A
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
      MICON           =   "FrmIWiseReport.frx":0E5C
      PICN            =   "FrmIWiseReport.frx":0E78
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
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "S.No  |Item Name                                  |I.Rate   |Qty     |Amt         "
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
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
      Format          =   97845251
      CurrentDate     =   42428
   End
   Begin MSComCtl2.DTPicker DTP_to 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
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
      Format          =   97845251
      CurrentDate     =   42428
   End
   Begin Project1.Button BtnReport 
      Height          =   495
      Left            =   2520
      TabIndex        =   9
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
      MICON           =   "FrmIWiseReport.frx":158A
      PICN            =   "FrmIWiseReport.frx":15A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report From"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report To"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmIWiseReport.frx":1CB8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM WISE REPORT"
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
      Width           =   3450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   7320
      Width           =   7935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "FrmIWiseReport"
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
xlsheet.Range("B1").EntireColumn.ColumnWidth = 25
xlsheet.Range("C1").EntireColumn.ColumnWidth = 10
xlsheet.Range("D1").EntireColumn.ColumnWidth = 10
xlsheet.Range("E1").EntireColumn.ColumnWidth = 10

xlsheet.Range("A1:E1").Font.Bold = True
xlsheet.Range("A1:E1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

Set i = Nothing
Set j = Nothing
For i = 0 To MSGrid.Rows - 1
    For j = 0 To MSGrid.Cols - 1
        xlsheet.Cells(i + 1, j + 1).Value = MSGrid.TextMatrix(i, j)
        
'        If Trim(xlsheet.Cells(i + 1, j + 1).Value) = "Total" Then
'            'xlsheet.Range("A" & i + 1 & ":D" & j + 1).Merge
'            'xlsheet.Range("A" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'            xlsheet.Range("A" & i + 1 & ":C" & i + 1).Font.Bold = True
'        End If
        
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

'If Trim(xlsheet.Cells(i, j).Value) = "Total" Then
'    xlsheet.Range("A" & i & ":D" & j).Merge
'    'xlsheet.Range("A" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'    xlsheet.Range("A" & i & ":E" & i).Font.Bold = True
'End If
        
xlapp.Application.Visible = True
End Sub

Private Sub BtnOK_Click()
If rs.State = 1 Then rs.Close
Debug.Print "select distinct itemcode,itemname from tbl_order where iscomplete=true and orderdate between #" & Format(DTP_from.Value, "m/d/yyyy") & "# and #" & Format(DTP_to.Value, "m/d/yyyy") & "# order by itemname"
rs.Open "select distinct itemcode,itemname from tbl_order where iscomplete=true and orderdate between #" & Format(DTP_from.Value, "m/d/yyyy") & "# and #" & Format(DTP_to.Value, "m/d/yyyy") & "# order by itemname", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    MSGrid.Rows = 1
    s = 1
    While Not rs.EOF
        If rs1.State = 1 Then rs1.Close
        On Error Resume Next
        rs1.Open "select itemcode, itemname, price, sum(quantity) as s from tbl_order where itemcode=" & Val(rs.Fields("itemcode")) & " and iscomplete=true and orderdate between #" & Format(DTP_from.Value, "m/d/yyyy") & "# and #" & Format(DTP_to.Value, "m/d/yyyy") & "# group by itemcode, itemname, price", db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            MSGrid.AddItem s & vbTab & rs1.Fields("itemname") & vbTab & Val(rs1.Fields("price")) & vbTab & rs1.Fields("s") & vbTab & Val(rs1.Fields("price")) * Val(rs1.Fields("s"))
        End If
        rs1.Close
        s = s + 1
        rs.MoveNext
    Wend
    
    totamt = 0
    For i = 1 To MSGrid.Rows - 1
        totamt = Val(totamt) + Val(MSGrid.TextMatrix(i, 4))
    Next i
    MSGrid.AddItem "Total" & vbTab & vbTab & vbTab & vbTab & Format(totamt, "0.00")
    
Else
    MsgBox "No Records Found. Check the Date", vbInformation, "Sri Saravana Bhavan"
End If
End Sub

Private Sub BtnReport_Click()
totamt = 0
For i = 1 To MSGrid.Rows - 2
    totamt = Val(totamt) + Val(MSGrid.TextMatrix(i, 4))
Next i
    
'----------Notepad print------------------
Open App.Path & "\iwisereport.txt" For Output As #1
'Print #1, Chr(18); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
Print #1, "          Sri Saravana Bhavan"
Print #1, ""
Print #1, "            Item Wise Report"
Print #1, ""
Print #1, "Report from " & Format(DTP_from.Value, "dd/mm/yyyy") & " to " & Format(DTP_to.Value, "dd/mm/yyyy")
Print #1, "------------------------------------------"      '42 characters"
Print #1, "ICo Item Name              Rate  Qty   Amt"
Print #1, "------------------------------------------"
For i = 1 To MSGrid.Rows - 2
    iico = 3 - Len(MSGrid.TextMatrix(i, 0))
    iiname = 22 - Len(Mid(MSGrid.TextMatrix(i, 1), 1, 22))
    iirate = 4 - Len(MSGrid.TextMatrix(i, 2))
    iqty = 4 - Len(MSGrid.TextMatrix(i, 3))
    iamt = 5 - Len(MSGrid.TextMatrix(i, 4))
    
    Print #1, Space(iico) & MSGrid.TextMatrix(i, 0) & Space(1) & Mid(MSGrid.TextMatrix(i, 1), 1, 22) & Space(iiname) & Space(1) & Space(iirate) & MSGrid.TextMatrix(i, 2) & Space(1) & Space(iqty) & MSGrid.TextMatrix(i, 3) & Space(1) & Space(iamt) & MSGrid.TextMatrix(i, 4)
Next i
Print #1, "------------------------------------------"
Print #1, "                        Total:" & Space(12 - Len(Format(Val(totamt), "0.00"))) & Format(Val(totamt), "0.00")
Print #1, "                              ------------"
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Chr$(&H1B); "m"; Chr$(&HA);
Close #1
retval = Shell("notepad.exe iwisereport.txt", vbHide)

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
Open App.Path & "\iwisereport.txt" For Input As #1
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

DTP_from.Value = Date
DTP_to.Value = Date

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i
End Sub
