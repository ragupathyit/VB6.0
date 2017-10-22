VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form SalesStatementRpt 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Period Wise Report"
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   15855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbcf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&OK"
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
      Left            =   13200
      TabIndex        =   10
      Top             =   960
      Width           =   1335
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
      Left            =   8040
      TabIndex        =   6
      Top             =   8640
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
      Left            =   6240
      TabIndex        =   1
      Top             =   8640
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   11280
      TabIndex        =   4
      Top             =   960
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   97714179
      CurrentDate     =   40537
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Amount       |Date          |Details                                                                         "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid1 
      Height          =   6375
      Left            =   7920
      TabIndex        =   9
      Top             =   1800
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Amount       |Date          |Details                                                                         "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   8760
      TabIndex        =   15
      Top             =   960
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   97714179
      CurrentDate     =   40537
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   960
      Width           =   3495
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "6165;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Carried Forward"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   8160
      Width           =   3900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
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
      Left            =   10560
      TabIndex        =   12
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      TabIndex        =   7
      Top             =   960
      Width           =   2520
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
      Left            =   10800
      TabIndex        =   5
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES STATEMENT REPORT"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   5625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Report From"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "SalesStatementRpt"
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
   
Private Sub cmdok_Click()
MSGrid.Rows = 1
MSGrid1.Rows = 1
'-----------------------------------Debit Side All Sales Bills------------------------------------------
If rs.State = 1 Then rs.Close
If Combo1.Text <> "" Then
    rs.Open "select distinct billno, custname, salesdate, cooly, lorryhire, gridtotqty, gridtotamt from tbl_sales where custname='" & Trim(Combo1.Text) & "' and salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true order by billno", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "Select distinct billno, salesdate, cooly, lorryhire, gridtotqty, gridtotamt from tbl_sales where salesdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# and isdel=true order by billno", db, adOpenDynamic, adLockOptimistic
End If
If Not rs.EOF Then
    While Not rs.EOF
        MSGrid.AddItem Format(Val(rs.Fields("cooly")) + Val(rs.Fields("lorryhire")) + Val(rs.Fields("gridtotamt")), "0.00") & vbTab & Format(rs.Fields("salesdate"), "DD/MM/YYYY") & vbTab & "B. No:" & rs.Fields("billno") & " T. Qty:" & Format(rs.Fields("gridtotqty"), "0.00")
        rs.MoveNext
    Wend
End If
rs.Close

'-----------------------------------Credit Side All Sales Bills------------------------------------------
If rs.State = 1 Then rs.Close
If Combo1.Text <> "" Then
    rs.Open "select id, custname, vdate, rbalance, paymenttype, details from tbl_salesvoucher where custname='" & Trim(Combo1.Text) & "' and vdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# order by id, vdate", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "Select id, vdate, rbalance, paymenttype, details from tbl_salesvoucher where vdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "# order by id, vdate", db, adOpenDynamic, adLockOptimistic
End If
If Not rs.EOF Then
    While Not rs.EOF
        MSGrid1.AddItem Format(Val(rs.Fields("rbalance")), "0.00") & vbTab & Format(rs.Fields("vdate"), "DD/MM/YYYY") & vbTab & "Pay Type:" & rs.Fields("paymenttype") & " Details:" & rs.Fields("details")
        rs.MoveNext
    Wend
End If
rs.Close

'------------------------Row adjust-------------------------------
If MSGrid.Rows - 1 > MSGrid1.Rows - 1 Then
    ims = Val(MSGrid.Rows - 1) - Val(MSGrid1.Rows - 1)
    For im = 1 To ims
        MSGrid1.AddItem ""
    Next im
ElseIf MSGrid.Rows - 1 < MSGrid1.Rows - 1 Then
    ims = Val(MSGrid1.Rows - 1) - Val(MSGrid.Rows - 1)
    For im = 1 To ims
        MSGrid.AddItem ""
    Next im
End If
MSGrid.AddItem "-"
MSGrid1.AddItem "-"

'-----------------------Total Debit Side--------------------------
tamt = 0
For i = 1 To MSGrid.Rows - 1
    If MSGrid.TextMatrix(i, 0) <> "" Then
        tamt = Val(tamt) + Val(MSGrid.TextMatrix(i, 0))
    End If
Next i
MSGrid.AddItem Format(tamt, "0.00") & vbTab & vbTab & "Total Amount"

'-----------------------Total Debit Side--------------------------
tsvamt = 0
For i = 1 To MSGrid1.Rows - 1
    If MSGrid1.TextMatrix(i, 0) <> "" Then
        tsvamt = Val(tsvamt) + Val(MSGrid1.TextMatrix(i, 0))
    End If
Next i
MSGrid1.AddItem Format(tsvamt, "0.00") & vbTab & vbTab & "Total Amount"

'----------------------------------Total Balance Carried Forward-----------------------------------------
txtbcf.Text = Format(Val(MSGrid.TextMatrix(MSGrid.Rows - 1, 0)) - Val(MSGrid1.TextMatrix(MSGrid1.Rows - 1, 0)), "0.00")

End Sub

Private Sub CmdReport_Click()
'----------Notepad print------------------
Open App.Path & "\rpt_sales_statement.txt" For Output As #1

Print #1, Chr(18); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
Print #1, "Report Created Date: " & Format(Date, "DD/MM/YYYY")
If Combo1.Text <> "" Then
    Print #1, Combo1.Text
End If
Print #1, "Sales Statement as on " & Format(DTPicker1.Value, "DD/MM/YYYY") & " to " & Format(DTPicker2.Value, "DD/MM/YYYY")
Print #1, "                       Credit                                                       Debit                               "
Print #1, "-----------------------------------------------------------||-----------------------------------------------------------"           ' 120
Print #1, "  Amount  |   Date   |   Particulars                       ||  Amount  |   Date   |   Particulars                       "
Print #1, "-----------------------------------------------------------||-----------------------------------------------------------"           ' 120
X = 7
For i = 1 To MSGrid.Rows - 1
    iamt = 10 - Len(Format(MSGrid.TextMatrix(i, 0), "0.00"))
    idate = 10 - Len(Format(Trim(MSGrid.TextMatrix(i, 1)), "DD/MM/YYYY"))
    ipart = 37 - Len(MSGrid.TextMatrix(i, 2))
    iamt1 = 10 - Len(Format(MSGrid1.TextMatrix(i, 0), "0.00"))
    idate1 = 10 - Len(Format(Trim(MSGrid1.TextMatrix(i, 1)), "DD/MM/YYYY"))
    ipart1 = 37 - Len(MSGrid1.TextMatrix(i, 2))
    
    If MSGrid.TextMatrix(i, 0) = "-" Then
        Print #1, "-----------------------------------------------------------||-----------------------------------------------------------"           ' 120
    Else
        Print #1, Space(iamt) & Format(MSGrid.TextMatrix(i, 0), "0.00") & "|" & Format(MSGrid.TextMatrix(i, 1), "DD/MM/YYYY") & Space(idate) & "|" & MSGrid.TextMatrix(i, 2) & Space(ipart) & "||" & Space(iamt1) & Format(MSGrid1.TextMatrix(i, 0), "0.00") & "|" & Format(MSGrid1.TextMatrix(i, 1), "DD/MM/YYYY") & Space(idate1) & "|" & MSGrid1.TextMatrix(i, 2) & Space(ipart1)
    End If
    
    X = X + 1
    If X = 64 Then
        X = 6
        Print #1, Chr(12);          ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
        Print #1, ""
        Print #1, "Sales Statement as on " & Format(DTPicker1.Value, "DD/MM/YYYY") & " to " & Format(DTPicker2.Value, "DD/MM/YYYY")
        Print #1, "                       Credit                                                       Debit                               "
        Print #1, "-----------------------------------------------------------||-----------------------------------------------------------"           ' 120
        Print #1, "  Amount  |   Date   |   Particulars                       ||  Amount  |   Date   |   Particulars                       "
        Print #1, "-----------------------------------------------------------||-----------------------------------------------------------"           ' 120
    End If
Next i
Print #1, "-----------------------------------------------------------||-----------------------------------------------------------"           ' 120
Print #1, Space(10 - Len(Format(Val(txtbcf.Text), "0.00"))) & Format(Val(txtbcf.Text), "0.00") & "  " & Label7.Caption
Close #1
retval = Shell("notepad.exe rpt_sales_statement.txt", vbMaximizedFocus)

'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'Print #1, "TYPE rpt_sales_statement.txt>PRN"
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
    Open App.Path & "\rpt_sales_statement.txt" For Input As #1
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
