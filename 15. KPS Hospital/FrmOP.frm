VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmOP 
   BackColor       =   &H00FF0000&
   Caption         =   "OP Creation"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   16455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   16455
   WindowState     =   2  'Maximized
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
      TabIndex        =   25
      Text            =   "Select the Doctor"
      Top             =   840
      Width           =   3495
   End
   Begin VB.ListBox lstfees 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7725
      Left            =   12240
      Style           =   1  'Checkbox
      TabIndex        =   23
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtpname 
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
      Height          =   480
      Left            =   8880
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17040
      TabIndex        =   9
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
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
      Height          =   480
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   9480
      Width           =   1695
   End
   Begin VB.TextBox txtpayamt 
      Alignment       =   1  'Right Justify
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
      Height          =   480
      Left            =   5280
      TabIndex        =   16
      Text            =   "0"
      Top             =   9480
      Width           =   1695
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   ">> >>"
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
      Left            =   14280
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "<< <<"
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
      Left            =   5400
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "S&ave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17040
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdcontinue 
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17040
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtbillno 
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
      Height          =   480
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17040
      TabIndex        =   10
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17040
      TabIndex        =   11
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton CmdBill 
      Caption         =   "&Bill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17040
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   14040
      TabIndex        =   2
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   97189889
      CurrentDate     =   40537
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   7725
      Left            =   2400
      TabIndex        =   20
      Top             =   1440
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   13626
      _Version        =   393216
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   16711680
      ForeColorFixed  =   16777215
      BackColorSel    =   16711680
      BackColorBkg    =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Particulars                                                                                        |Amount        "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "B. No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2400
      TabIndex        =   26
      Top             =   240
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Dr Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2400
      TabIndex        =   24
      Top             =   960
      Width           =   1290
   End
   Begin VB.Label lblcancel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   17040
      TabIndex        =   22
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label lblbilldate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   17130
      TabIndex        =   21
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   9600
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Payment Amount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   9600
      Width           =   2160
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Bill Amt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   12570
      TabIndex        =   13
      Top             =   9360
      Width           =   2085
   End
   Begin VB.Label lblbill 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   14880
      TabIndex        =   12
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "OP BILL DETAILS"
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
      Left            =   8160
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   13080
      TabIndex        =   3
      Top             =   960
      Width           =   690
   End
End
Attribute VB_Name = "FrmOP"
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

Private Sub CmdBill_Click()
a = MsgBox("Do you want to print the bill", vbYesNo)
If a = vbYes Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_op where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then

        '----------Notepad print------------------
        Open App.Path & "\rptbill.txt" For Output As #1

        Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
        Print #1, ""
        Print #1, Space(4) & "      K.P.S HOSPITALS (P) LTD     "
        Print #1, Space(4) & "ANNUR ROAD, METTUPALAYAM - 641 301"
        Print #1, Space(4) & "        04254-224314, 224315      "
        Print #1, ""
        Print #1, Space(10) & rs.Fields("doctorname")
        Print #1, ""
        Print #1, "Bill No: " & rs.Fields("billno") & Space(10 - Len(rs.Fields("billno"))) & Space(9) & "Date: " & Format(rs.Fields("opdate"), "DD/MM/YY")
        Print #1, "Name: " & Mid(rs.Fields("patientname"), 1, 21) & Space(21 - Len(Mid(rs.Fields("patientname"), 1, 21))) & Space(1) & "Time: " & rs.Fields("optime")
        Print #1, "------------------------------------------"    '42
        Print #1, "Particulars " & Space(19) & Space(1) & "Amount(Rs)"
        Print #1, "------------------------------------------"
        pamt = rs.Fields("payamt")
        ipamt = 10 - Len(Format(pamt, "0.00"))
        'word = ConNumToEngLish(Val(pamt))
        'i = 1
        While Not rs.EOF
            'ii = 2 - Len(i)
            IName = 31 - Len(Mid(rs.Fields("feename"), 1, 29))
            iamt = 10 - Len(Format(rs.Fields("charges"), "0.00"))
            Print #1, UCase(Mid(rs.Fields("feename"), 1, 29)) & Space(IName) & Space(1) & Space(iamt) & Format(rs.Fields("charges"), "0.00")
            i = i + 1
            rs.MoveNext
        Wend
        Print #1, ""
        Print #1, "------------------------------------------"
        Print #1, Space(25) & "Total: " & Space(ipamt) & Format(pamt, "0.00")
        Print #1, Space(32) & "----------"
        'Print #1, word & " Rupees Only"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Close #1
        retval = Shell("notepad.exe rptbill.txt", vbHide)
    End If
    rs.Close

    'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
    'Print #1, "TYPE rptbill.txt>PRN"
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
    Open App.Path & "\rptbill.txt" For Input As #1
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
Call cmdclear_Click
End Sub

Private Sub cmdclear_Click()
Unload Me
Load Me
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdcontinue_Click()
MSGrid.Row = MSGrid.Rows - 1
MSGrid.Col = 1
MSGrid.SetFocus
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdDelete_Click()
db.Execute "update tbl_op set cancel1='Y' where billno=" & Val(txtbillno.Text)
MsgBox "Successfully Bill Canceled...", vbInformation, "KPS Hospital"
Call cmdclear_Click
End Sub

Private Sub cmdcontinue_KeyPress(KeyAscii As Integer)
If KeyAscii = 115 Then  's Key
    If CmdSave.Enabled = False Then
        CmdSave.Enabled = True
    End If
    CmdSave.SetFocus
    'Call CmdSave_Click
End If
End Sub

Private Sub cmdnext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_op where billno=" & Val(txtbillno.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtbillno.Text = ""
    cmbdoctor.Text = ""
    txtpname.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    
    txtbillno.Text = rs.Fields("billno")
    cmbdoctor.Text = rs.Fields("doctorname")
    txtpname.Text = rs.Fields("patientname")
    DTPicker1.Value = rs.Fields("opdate")

    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    lblbill.Caption = Format(rs.Fields("payamt"), "0.00")
    lblbilldate.Caption = rs.Fields("optime")

    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("feename")
        MSGrid.TextMatrix(i, 1) = Format(Val(rs.Fields("charges")), "0.00")
        MSGrid.Rows = MSGrid.Rows + 1
        i = i + 1
        rs.MoveNext
    Wend

    cmdcontinue.Enabled = False
    CmdSave.Enabled = False
    CmdBill.Enabled = True
    cmddelete.Enabled = True
    cmdclear.Enabled = True

    CmdBill.SetFocus
Else
    Call cmdclear_Click
End If

'--------------------Cancel Bill Information---------------------------------------
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_op where cancel1='Y' and billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lblcancel.Caption = "CANCEL"
    cmddelete.Enabled = False
Else
    lblcancel.Caption = ""
    cmddelete.Enabled = True
End If
'--------------------Cancel Bill Information---------------------------------------
End Sub

Private Sub cmdprevious_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_op where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtbillno.Text = ""
    cmbdoctor.Text = ""
    txtpname.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    
    txtbillno.Text = rs.Fields("billno")
    cmbdoctor.Text = rs.Fields("doctorname")
    txtpname.Text = rs.Fields("patientname")
    DTPicker1.Value = rs.Fields("opdate")

    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    lblbill.Caption = Format(rs.Fields("payamt"), "0.00")
    lblbilldate.Caption = rs.Fields("optime")

    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("feename")
        MSGrid.TextMatrix(i, 1) = Format(Val(rs.Fields("charges")), "0.00")
        MSGrid.Rows = MSGrid.Rows + 1
        i = i + 1
        rs.MoveNext
    Wend

    cmdcontinue.Enabled = False
    CmdSave.Enabled = False
    CmdBill.Enabled = True
    cmddelete.Enabled = True
    cmdclear.Enabled = True

    CmdBill.SetFocus
End If

'--------------------Cancel Bill Information---------------------------------------
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_op where cancel1='Y' and billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lblcancel.Caption = "CANCEL"
    cmddelete.Enabled = False
Else
    lblcancel.Caption = ""
    cmddelete.Enabled = True
End If
'--------------------Cancel Bill Information---------------------------------------
End Sub

Private Sub CmdSave_Click()
If txtbillno.Text = "" Then
    MsgBox "Enter the bill no properly...", vbInformation, "KPS Hospital"
    txtbillno.SetFocus
ElseIf cmbdoctor.Text = "Select the Doctor" Then
    MsgBox "Select the doctor name properly...", vbInformation, "KPS Hospital"
    cmbdoctor.SetFocus
ElseIf txtpname.Text = "" Then
    MsgBox "Enter the patient name properly...", vbInformation, "KPS Hospital"
    txtpname.SetFocus
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_op", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 1
        rs.AddNew
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.Fields("doctorname") = Trim(UCase(cmbdoctor.Text))
        rs.Fields("patientname") = Trim(UCase(txtpname.Text))
        rs.Fields("opdate") = DTPicker1.Value
        rs.Fields("optime") = Format(Time, "HH:MM AMPM")
        rs.Fields("feename") = MSGrid.TextMatrix(i, 0)
        rs.Fields("charges") = MSGrid.TextMatrix(i, 1)
        rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
        rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
        rs.Fields("cancel1") = "N"
        rs.Update
    Next i
    rs.Close
    Call CmdBill_Click
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.Row = 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub MSGrid_EnterCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub MSGrid_LeaveCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = vbWhite
End Sub

Private Sub MSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 188 Then '  , Key for Previous Record
    Call cmdprevious_Click
    MSGrid.Row = 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If

If KeyCode = 190 Then '    . Key for Next Record
    Call cmdnext_Click
    MSGrid.Row = 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If

If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If

End Sub

Private Sub MSGrid_KeyPress(KeyAscii As Integer)
If MSGrid.Col = 0 Or MSGrid.Col = 1 Then    'Fees Name, Amount only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
    Case 32         ' 32 keyascii is for space bar key
        If MSGrid.Col = 0 Then  ' For Fees Name Coloumn Only
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 46         ' 46 keyascii is for dot symbol
        If MSGrid.Col = 1 Then 'For Amount Coloumn Only
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        If MSGrid.Col = 1 Then 'For Amount Coloumn Only
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
        If MSGrid.Col = 0 Then  ' For Fees Name Coloumn Only
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 97 To 122  ' 97-122 keyascii is for small a to z
        If MSGrid.Col = 0 Then  ' For Fees Name Coloumn Only
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 0 Then  ' Item Code
            If MSGrid.TextMatrix(MSGrid.Row, 0) <> "" Then
                MSGrid.Row = MSGrid.Row
                MSGrid.Col = 1
                MSGrid.CellBackColor = RGB(117, 145, 233)
                MSGrid.SetFocus
            End If
        Else
            If MSGrid.TextMatrix(MSGrid.Row, 1) <> "" Then
                txttotamt.Text = 0
                txtpayamt.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotamt.Text = Format(Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 1)), "0.00")
                    txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
                    lblbill.Caption = Format(Val(txtpayamt.Text), "0.00")
                Next i
                
                MSGrid.Row = MSGrid.Row
                MSGrid.Col = 0
                MSGrid.CellBackColor = vbWhite
                If CmdSave.Enabled = True Then
                    CmdSave.SetFocus  'cursor navigation to the Continue Button
                Else
                    cmdclear.SetFocus
                End If
            End If
        End If
    End Select
End If
End Sub

Private Sub Form_Load()
Call connect
Call Fill

DTPicker1.Value = Date

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_op order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = Val(rs.Fields("billno")) + 1
Else
    txtbillno.Text = 1
End If
rs.Close

CmdSave.Enabled = True
CmdBill.Enabled = False
cmddelete.Enabled = False
cmdclear.Enabled = True
End Sub

Private Function Fill()
stmt = "select * from tbl_feesmaster order by fcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        lstfees.AddItem rs.Fields("feename")
        rs.MoveNext
    Loop
End If

If rs.State = 1 Then rs.Close
rs.Open "select doctorname from tbl_doctormaster order by dcode", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        cmbdoctor.AddItem rs.Fields("doctorname")
        rs.MoveNext
    Loop
End If
End Function

Private Sub lstfees_ItemCheck(Item As Integer)
If lstfees.Selected(Item) = True Then
    If MSGrid.TextMatrix(1, 0) = "" Then
        MSGrid.Rows = 1
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_feesmaster where feename='" & Trim(lstfees.List(Item)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        If Trim(lstfees.List(Item)) = "OTHERS" Then
            sa = InputBox("Enter the Fees Name", "Fees Name")
            MSGrid.AddItem UCase(sa) & vbTab & ""
        Else
            MSGrid.AddItem rs.Fields("feename") & vbTab & Format(rs.Fields("charges"), "0.00")
            txttotamt.Text = Format(Val(txttotamt.Text) + Val(rs.Fields("charges")), "0.00")
            txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
            lblbill.Caption = Format(Val(txtpayamt.Text), "0.00")
        End If
        
        MSGrid.Row = MSGrid.Rows - 1
        MSGrid.Col = 1
        MSGrid.CellBackColor = RGB(117, 145, 233)
        MSGrid.SetFocus
    End If
    rs.Close
Else
    For i = 1 To MSGrid.Rows - 1
        If MSGrid.TextMatrix(i, 0) = lstfees.List(Item) Then
            If MSGrid.Rows = 2 Then
                MSGrid.TextMatrix(1, 0) = ""
                MSGrid.TextMatrix(1, 1) = ""
                txttotamt.Text = 0
                txtpayamt.Text = 0
            Else
                txttotamt.Text = Format(Val(txttotamt.Text) - Val(MSGrid.TextMatrix(i, 1)), "0.00")
                txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
                lblbill.Caption = Format(Val(txtpayamt.Text), "0.00")
                MSGrid.RemoveItem i
            End If
            
            MSGrid.Row = MSGrid.Rows - 1
            MSGrid.Col = 1
            MSGrid.CellBackColor = RGB(117, 145, 233)
            MSGrid.SetFocus
            Exit For
        End If
    Next
End If
End Sub

Private Sub txtpname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MSGrid.Row = 1
    MSGrid.Col = 1
    MSGrid.CellBackColor = RGB(117, 145, 233)
    MSGrid.SetFocus
End If
End Sub

Private Sub txtbillno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not txtbillno.Text = "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_op where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            txtbillno.Text = rs.Fields("billno")
            txtpname.Text = rs.Fields("patientname")
            DTPicker1.Value = rs.Fields("opdate")

            txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
            txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
            lblbill.Caption = Format(rs.Fields("payamt"), "0.00")
            lblbilldate.Caption = rs.Fields("optime")

            i = 1
            While Not rs.EOF
                MSGrid.TextMatrix(i, 0) = rs.Fields("feename")
                MSGrid.TextMatrix(i, 1) = Format(Val(rs.Fields("charges")), "0.00")
                MSGrid.Rows = MSGrid.Rows + 1
                i = i + 1
                rs.MoveNext
            Wend

            cmdcontinue.Enabled = False
            CmdSave.Enabled = False
            CmdBill.Enabled = True
            cmddelete.Enabled = True
            cmdclear.Enabled = True

            CmdBill.SetFocus
        End If
    Else
        MsgBox "Enter the bill no properly...", vbInformation, "KPS Hospital"
        txtbillno.SetFocus
    End If
End If
End Sub

Private Sub txtpayamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    End If
End If
End Sub

'112 keycode=F1
'to
'123 keycode=F12
