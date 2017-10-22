VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesFrm1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Item Retail Sales"
   ClientHeight    =   9705
   ClientLeft      =   -60
   ClientTop       =   -75
   ClientWidth     =   17865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   17865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcustname 
      Appearance      =   0  'Flat
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "sa"
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Delete"
      Enabled         =   0   'False
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
      Left            =   6600
      TabIndex        =   11
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox txtgridtotamt 
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
      Height          =   420
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtdiscount 
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
      Height          =   420
      Left            =   360
      TabIndex        =   19
      Text            =   "0"
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox txtpayamt 
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
      Height          =   420
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
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
      Left            =   12720
      TabIndex        =   17
      Top             =   0
      Width           =   5175
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C0E0FF&
      Caption         =   ">> >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdprevious 
      BackColor       =   &H00C0E0FF&
      Caption         =   "<< <<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "S&ave"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdcontinue 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Continue"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox txtbillno 
      Appearance      =   0  'Flat
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C&lear"
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
      Left            =   8400
      TabIndex        =   12
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Exit"
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
      Left            =   10200
      TabIndex        =   13
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton CmdBill 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Bill"
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
      Left            =   4800
      TabIndex        =   10
      Top             =   9120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   10200
      TabIndex        =   2
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
      Format          =   108593153
      CurrentDate     =   43019
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16761024
      BackColorBkg    =   16744576
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   $"SalesFrm1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid1 
      Height          =   9495
      Left            =   12720
      TabIndex        =   26
      Top             =   360
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   16748
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16761024
      BackColorBkg    =   16744576
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "I Code |Item Name                                |Item Type   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   9000
      Width           =   12735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "SALES AND BILLING DETAILS"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   5115
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   10200
      TabIndex        =   25
      Top             =   8160
      Width           =   465
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amt"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3960
      TabIndex        =   23
      Top             =   8520
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Discount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      TabIndex        =   22
      Top             =   8520
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Amt"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1800
      TabIndex        =   21
      Top             =   8520
      Width           =   1245
   End
   Begin VB.Label lbllastbill 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   8160
      TabIndex        =   14
      Top             =   8040
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   9600
      TabIndex        =   4
      Top             =   1080
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "SalesFrm1"
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
If rs.State = 1 Then rs.Close
SQL = "SELECT billno, salesdate, itemcode, itemname, itemrate, sum(quantity) as qty, sum(itemamt) as amt, discount, totamt From tbl_sales GROUP BY billno, salesdate, itemcode, itemname, itemrate, discount, totamt HAVING billno=" & Val(txtbillno.Text)
Debug.Print SQL
rs.Open SQL, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\bill.txt" For Output As #1
        Print #1, Chr(27); Chr(77);         ' Printer 12 Pitch=Chr(27)    Form feed=Chr(12); 10 pitch=Chr(18);
        Print #1, Space(20) & "BILL"
        Print #1, Space(9) & "Rotary Mettupalayam"
        Print #1, Space(10) & "Padma Sri Agencies"
        Print #1, Space(7) & "GSTN No: 33AGIPL4304J1ZL"
        Print #1, ""
        'Print #1, "To: " & Mid(rs.Fields("custname"), 1, 29) & Space(31 - Len(Mid(rs.Fields("custname"), 1, 29))) & "Bill No: " & rs.Fields("billno")
        billno = rs.Fields("billno")
        Print #1, "Bill No: " & rs.Fields("billno") & Space(2) & "      " & Format(Date, "DD/MM/YY") & "(" & Format(Time, "HH:MM AMPM") & ")"
        Print #1, "------------------------------------------"      '42 characters
        Print #1, "Item Name " & Space(11) & Space(1) & "I.Rate" & Space(1) & " Qty" & Space(1) & "  Amount"
        Print #1, "------------------------------------------"
        tamt = Round(Format(rs.Fields("totamt"), "0.00")) & ".00"
        word = ConNumToEngLish(Val(tamt))
        itamt = 10 - Len(Format(tamt, "0.00"))

'        tsround = Round(Val(rs.Fields("totamt"))) - Val(rs.Fields("totamt"))
'        itsround = 10 - Len(Format(tsround, "0.00"))

        tsdis = rs.Fields("discount")
        idis = 10 - Len(Format(tsdis, "0.00"))

        i = 1
        While Not rs.EOF
            ii = 2 - Len(i)
            iname = 21 - Len(Mid(rs.Fields("itemname"), 1, 21))
            irate = 6 - Len(rs.Fields("itemrate"))
            iqty = 4 - Len(rs.Fields("qty"))
            iamt = 8 - Len(Format(rs.Fields("amt"), "0.00"))

            Print #1, UCase(Mid(rs.Fields("itemname"), 1, 21)) & Space(iname) & Space(1) & Space(irate) & rs.Fields("itemrate") & Space(1) & Space(iqty) & rs.Fields("qty") & Space(1) & Space(iamt) & Format(rs.Fields("amt"), "0.00")
            i = i + 1
            rs.MoveNext
        Wend

        If Val(tsdis) <> 0 Then
            Print #1, ""
            Print #1, "Discount              " & Space(idis) & Format(tsdis, "0.00")
        End If

'        If Val(tsround) <> 0 Then
'            Print #1, Space(34) & "Round: " & Space(itsround) & Format(tsround, "0.00")
'        End If

        Print #1, "------------------------------------------"
        Print #1, "Items: " & Val(i) - 1 & Space(ii) & Space(16) & "Total: " & Space(itamt) & Round(Format(tamt, "0.00")) & ".00"
        Print #1, Space(32) & "----------"
        'Print #1, word & " Rupees Only"
        Print #1, "           ***Including GST***            "
        'Print #1, Space(43) & "Authorized Signatory"
        Print #1, ""
        Print #1, Space(2) & "The proceeds of this bill will be used "
        Print #1, Space(2) & "         for public cause"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Chr$(&H1B); "m"; Chr$(&HA);   'Cutter Code
        'Print #1, ""
        'Print #1, ""
    Close #1
    retval = Shell("notepad.exe bill.txt", vbHide)
End If
rs.Close

'----------Notepad print------------------
Open App.Path & "\token.txt" For Output As #1
    Print #1, "Bill No: " & billno
    Print #1, ""
    Print #1, "       Total: " & Space(itamt) & Round(Format(tamt, "0.00")) & ".00"
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Chr$(&H1B); "m"; Chr$(&HA);   'Cutter Code
Close #1
retval = Shell("notepad.exe token.txt", vbHide)
'----------Notepad print------------------


'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'Print #1, "TYPE bill.txt>PRN"
'Print #1, "EXIT"
'Close #1
'retval = Shell(App.Path & "\print.bat", vbHide)

'<==================== Printing Code ========================>
Dim lhPrinter As Long
Dim lReturn As Long
Dim lpcWritten As Long
Dim lDoc As Long
Dim sWrittenData As String
Dim sWrittenData1 As String
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
Open App.Path & "\bill.txt" For Input As #1
var1 = Input(LOF(1), #1)
Do While Not EOF(1)
    Line Input #1, sWrittenData
Loop
Close #1
sWrittenData = var1 '& vbFormFeed

Dim var2 As String
Open App.Path & "\token.txt" For Input As #1
var2 = Input(LOF(1), #1)
Do While Not EOF(1)
    Line Input #1, sWrittenData1
Loop
Close #1
sWrittenData1 = var2 '& vbFormFeed

lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
Len(sWrittenData), lpcWritten)

lReturn = WritePrinter(lhPrinter, ByVal sWrittenData1, _
Len(sWrittenData1), lpcWritten)

lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)
'<==================== Printing Code ========================>

Call cmdclear_Click

End Sub

Private Sub cmdclear_Click()
Unload Me
SalesFrm1.Show
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

Private Sub cmdcontinue_KeyPress(KeyAscii As Integer)
If KeyAscii = 97 Then  'a Key
     If CmdSave.Enabled = False Then
        CmdSave.Enabled = True
    End If
    CmdSave.SetFocus
    Call CmdSave_Click
End If
If KeyAscii = 100 Then  'd Key
    txtdiscount.SetFocus
    txtdiscount.SelStart = 0
    txtdiscount.SelLength = Len(txtdiscount.Text)    'select the text
End If
'MsgBox KeyAscii
If KeyAscii = 112 Then  'p Key for save and bill
    Call CmdSave_Click
End If
End Sub

Private Sub CmdDelete_Click()
db.Execute "delete from tbl_sales where billno=" & Val(txtbillno.Text)
'Stock Minus to Plus Update-------------------------------
For i = 1 To MSGrid.Rows - 2
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_stock where itemcode=" & Val(MSGrid.TextMatrix(i, 1)), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.Fields("qty") = Format(Val(rs.Fields("qty")) + Val(MSGrid.TextMatrix(i, 5)), "0.00")
        rs.Update
    End If
    rs.Close
Next i
'Stock Minus to Plus Update-------------------------------

MsgBox "Successfully Deleted...", vbInformation, "Rotary Club of Mettupalayam"
Call cmdclear_Click
End Sub

Private Sub cmdnext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtbillno.Text = ""
    txtcustname.Text = ""
    txtgridtotamt.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""
    MSGrid.TextMatrix(1, 6) = ""

    txtbillno.Text = rs.Fields("billno")
    txtcustname.Text = rs.Fields("custname")
    DTPicker1.Value = rs.Fields("salesdate")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtdiscount.Text = Format(rs.Fields("discount"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    lbllastbill.Caption = Format(Val(txtpayamt.Text), "0.00")
    
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = i
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 3) = rs.Fields("itemtype")
        MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 5) = rs.Fields("quantity")
        MSGrid.TextMatrix(i, 6) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    cmdcontinue.Enabled = False
    CmdSave.Enabled = False
    CmdBill.Enabled = True
    cmddelete.Enabled = True

    CmdBill.SetFocus
Else
    Call cmdclear_Click
End If
End Sub

Private Sub cmdprevious_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtbillno.Text = ""
    txtcustname.Text = ""
    txtgridtotamt.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""
    MSGrid.TextMatrix(1, 6) = ""

    txtbillno.Text = rs.Fields("billno")
    txtcustname.Text = rs.Fields("custname")
    DTPicker1.Value = rs.Fields("salesdate")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtdiscount.Text = Format(rs.Fields("discount"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    lbllastbill.Caption = Format(Val(txtpayamt.Text), "0.00")
    
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = i
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 3) = rs.Fields("itemtype")
        MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 5) = rs.Fields("quantity")
        MSGrid.TextMatrix(i, 6) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    cmdcontinue.Enabled = False
    CmdSave.Enabled = False
    CmdBill.Enabled = True
    cmddelete.Enabled = True

    CmdBill.SetFocus
Else
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub CmdSave_Click()
'If txtcustname.Text = "" Then
'    MsgBox "Enter the customer name properly...", vbInformation, "Rotary Club of Mettupalayam"
'    txtcustname.SetFocus
'Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then     '  If the record is allready stored means we delete it and then update it
        db.Execute "delete from tbl_sales where billno=" & Val(txtbillno.Text)
        'Stock Minus to Plus Update-------------------------------
        For i = 1 To MSGrid.Rows - 2
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_stock where itemcode=" & Val(MSGrid.TextMatrix(i, 1)), db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                rs.Fields("qty") = Format(Val(rs.Fields("qty")) + Val(MSGrid.TextMatrix(i, 5)), "0.00")
                rs.Update
            End If
            rs.Close
        Next i
        'Stock Minus to Plus Update-------------------------------
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        rs.AddNew
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.Fields("custname") = UCase(txtcustname.Text)
        rs.Fields("salesdate") = DTPicker1.Value
        rs.Fields("itemcode") = MSGrid.TextMatrix(i, 1)
        rs.Fields("itemname") = MSGrid.TextMatrix(i, 2)
        rs.Fields("itemtype") = MSGrid.TextMatrix(i, 3)
        rs.Fields("itemrate") = Format(Val(MSGrid.TextMatrix(i, 4)), "0.00")
        rs.Fields("quantity") = Val(MSGrid.TextMatrix(i, 5))
        rs.Fields("itemamt") = Format(Val(MSGrid.TextMatrix(i, 6)), "0.00")
        rs.Fields("gridtotamt") = Format(Val(txtgridtotamt.Text), "0.00")
        rs.Fields("discount") = Format(Val(txtdiscount.Text), "0.00")
        rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
        rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
        rs.Update
    Next i
    rs.Close

    'Stock Minus Update-------------------------------
    For i = 1 To MSGrid.Rows - 2
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_stock where itemcode=" & Val(MSGrid.TextMatrix(i, 1)), db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            rs.Fields("qty") = Format(Val(rs.Fields("qty")) - Val(MSGrid.TextMatrix(i, 5)), "0.00")
            rs.Update
        End If
        rs.Close
    Next i
    'Stock Minus Update------------------------------

    'MsgBox "Successfully Saved...", vbInformation, "Rotary Club of Mettupalayam"
'End If

'prn = MsgBox("Do You Want Bill", vbYesNo, "Rotary Club of Mettupalayam")
'If prn = vbYes Then
    Call CmdBill_Click
'Else
'    Call cmdclear_Click
'End If
End Sub

Private Sub CmdSave_KeyPress(KeyAscii As Integer)
If KeyAscii = 99 Then  'c keyascii
    cmdcontinue.SetFocus
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.Col = 1
    MSGrid.Row = 1
    MSGrid.SetFocus
End If

If KeyCode = 188 Then '  , Key for Previous Record
    Call cmdprevious_Click
End If

If KeyCode = 190 Then '    . Key for Next Record
    Call cmdnext_Click
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
If KeyCode = 112 Then 'F1 Key
    MSGrid.CellBackColor = vbWhite
    txtsearch.BackColor = RGB(117, 145, 233)
    txtsearch.SetFocus
End If

'MsgBox KeyCode
If KeyCode = 80 Then    ' Character P is for save and print
    Call CmdSave_Click
End If

'If KeyCode = 117 Then 'F6 Key for Delete the row
If KeyCode = 46 Then 'Delete Key for Delete the row
    txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 6)), "0.00")
    txttotamt.Text = Format(Val(txtgridtotamt.Text), "0.00")
    txtpayamt.Text = Format(Val(txtgridtotamt.Text), "0.00")
    lbllastbill.Caption = Format(Val(txtpayamt.Text), "0.00")
    
    MSGrid.Row = MSGrid.Row
    MSGrid.Col = 1
    If MSGrid.Row = 1 Then
        MSGrid.TextMatrix(1, 0) = ""
        MSGrid.TextMatrix(1, 1) = ""
        MSGrid.TextMatrix(1, 2) = ""
        MSGrid.TextMatrix(1, 3) = ""
        MSGrid.TextMatrix(1, 4) = ""
        MSGrid.TextMatrix(1, 5) = ""
        MSGrid.TextMatrix(1, 6) = ""
    Else
        MSGrid.RemoveItem MSGrid.Row
    End If
    MSGrid.CellBackColor = RGB(117, 145, 233)
End If

If KeyCode = 119 Then 'F8 Key
    txtcustname.SetFocus
    txtcustname.SelStart = 0
    txtcustname.SelLength = Len(txtcustname.Text)
End If

If KeyCode = 120 Then 'F9 Key
    DTPicker1.SetFocus
End If

If KeyCode = 122 Then 'F11 Key
    MSGrid.Row = MSGrid.Row
    MSGrid.Col = 4      ' item rate
    MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
    MSGrid.SetFocus
End If

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

If MSGrid.Col = 1 Or MSGrid.Col = 4 Or MSGrid.Col = 5 Then ' itemcode, itemrate and quantity grid coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
    Case 46         ' 46 keyascii is for dot symbol
        If MSGrid.Col = 4 Then
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
'        If MSGrid.Col = 0 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
    Case 68         ' d is for discount
        txtdiscount.SetFocus
        txtdiscount.SelStart = 0
        txtdiscount.SelLength = Len(txtdiscount.Text)
'    Case 97 To 122  ' 97-122 keyascii is for small a to z
'        If MSGrid.Col = 0 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
    Case 100        ' d is for discount
        txtdiscount.SetFocus
        txtdiscount.SelStart = 0
        txtdiscount.SelLength = Len(txtdiscount.Text)
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 1 Then      ' item code
            If MSGrid.TextMatrix(MSGrid.Row, 1) <> "" Then
                If rs.State = 1 Then rs.Close
                rs.Open "select * from tbl_itemmaster where itemcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 1)), db, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    MSGrid.TextMatrix(MSGrid.Row, 1) = rs.Fields("itemcode")
                    MSGrid.TextMatrix(MSGrid.Row, 2) = rs.Fields("itemname")
                    MSGrid.TextMatrix(MSGrid.Row, 3) = rs.Fields("itemtype")
                    MSGrid.TextMatrix(MSGrid.Row, 4) = Format(rs.Fields("salesrate"), "0.00")
                    MSGrid.Col = MSGrid.Col + 4  ' Grid entry was changed to qty coloumn
                Else
                    MsgBox "There is no stock", vbInformation, "Rotary Club of Mettupalayam"
                    MSGrid.TextMatrix(MSGrid.Row, 1) = ""
                    MSGrid.Row = MSGrid.Row
                    MSGrid.Col = 1
                End If
                rs.Close
            Else
                MSGrid.CellBackColor = vbWhite
                If cmdcontinue.Enabled = False Then
                    CmdBill.SetFocus
                Else
                    cmdcontinue.SetFocus
                End If
            End If
        End If

        If MSGrid.Col = 4 Then  '   itemrate
            If MSGrid.TextMatrix(MSGrid.Row, 4) <> "" Then
                MSGrid.Row = MSGrid.Row
                MSGrid.Col = 5
                MSGrid.SetFocus
            End If
        End If

        If MSGrid.Col = 5 Then  '   Qty
            If MSGrid.TextMatrix(MSGrid.Row, 5) = "" And flag = 1 Then
                MSGrid.TextMatrix(MSGrid.Row, 5) = 1
                flag = 0
            Else
                flag = 1
            End If
            
            If MSGrid.TextMatrix(MSGrid.Row, 5) <> "" Then
                flag = 0
                'Stock Checking Whether the item is in stock------------------------------------------>
                If rs1.State = 1 Then rs1.Close
                rs1.Open "select * from tbl_stock where itemcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 1)), db, adOpenDynamic, adLockOptimistic
                If Not rs1.EOF Then
                    If Val(rs1.Fields("qty")) < Val(MSGrid.TextMatrix(MSGrid.Row, 5)) Then
                        MsgBox MSGrid.TextMatrix(MSGrid.Row, 2) & " is in stock but " & rs1.Fields("qty") & " quantities only.", vbInformation, "Rotary Club of Mettupalayam"
                        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
                        MSGrid.Col = MSGrid.Col
                        MSGrid.Row = MSGrid.Row
                        MSGrid.SetFocus
                    Else    ' Amount Coloumn Calculation
                        MSGrid.TextMatrix(MSGrid.Row, 6) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 4)) * Val(MSGrid.TextMatrix(MSGrid.Row, 5)), "0.00")
                    End If
                End If
                rs1.Close
                'Stock Checking Whether the item is in stock------------------------------------------>
                
                txtgridtotamt.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 6)), "0.00")   'Total grid amount calculation
                    txttotamt.Text = Format(Val(txtgridtotamt.Text), "0.00")
                    txtpayamt.Text = Format(Val(txtgridtotamt.Text), "0.00")
                    lbllastbill.Caption = Format(Val(txtpayamt.Text), "0.00")
                Next i

                If MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = "" Then
                    MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                End If
                
                MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                MSGrid.Row = MSGrid.Row + 1     'cursor position changed to the newlly created row
                MSGrid.Col = 1                  'cursor position changed to the second coloumn of that newly created row
                MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = Val(MSGrid.TextMatrix(MSGrid.Rows - 2, 0)) + 1  'This is for S.No
                MSGrid.SetFocus
                
                'SendKeys vbKeyDown '-------------------This will not work in vista, 7, 8 and later os
                Set WshShell = CreateObject("WScript.Shell")
                WshShell.SendKeys "{DOWN}"
                
            Else
                MSGrid.Row = MSGrid.Row
                MSGrid.Col = 5
                MSGrid.SetFocus
            End If
        End If
    End Select
End If
End Sub

Private Sub Form_Load()

'Me.BackColor = RGB(255, 204, 203)
'MSGrid.BackColorBkg = RGB(255, 204, 203)
'MSGrid1.BackColorBkg = RGB(255, 204, 203)

Call connect
Call Fill

DTPicker1.Value = Date

If rs.State = 1 Then rs.Close
rs.Open "select billno from tbl_sales order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = Val(rs.Fields("billno")) + 1
Else
    txtbillno.Text = 30001  'Bill No. Starts From----------------------------------------
End If
rs.Close

'If rs.State = 1 Then rs.Close
'rs.Open "select totamt from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
    lbllastbill.Caption = "0.00"
'    lbllastbill.Caption = Format(rs.Fields("totamt"), "0.00")
'End If
'rs.Close

MSGrid.Row = 1
MSGrid.Col = 1
MSGrid.TextMatrix(1, 0) = 1
MSGrid.CellBackColor = RGB(117, 145, 233)
'MSGrid.SetFocus

CmdBill.Enabled = False
cmddelete.Enabled = False
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid1.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("itemtype")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub MSGrid1_Click()

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_itemmaster where itemcode=" & Val(MSGrid1.TextMatrix(MSGrid1.Row, 0)), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    MSGrid.TextMatrix(MSGrid.Row, 1) = rs1.Fields("itemcode")
    MSGrid.TextMatrix(MSGrid.Row, 2) = rs1.Fields("itemname")
    MSGrid.TextMatrix(MSGrid.Row, 3) = rs1.Fields("itemtype")
    MSGrid.TextMatrix(MSGrid.Row, 4) = Format(rs1.Fields("salesrate"), "0.00")
    If MSGrid.Col = 1 Then
        MSGrid.Col = MSGrid.Col + 4  ' Grid entry was changed to qty coloumn
    End If
    MSGrid.SetFocus
End If
rs1.Close
End Sub

Private Sub MSGrid1_EnterCell()
MSGrid1.Row = MSGrid1.Row
MSGrid1.Col = MSGrid1.Col
MSGrid1.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub MSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then 'F1 Key
    MSGrid1.CellBackColor = vbWhite
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
    MSGrid.CellBackColor = RGB(117, 145, 233)
End If
End Sub

Private Sub MSGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call MSGrid1_Click
End If
End Sub

Private Sub MSGrid1_LeaveCell()
MSGrid1.Row = MSGrid1.Row
MSGrid1.Col = MSGrid1.Col
MSGrid1.CellBackColor = vbWhite
End Sub

Private Sub txtcustname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 188 Then '  , Key for Previous Record
    Call cmdprevious_Click
End If

If KeyCode = 190 Then '    . Key for Next Record
    Call cmdnext_Click
End If
End Sub

Private Sub txtcustname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MSGrid.Col = 1
    MSGrid.Row = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub txtdiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttotamt.Text = Format(Val(txtgridtotamt.Text) - Val(txtdiscount.Text), "0.00")
    txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
    lbllastbill.Caption = Format(Val(txtpayamt.Text), "0.00")
    CmdSave.SetFocus
End If
End Sub

'112 keycode=F1
'to
'123 keycode=F12
Private Sub txtsearch_Change()
stmt = "select * from tbl_itemmaster where itemname like'" & Trim(txtsearch.Text) & "%' order by itemcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid1.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("itemtype")
        rs.MoveNext
    Loop
End If
rs.Close
End Sub

Private Sub txtsearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then 'F1 Key
    txtsearch.BackColor = vbWhite
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
    MSGrid.CellBackColor = RGB(117, 145, 233)
End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsearch.BackColor = vbWhite
    MSGrid1.Row = 1
    MSGrid1.Col = 1
    MSGrid1.SetFocus
    MSGrid1.CellBackColor = RGB(117, 145, 233)
End If
End Sub
