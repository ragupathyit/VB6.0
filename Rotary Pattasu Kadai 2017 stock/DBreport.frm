VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form daybookRpt 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "DayBook Report"
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
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "E&XIT"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
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
      Left            =   3480
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton CmdReport 
      BackColor       =   &H00C0E0FF&
      Caption         =   "RE&PORT"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
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
      Left            =   5520
      TabIndex        =   4
      Top             =   2400
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cr Item"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "DAYBOOK REPORT"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   200
      Width           =   3300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      TabIndex        =   2
      Top             =   2520
      Width           =   1185
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
Attribute VB_Name = "daybookRpt"
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

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub CmdReport_Click()
If Combo1.Text = "" Then
    MsgBox "Select the item properly...", vbInformation, "Rotary Club of Mettupalayam"
    Combo1.SetFocus
Else
    '----------Notepad print------------------
    Open App.Path & "\DBreport1.txt" For Output As #1
    Print #1, Space(16) & "Rotary Club of Mettupalayam"
    Print #1, ""
    Print #1, Space(1) & "Date  =" & Format(Date, "DD/MM/YYYY")
    Print #1, ""
    Print #1, Space(1) & "Name of the Particulars =" & Combo1.Text
    
    If rs.State = 1 Then rs.Close
    If rs1.State = 1 Then rs1.Close
    rs.Open "select * from tbl_DBcredit where particulars='" & Combo1.Text & "' and cdate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
    rs1.Open "select * from tbl_DBdebit where particulars='" & Combo1.Text & "' and ddate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
    
    Print #1, Space(1) & "================================================="
    Print #1, Space(1) & "Date" & Space(28) & "Cr Amt     Dr Amt"
    Print #1, Space(1) & "================================================="
    
    totamt = 0
    While Not rs.EOF
        icdate = 26 - Len(rs.Fields("cdate"))
        iamt = 10 - Len(Format(rs.Fields("amt"), "0.00"))
        Print #1, Space(1) & Format(rs.Fields("cdate"), "DD/MM/YYYY") & Space(icdate) & Space(iamt) & Format(rs.Fields("amt"), "0.00")
        totamt = Val(totamt) + Val(rs.Fields("amt"))
        i = i + 1
        rs.MoveNext
    Wend
    
    totamt1 = 0
    While Not rs1.EOF
        iddate = 38 - Len(rs1.Fields("ddate"))
        iamt1 = 10 - Len(Format(rs1.Fields("amt"), "0.00"))
        Print #1, Space(1) & Format(rs1.Fields("ddate"), "DD/MM/YYYY") & Space(iddate) & Space(iamt1) & Format(rs1.Fields("amt"), "0.00")
        totamt1 = Val(totamt1) + Val(rs1.Fields("amt"))
        i = i + 1
        rs1.MoveNext
    Wend
    
    Print #1, Space(1) & "-------------------------------------------------"
    itotal = 29 - Len("Total:")
    If totamt1 = 0 Then
        Print #1, Space(itotal) & "Total:" & Space(iamt) & Format(totamt, "0.00")
    Else
        Print #1, Space(itotal) & "Total:" & Space(iamt) & Format(totamt, "0.00") & Space(iamt1) & Format(totamt1, "0.00")
    End If
    Print #1, Space(29) & "---------------------"
    
    Print #1, ""
    Close #1
    retval = Shell("notepad.exe DBreport1.txt", vbMaximizedFocus)
    '------------------------------------------------------------------------------------------
    s = MsgBox("Do You Want Print", vbYesNo)
    If s = vbYes Then
        'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
        'Print #1, "TYPE DBreport1.txt>PRN"
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
        Open App.Path & "\DBreport1.txt" For Input As #1
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
    
End If
End Sub

Private Sub Form_Load()

'Me.BackColor = RGB(255, 204, 203)

Call connect
If rs.State = 1 Then rs.Close
rs.Open "select distinct(particulars) from tbl_DBcredit", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    Combo1.AddItem rs.Fields("particulars")
    rs.MoveNext
Wend

DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
