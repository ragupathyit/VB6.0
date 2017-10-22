VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form VoucherPurchaseFrm 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Day Book"
   ClientHeight    =   8775
   ClientLeft      =   -60
   ClientTop       =   -75
   ClientWidth     =   10575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   7920
      Width           =   1215
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
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtdetails 
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
      Height          =   1260
      Left            =   2160
      TabIndex        =   6
      Top             =   6240
      Width           =   7575
   End
   Begin VB.TextBox txtbalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   21
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox txtrbalance 
      Alignment       =   1  'Right Justify
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
      Left            =   4320
      TabIndex        =   1
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txttotbalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   20
      Top             =   2280
      Width           =   3015
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
      Left            =   8040
      TabIndex        =   19
      Top             =   120
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
      Left            =   1560
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Payment Type"
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
      Height          =   1335
      Left            =   720
      TabIndex        =   11
      Top             =   4560
      Width           =   9015
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00C00000&
         Caption         =   "Cash"
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
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton OptBank 
         BackColor       =   &H00C00000&
         Caption         =   "Bank"
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
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton OptCheque 
         BackColor       =   &H00C00000&
         Caption         =   "Cheque"
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
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton OptHawala 
         BackColor       =   &H00C00000&
         Caption         =   "Hawala"
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
         Height          =   495
         Left            =   6720
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
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
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   22
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   741
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
      Format          =   97452035
      CurrentDate     =   42430
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Voucher Date"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   840
      Width           =   2145
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Supplier Name"
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
      Left            =   1320
      TabIndex        =   16
      Top             =   1560
      Width           =   2295
   End
   Begin MSForms.ComboBox cmbsupname 
      Height          =   420
      Left            =   4320
      TabIndex        =   0
      Top             =   1560
      Width           =   4575
      VariousPropertyBits=   746604571
      ForeColor       =   0
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8070;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Total Balance"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   2280
      Width           =   2190
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Received Balance"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   3000
      Width           =   2850
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Balance"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   3720
      Width           =   1290
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Details"
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
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "PURCHASE VOUCHER"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   4350
   End
End
Attribute VB_Name = "VoucherPurchaseFrm"
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
   
Private Sub cmbsupname_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchasebalance where supname='" & Trim(cmbsupname.Text) & "' order by id", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txttotbalance.Text = Format(Val(rs.Fields("balamt")), "0.00")
End If
txtrbalance.SetFocus
End Sub

Private Sub CmdBill_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchasevoucher where id=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\VoucherPurchaseBill.txt" For Output As #1
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, ""
    Print #1, Space(17) & "Sumathi Stores"
    Print #1, Space(13) & "New Vegitable Market"
    Print #1, Space(1) & "Vegitable Supplier and Commission Agency"
    Print #1, Space(13) & "METTUPALAYAM - 641301"
    Print #1, Space(6) & "CELL NO - 93641 33333, 90034 00000"
    Print #1, "--------------------------------------------"      '44 characters
    Print #1, Space(11) & "VOUCHER PURCHASE BILL"
    Print #1, "To:" & Mid(rs.Fields("supname"), 1, 25) & Space(25 - Len(Mid(rs.Fields("supname"), 1, 25)))
    Print #1, "Voucher No: " & rs.Fields("id") & Space(6 - Len(rs.Fields("id"))) & " Date: " & Format(Date, "DD/MM/YY") & " (" & Format(Time, "HH:MM AMPM") & ")"
    Print #1, "--------------------------------------------"      '44 characters
    Print #1, "Particulars " & Space(21) & Space(1) & "    Amount"
    Print #1, "--------------------------------------------"
    rbalance = Round(Format(rs.Fields("rbalance"), "0.00")) & ".00"
    word = ConNumToEngLish(Val(rbalance))
    irbalance = 10 - Len(Format(rbalance, "0.00"))
    
    totobalance = Round(Format(rs.Fields("totobalance"), "0.00")) & ".00"
    itotobalance = 10 - Len(Format(totobalance, "0.00"))
        
    balance = Round(Format(rs.Fields("balance"), "0.00"))
    ibalance = 10 - Len(Format(balance, "0.00"))
        
    itb = 31 - Len("Total Old Balance")
    Print #1, UCase("Total Old Balance") & Space(itb) & Space(2) & Space(itotobalance) & Format(totobalance, "0.00")
    
    irb = 31 - Len("Received Balance")
    Print #1, UCase("Received Balance") & Space(irb) & Space(2) & Space(irbalance) & Format(rbalance, "0.00")
    
    ib = 31 - Len("Remaining Balance")
    Print #1, UCase("Remaining Balance") & Space(ib) & Space(2) & Space(ibalance) & Format(balance, "0.00")
    Print #1, ""
    Print #1, "Payment Type: " & UCase(rs.Fields("paymenttype"))
    If rs.Fields("details") <> "" Then
        Print #1, "Details: "
        Print #1, UCase(rs.Fields("details"))
    End If
    Print #1, "--------------------------------------------"
    Print #1, word & " Rupees Only"
    Print #1, ""
    Print #1, Space(24) & "Authorized Signatory"
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
    End If
    Close #1
    retval = Shell("notepad.exe VoucherPurchaseBill.txt", vbHide)
rs.Close

'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'Print #1, "TYPE VoucherPurchaseBill.txt>PRN"
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
    Open App.Path & "\VoucherPurchaseBill.txt" For Input As #1
    var1 = Input(LOF(1), #1)
    Close #1
    
    sWrittenData = var1 '& vbFormFeed

    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    '<==================== Printing Code ========================>
    
Call cmdclear_Click
    
End Sub

Private Sub cmdclear_Click()
Unload Me
VoucherPurchaseFrm.Show
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnext_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_purchasevoucher where id=" & Val(txtid.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    cmbsupname.Text = ""
    txttotbalance.Text = ""
    txtrbalance.Text = ""
    txtbalance.Text = ""
    txtdetails.Text = ""
    
    txtid.Text = rs1.Fields("id")
    DTPicker1.Value = rs1.Fields("vdate")
    cmbsupname.Text = rs1.Fields("supname")
    txttotbalance.Text = Format(rs1.Fields("totobalance"), "0.00")
    txtrbalance.Text = Format(rs1.Fields("rbalance"), "0.00")
    txtbalance.Text = Format(rs1.Fields("balance"), "0.00")
    If rs1.Fields("paymenttype") = "Cash" Then
        OptCash.Value = True
    ElseIf rs1.Fields("paymenttype") = "Bank" Then
        OptBank.Value = True
    ElseIf rs1.Fields("paymenttype") = "Cheque" Then
        OptCheque.Value = True
    ElseIf rs1.Fields("paymenttype") = "Hawala" Then
        OptHawala.Value = True
    End If
    txtdetails.Text = rs1.Fields("details")
    
    CmdSave.Enabled = False
Else
    Call cmdclear_Click
End If
End Sub

Private Sub cmdprevious_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_purchasevoucher where id=" & Val(txtid.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    cmbsupname.Text = ""
    txttotbalance.Text = ""
    txtrbalance.Text = ""
    txtbalance.Text = ""
    txtdetails.Text = ""
    
    txtid.Text = rs1.Fields("id")
    DTPicker1.Value = rs1.Fields("vdate")
    cmbsupname.Text = rs1.Fields("supname")
    txttotbalance.Text = Format(rs1.Fields("totobalance"), "0.00")
    txtrbalance.Text = Format(rs1.Fields("rbalance"), "0.00")
    txtbalance.Text = Format(rs1.Fields("balance"), "0.00")
    If rs1.Fields("paymenttype") = "Cash" Then
        OptCash.Value = True
    ElseIf rs1.Fields("paymenttype") = "Bank" Then
        OptBank.Value = True
    ElseIf rs1.Fields("paymenttype") = "Cheque" Then
        OptCheque.Value = True
    ElseIf rs1.Fields("paymenttype") = "Hawala" Then
        OptHawala.Value = True
    End If
    txtdetails.Text = rs1.Fields("details")
    
    CmdSave.Enabled = False
End If
End Sub

Private Sub CmdSave_Click()
If cmbsupname.Text = "" Then
    MsgBox "Select the Supplier Name Properly...", vbInformation, "Sumathi Stores"
    cmbsupname.SetFocus
ElseIf txtrbalance.Text = "" Then
    MsgBox "Enter the Received Balance Properly...", vbInformation, "Sumathi Stores"
    txtrbalance.SetFocus
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select sid from tbl_suppliermaster where suppliername='" & cmbsupname.Text & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        s = rs.Fields("sid")
    End If
    rs.Close
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchasevoucher", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields("id") = txtid.Text
    rs.Fields("vdate") = DTPicker1.Value
    rs.Fields("sid") = Val(s)
    rs.Fields("supname") = cmbsupname.Text
    rs.Fields("totobalance") = Format(Val(txttotbalance.Text), "0.00")
    rs.Fields("rbalance") = Format(Val(txtrbalance.Text), "0.00")
    rs.Fields("balance") = Format(Val(txtbalance.Text), "0.00")
    If OptCash.Value = True Then
        rs.Fields("paymenttype") = "Cash"
    ElseIf OptBank.Value = True Then
        rs.Fields("paymenttype") = "Bank"
    ElseIf OptCheque.Value = True Then
        rs.Fields("paymenttype") = "Cheque"
    ElseIf OptHawala.Value = True Then
        rs.Fields("paymenttype") = "Hawala"
    End If
    rs.Fields("details") = txtdetails.Text
    rs.Update
    
    '====================Sales Balance====================
    'If rs.State = 1 Then rs.Close
    'rs.Open "select * from tbl_purchasebalance", db, adOpenDynamic, adLockOptimistic
    'rs.AddNew
    '    rs.Fields("purdate") = DTPicker1.Value
    '    rs.Fields("sid") = Val(s)
    '    rs.Fields("supname") = Trim(cmbsupname.Text)
    '    rs.Fields("balamt") = Format(Val(txtbalance.Text), "0.00")
    '    rs.Fields("obalance") = Format(Val(txttotbalance.Text), "0.00")
    '    rs.Fields("totamt") = Format(Val(txttotbalance.Text), "0.00")
    '    rs.Fields("payamt") = Format(Val(txtrbalance.Text), "0.00")
    '    rs.Fields("baldesc") = txttotbalance.Text & "-" & txtrbalance.Text
    'rs.Update
    'rs.Close
    '====================Sales Balance====================
    
    MsgBox "Voucher Successfully Saved...", vbInformation, "Sumathi Stores"
    
End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(35, 29, 29)
Label2.BackColor = RGB(35, 29, 29)
Label4.BackColor = RGB(35, 29, 29)
Label5.BackColor = RGB(35, 29, 29)
Label6.BackColor = RGB(35, 29, 29)
Label7.BackColor = RGB(35, 29, 29)
Label8.BackColor = RGB(35, 29, 29)
Label10.BackColor = RGB(35, 29, 29)
Frame1.BackColor = RGB(35, 29, 29)
OptCash.BackColor = RGB(35, 29, 29)
OptBank.BackColor = RGB(35, 29, 29)
OptCheque.BackColor = RGB(35, 29, 29)
OptHawala.BackColor = RGB(35, 29, 29)

Call connect
DTPicker1.Value = Date

rs.Open "select suppliername from tbl_suppliermaster order by suppliername", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    cmbsupname.AddItem Trim(rs.Fields("suppliername"))
    rs.MoveNext
Wend
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchasevoucher order by id", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtid.Text = Val(rs.Fields("id")) + 1
Else
    txtid.Text = 1
End If
rs.Close

CmdSave.Enabled = True

End Sub

Private Sub OptBank_Click()
txtdetails.SetFocus
End Sub

Private Sub OptCash_Click()
txtdetails.SetFocus
End Sub

Private Sub OptCheque_Click()
txtdetails.SetFocus
End Sub

Private Sub OptHawala_Click()
txtdetails.SetFocus
End Sub

Private Sub txtrbalance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OptCash.SetFocus
End If
End Sub

Private Sub txtrbalance_LostFocus()
txtrbalance.Text = Format(Val(txtrbalance.Text), "0.00")
txtbalance.Text = Format(Val(txttotbalance.Text) - Val(txtrbalance.Text), "0.00")
End Sub
