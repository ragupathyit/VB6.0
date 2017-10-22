VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SalesFrm3 
   Caption         =   "Item Sales"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11400
      TabIndex        =   22
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton cmdnext 
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
      Left            =   8640
      TabIndex        =   20
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdprevious 
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
      Left            =   1080
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton cmdcontinue 
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox txtcustname 
      Appearance      =   0  'Flat
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
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtbillno 
      Appearance      =   0  'Flat
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
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdclear 
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton CmdBill 
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton CmdSavePrint 
      Caption         =   "&Print"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9480
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   92536833
      CurrentDate     =   40537
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   16761024
      AllowUserResizing=   2
      FormatString    =   "S.No |Item Code | Item Name                                              |Item Rate  |Quantity   | Amount       "
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
      Height          =   9855
      Left            =   10680
      TabIndex        =   12
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   17383
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16761024
      AllowUserResizing=   2
      FormatString    =   "Code |Item Name                               |Rate      "
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   7920
      TabIndex        =   21
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Last Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label lbllastbill 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
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
      Left            =   1800
      TabIndex        =   17
      Top             =   8640
      Width           =   630
   End
   Begin VB.Label lblbillamt 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   690
      Left            =   6120
      TabIndex        =   16
      Top             =   8520
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALES AND BILL DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Width           =   4005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   11415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   9240
      Width           =   11415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   8280
      TabIndex        =   5
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   4965
      TabIndex        =   4
      Top             =   8760
      Width           =   900
   End
End
Attribute VB_Name = "SalesFrm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub CmdBill_Click()

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\bill.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12
    Print #1, Space(13) & "FRESH PARK"
    Print #1, Space(10) & "Palamudir Nilayam"
    Print #1, Space(8) & "CBE.Road, Mettupalayam"
    Print #1, Space(14) & "PH-224977"
    Print #1, ""
    Print #1, "B. No: " & rs.Fields("billno") & Space(20) & Format(Date, "DD/MM/YY") & " (" & Format(Time, "HH:MM") & ")"
    Print #1, "---------------------------------------------"
    Print #1, "Item" & Space(20) & "Rate" & Space(3) & "Qty" & Space(4) & " Amt"
    Print #1, "---------------------------------------------"
    bamt = rs.Fields("billamt")
    ibamt = 9 - Len(Format(bamt, "0.00"))
    i = 1
    While Not rs.EOF
        ii = 2 - Len(i)
        iname = 21 - Len(rs.Fields("itemname"))
        irate = 7 - Len(Format(rs.Fields("itemrate"), "0.00"))
        iqty = 7 - Len(Format(rs.Fields("itemqty"), "0.00"))
        iamt = 9 - Len(Format(rs.Fields("itemamt"), "0.00"))
        Print #1, UCase(rs.Fields("itemname")) & Space(iname) & Space(irate) & Format(rs.Fields("itemrate"), "0.00") & Space(iqty) & Format(rs.Fields("itemqty"), "0.000") & Space(iamt) & Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        rs.MoveNext
    Wend
    Print #1, "---------------------------------------------"
    Print #1, "Items :" & Val(i) - 1 & Space(ii) & Space(19) & "Total : " & Space(ibamt) & Format(bamt, "0.00")
    Print #1, Space(35) & "----------"
    Print #1, ""
    Print #1, Space(10) & "Thank You! Come Again..."
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Close #1
    retval = Shell("notepad.exe bill.txt", vbHide)
    retval = Shell(App.Path & "\print.bat", vbHide)
End If

'If rs.State = 1 Then rs.Close
'rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    '----------------Printer print--------------------
'    Printer.FontName = "Lucida Console"
'    Printer.FontSize = 8
'    Printer.Print Space(13) & "FRESH PARK"
'    Printer.Print Space(10) & "Palamudir Nilayam"
'    Printer.Print Space(8) & "CBE.Road, Mettupalayam"
'    Printer.Print Space(14) & "PH-224977"
'    Printer.FontSize = 7
'    Printer.Print ""
'    Printer.Print Space(1) & "B. No: " & rs.Fields("billno") & Space(18) & Format(Date, "DD/MM/YY") & " (" & Format(Time, "HH:MM") & ")"
'    Printer.Print "---------------------------------------------"
'    Printer.Print ""
'    Printer.Print "Item" & Space(20) & "Rate" & Space(3) & "Qty" & Space(4) & " Amt"
'    Printer.Print ""
'    Printer.Print "---------------------------------------------"
'    Printer.Print ""
'    bamt = rs.Fields("billamt")
'    ibamt = 9 - Len(Format(bamt, "0.00"))
'    i = 1
'    While Not rs.EOF
'        ii = 2 - Len(i)
'        iname = 21 - Len(rs.Fields("itemname"))
'        irate = 7 - Len(Format(rs.Fields("itemrate"), "0.00"))
'        iqty = 7 - Len(Format(rs.Fields("itemqty"), "0.00"))
'        iamt = 9 - Len(Format(rs.Fields("itemamt"), "0.00"))
'        Printer.Print UCase(rs.Fields("itemname")) & Space(iname) & Space(irate) & Format(rs.Fields("itemrate"), "0.00") & Space(iqty) & Format(rs.Fields("itemqty"), "0.000") & Space(iamt) & Format(rs.Fields("itemamt"), "0.00")
'        i = i + 1
'        rs.MoveNext
'    Wend
'    Printer.Print ""
'    Printer.Print "---------------------------------------------"
'    Printer.Print ""
'    Printer.Print "Items :" & Val(i) - 1 & Space(ii) & Space(19) & "Total : " & Space(ibamt) & Format(bamt, "0.00")
'    Printer.Print Space(35) & "----------"
'    Printer.Print ""
'    Printer.Print Space(10) & "Thank You! Come Again..."
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.EndDoc
'    Printer.KillDoc
'End If

Call cmdclear_Click

End Sub

Private Sub CmdBill_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 118 Then 'F7 Key
    txtbillno.SetFocus
    txtbillno.SelStart = 0
    txtbillno.SelLength = Len(txtbillno.Text)    'select the text
End If
End Sub

Private Sub cmdclear_Click()
txtbillno.Text = ""
txtcustname.Text = ""
lblbillamt.Caption = "0"
MSGrid.Rows = 2
MSGrid.TextMatrix(1, 0) = ""
MSGrid.TextMatrix(1, 1) = ""
MSGrid.TextMatrix(1, 2) = ""
MSGrid.TextMatrix(1, 3) = ""
MSGrid.TextMatrix(1, 4) = ""
MSGrid.TextMatrix(1, 5) = ""

Unload Me
Load Me

Call Form_Load

CmdSave.Enabled = True
CmdBill.Enabled = False
cmdclear.Enabled = True
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
If KeyAscii = 115 Then  's Key
    CmdSavePrint.Enabled = False
    
    If CmdSave.Enabled = False Then
        CmdSave.Enabled = True
    End If
    CmdSave.SetFocus
    Call CmdSave_Click
End If
If KeyAscii = 112 Then   'p Key
    CmdSave.Enabled = False
    
    If CmdSavePrint.Enabled = False Then
        CmdSavePrint.Enabled = True
    End If
    CmdSavePrint.SetFocus
    Call CmdSavePrint_Click
End If
End Sub

Private Sub cmdnext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtbillno.Text = ""
    txtcustname.Text = ""
    lblbillamt.Caption = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""
    
    txtbillno.Text = rs.Fields("billno")
    txtcustname.Text = rs.Fields("custname")
    DTPicker1.Value = rs.Fields("salesdate")
    lblbillamt.Caption = Format(rs.Fields("billamt"), "0.00")
            
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = i
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 3) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemqty"), "0.00")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend
            
    cmdcontinue.Enabled = False
    CmdSavePrint.Enabled = False
    CmdSave.Enabled = False
    CmdBill.Enabled = True
    cmdclear.Enabled = True
        
    CmdBill.SetFocus
Else
    Call cmdclear_Click
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub cmdprevious_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtbillno.Text = ""
    txtcustname.Text = ""
    lblbillamt.Caption = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""

    txtbillno.Text = rs.Fields("billno")
    txtcustname.Text = rs.Fields("custname")
    DTPicker1.Value = rs.Fields("salesdate")
    lblbillamt.Caption = Format(rs.Fields("billamt"), "0.00")
            
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = i
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 3) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemqty"), "0.00")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend
            
    cmdcontinue.Enabled = False
    CmdSavePrint.Enabled = False
    CmdSave.Enabled = False
    CmdBill.Enabled = True
    cmdclear.Enabled = True
        
    CmdBill.SetFocus
Else
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub CmdSave_Click()

If txtbillno.Text = "" Then
    MsgBox "Enter the bill no properly...", vbInformation, "Fresh Park"
    txtbillno.SetFocus
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        rs.AddNew
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.Fields("custname") = UCase(txtcustname.Text)
        rs.Fields("salesdate") = DTPicker1.Value
        rs.Fields("itemcode") = MSGrid.TextMatrix(i, 1)
        rs.Fields("itemname") = MSGrid.TextMatrix(i, 2)
        rs.Fields("itemrate") = Val(MSGrid.TextMatrix(i, 3))
        rs.Fields("itemqty") = Val(MSGrid.TextMatrix(i, 4))
        rs.Fields("itemamt") = Val(MSGrid.TextMatrix(i, 5))
        rs.Fields("billamt") = Val(lblbillamt.Caption)
        rs.Update
    Next i
    rs.Close
    
End If
Call cmdclear_Click
End Sub

Private Sub CmdSave_KeyPress(KeyAscii As Integer)
If KeyAscii = 99 Then  'c keyascii
    cmdcontinue.SetFocus
    CmdSave.Enabled = True
    CmdSavePrint.Enabled = True
End If

If KeyAscii = 112 Then   'p Key
    CmdSave.Enabled = False
    
    If CmdSavePrint.Enabled = False Then
        CmdSavePrint.Enabled = True
    End If
    CmdSavePrint.SetFocus
End If
End Sub

Private Sub CmdSavePrint_Click()
If txtbillno.Text = "" Then
    MsgBox "Enter the bill no properly...", vbInformation, "Fresh Park"
    txtbillno.SetFocus
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        rs.AddNew
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.Fields("custname") = UCase(txtcustname.Text)
        rs.Fields("salesdate") = DTPicker1.Value
        rs.Fields("itemcode") = MSGrid.TextMatrix(i, 1)
        rs.Fields("itemname") = MSGrid.TextMatrix(i, 2)
        rs.Fields("itemrate") = Val(MSGrid.TextMatrix(i, 3))
        rs.Fields("itemqty") = Val(MSGrid.TextMatrix(i, 4))
        rs.Fields("itemamt") = Val(MSGrid.TextMatrix(i, 5))
        rs.Fields("billamt") = Val(lblbillamt.Caption)
        rs.Update
    Next i
    rs.Close
    
End If
Call CmdBill_Click
Call cmdclear_Click
End Sub

Private Sub CmdSavePrint_KeyPress(KeyAscii As Integer)
If KeyAscii = 99 Then 'c key
    cmdcontinue.SetFocus
    CmdSave.Enabled = True
    CmdSavePrint.Enabled = True
End If

If KeyAscii = 115 Then  's Key
    CmdSavePrint.Enabled = False
    
    If CmdSave.Enabled = False Then
        CmdSave.Enabled = True
    End If
    CmdSave.SetFocus
End If

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.Col = 1
    MSGrid.Row = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub MSGrid_EnterCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub MSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then 'F1 Key
    If Not MSGrid.Rows = 1 Then
        MSGrid.CellBackColor = vbWhite
        txtsearch.BackColor = RGB(117, 145, 233)
        txtsearch.SetFocus
    End If
End If

If KeyCode = 118 Then 'F7 Key
    txtbillno.SetFocus
    txtbillno.SelStart = 0
    txtbillno.SelLength = Len(txtbillno.Text)    'select the text
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

    If MSGrid.TextMatrix(MSGrid.Row, 2) = "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_itemmaster where itemcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 1)), db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            MSGrid.TextMatrix(MSGrid.Row, 2) = rs.Fields("itemname")
        End If
        rs.Close
    End If
    
    MSGrid.Row = MSGrid.Row
    MSGrid.Col = 3
    MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
    MSGrid.SetFocus
End If

If KeyCode = 123 Then 'F12 Key

    If MSGrid.TextMatrix(MSGrid.Row, 2) = "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_itemmaster where itemcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 1)), db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            MSGrid.TextMatrix(MSGrid.Row, 2) = rs.Fields("itemname")
            MSGrid.TextMatrix(MSGrid.Row, 3) = rs.Fields("rate")
        End If
        rs.Close
    End If
    
    MSGrid.Row = MSGrid.Row
    MSGrid.Col = 4
    MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
    MSGrid.SetFocus
End If

If KeyCode = 117 Then 'F6 Key for Delete the row
    If Not MSGrid.Rows = 1 Then
        lblbillamt.Caption = Format(Val(lblbillamt.Caption) - Val(MSGrid.TextMatrix(MSGrid.Row, 5)), "0.00")
        MSGrid.Row = MSGrid.Row
        MSGrid.Col = 1
        If MSGrid.Row = 1 Then
            MSGrid.TextMatrix(1, 0) = 1
            MSGrid.TextMatrix(1, 1) = ""
            MSGrid.TextMatrix(1, 2) = ""
            MSGrid.TextMatrix(1, 3) = ""
            MSGrid.TextMatrix(1, 4) = ""
            MSGrid.TextMatrix(1, 5) = ""
        Else
            MSGrid.RemoveItem MSGrid.Row
        End If
        MSGrid.CellBackColor = RGB(117, 145, 233)
    End If
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

If KeyCode = 113 Then   'F2 Key for Opening another SalesFrm1
    Me.WindowState = 1
    SalesFrm1.WindowState = 2
End If

If KeyCode = 114 Then   'F3 Key for Opening another SalesFrm2
    Me.WindowState = 1
    SalesFrm2.WindowState = 2
End If

'If KeyCode = 46 Then   'Delete Key for Opening another SalesFrm2
'    Unload Me
'End If

If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If

End Sub

Private Sub MSGrid_KeyPress(KeyAscii As Integer)

If MSGrid.Col = 1 Or MSGrid.Col = 4 Then ' second and fourth grid coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
    Case 46         ' 46 keyascii is for dot symbol
        If MSGrid.Col = 4 Then
            If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "." Then
                MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
            End If
        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 1 Then
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_itemmaster where itemcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 1)), db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                MSGrid.TextMatrix(MSGrid.Row, 2) = rs.Fields("itemname")
                MSGrid.TextMatrix(MSGrid.Row, 3) = Format(rs.Fields("rate"), "0.00")
                
                If Not MSGrid.TextMatrix(MSGrid.Row, 4) = "" Then
                    MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                    MSGrid.Row = MSGrid.Row + 1     'cursor position changed to the newlly created row
                    MSGrid.Col = 1                  'cursor position changed to the second coloumn of that newly created row
                    MSGrid.TextMatrix(MSGrid.Row, 0) = MSGrid.TextMatrix(MSGrid.Row - 1, 0) + 1    'Sl.no
                    
                    If MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = "" Then
                        MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                    End If
                Else
                    MSGrid.Col = MSGrid.Col + 3  ' Grid entry was changed to 4th coloumn
                End If
            Else
                MSGrid.CellBackColor = vbWhite
                
                If cmdcontinue.Enabled = False Then
                    CmdBill.SetFocus
                Else
                    cmdcontinue.SetFocus
                End If
                
            End If
            rs.Close
        End If
        
        If MSGrid.Col = 4 Then
            If Not MSGrid.TextMatrix(MSGrid.Row, 4) = "" Then
                MSGrid.TextMatrix(MSGrid.Row, 5) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 3)) * Val(MSGrid.TextMatrix(MSGrid.Row, 4)), "0.00")
                
                lblbillamt.Caption = 0
                For i = 1 To MSGrid.Rows - 1
                    lblbillamt.Caption = Format(Val(lblbillamt.Caption) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Total bill amount calculation
                Next i
                
                
                MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                MSGrid.Row = MSGrid.Row + 1     'cursor position changed to the newlly created row
                MSGrid.Col = 1                  'cursor position changed to the second coloumn of that newly created row
                MSGrid.TextMatrix(MSGrid.Row, 0) = MSGrid.TextMatrix(MSGrid.Row - 1, 0) + 1    'Sl.no
                
                If MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = "" Then
                    MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                End If
            End If
        End If
    End Select
End If

If MSGrid.Col = 3 Then 'fourth Coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
    Case 46         ' 46 keyascii is for dot symbol
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 13         ' 13 keyascii is for enter key
        If Not MSGrid.TextMatrix(MSGrid.Row, 3) = "" Then
            MSGrid.Row = MSGrid.Row     'cursor position maintains the same row
            MSGrid.Col = 4              'cursor position changed to the fifth coloumn of that same row
            MSGrid.SetFocus
        Else
            MsgBox "Enter the rate properly", vbInformation, "Fresh Park"
        End If
    End Select
End If

End Sub

Private Sub Form_Load()

If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

Call Fill

DTPicker1.Value = Format(Date, "DD/MM/YYYY")

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = Val(rs.Fields("billno")) + 3
Else
    txtbillno.Text = 1
End If
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select billamt from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lbllastbill.Caption = "0.00"
    lbllastbill.Caption = Format(rs.Fields("billamt"), "0.00")
End If
rs.Close

MSGrid.TextMatrix(1, 0) = 1
MSGrid.Row = 1
MSGrid.Col = 1
MSGrid.CellBackColor = RGB(117, 145, 233)

CmdSave.Enabled = True
CmdBill.Enabled = False
cmdclear.Enabled = True
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid1.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("rate")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub MSGrid_LeaveCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = vbWhite
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
    MSGrid1.CellBackColor = vbWhite
    MSGrid.TextMatrix(MSGrid.Rows - 1, 1) = MSGrid1.TextMatrix(MSGrid1.Row, 0)
    MSGrid.TextMatrix(MSGrid.Rows - 1, 2) = MSGrid1.TextMatrix(MSGrid1.Row, 1)
    MSGrid.TextMatrix(MSGrid.Rows - 1, 3) = MSGrid1.TextMatrix(MSGrid1.Row, 2)
    txtsearch.Text = ""
    
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 4
    MSGrid.SetFocus
End If
End Sub

Private Sub MSGrid1_LeaveCell()
MSGrid1.Row = MSGrid1.Row
MSGrid1.Col = MSGrid1.Col
MSGrid1.CellBackColor = vbWhite
End Sub

Private Sub txtbillno_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then         ' 8 keyascii is for Back Space key
    txtbillno.Text = Left(txtbillno.Text, Len(txtbillno.Text) - 1)
End If

If KeyAscii = 13 Then
    If Not txtbillno.Text = "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            txtbillno.Text = rs.Fields("billno")
            txtcustname.Text = rs.Fields("custname")
            DTPicker1.Value = rs.Fields("salesdate")
            lblbillamt.Caption = Format(rs.Fields("billamt"), "0.00")
            
            i = 1
            While Not rs.EOF
                MSGrid.TextMatrix(i, 0) = i
                MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
                MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
                MSGrid.TextMatrix(i, 3) = Format(rs.Fields("itemrate"), "0.00")
                MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemqty"), "0.00")
                MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
                i = i + 1
                MSGrid.Rows = MSGrid.Rows + 1
                rs.MoveNext
            Wend
            
            cmdcontinue.Enabled = False
            CmdSavePrint.Enabled = False
            CmdSave.Enabled = False
            CmdBill.Enabled = True
            cmdclear.Enabled = True
        
            CmdBill.SetFocus
        Else
            MSGrid.Row = MSGrid.Rows - 1
            MSGrid.Col = 1
            MSGrid.SetFocus
        End If
    Else
        MsgBox "Enter the bill no properly...", vbInformation, "Fresh Park"
        txtbillno.SetFocus
    End If
End If
End Sub

Private Sub txtcustname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DTPicker1.SetFocus
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
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("rate")
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

Private Sub cmdcontinue_GotFocus()
cmdcontinue.BackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdcontinue_LostFocus()
cmdcontinue.BackColor = RGB(239, 234, 219)
End Sub

Private Sub CmdSave_GotFocus()
CmdSave.BackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdSave_LostFocus()
CmdSave.BackColor = RGB(239, 234, 219)
End Sub

Private Sub CmdSavePrint_GotFocus()
CmdSavePrint.BackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdSavePrint_LostFocus()
CmdSavePrint.BackColor = RGB(239, 234, 219)
End Sub

Private Sub CmdBill_GotFocus()
CmdBill.BackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdBill_LostFocus()
CmdBill.BackColor = RGB(239, 234, 219)
End Sub

Private Sub cmdclear_GotFocus()
cmdclear.BackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdclear_LostFocus()
cmdclear.BackColor = RGB(239, 234, 219)
End Sub

Private Sub CmdClose_GotFocus()
CmdClose.BackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdClose_LostFocus()
cmdcontinue.BackColor = RGB(239, 234, 219)
End Sub
