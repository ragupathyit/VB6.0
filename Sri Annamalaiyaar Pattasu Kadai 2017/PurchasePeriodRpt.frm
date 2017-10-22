VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchasePeriodRpt 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Period Wise Report"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
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
      Left            =   4800
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
      Format          =   100532225
      CurrentDate     =   40537
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4800
      TabIndex        =   4
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
      Format          =   100532225
      CurrentDate     =   40537
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the To Date"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   2925
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE PERIOD WISE REPORT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the From Date"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   3315
   End
End
Attribute VB_Name = "PurchasePeriodRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdReport_Click()
If rs.State = 1 Then rs.Close
rs.Open "Select * from tbl_purchase where purchasedate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "No Records Check the Date", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    Exit Sub
End If

If rs.State = 1 Then rs.Close
rs.Open "Select sum(itemamt) from tbl_purchase where purchasedate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
gamt = rs.Fields(0)

If rs.State = 1 Then rs.Close
rs.Open "Select distinct pid,discount from tbl_purchase where purchasedate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#", db, adOpenDynamic, adLockOptimistic
dis = 0
While Not rs.EOF
    dis = Val(dis) + Val(rs.Fields("discount"))
    rs.MoveNext
Wend

stmt = "select itemcode,itemname,itemtype,itemrate,quantity,itemamt,gridtotamt,discount,totamt from tbl_purchase where purchasedate between #" & DTPicker1.Value & "# and #" & DTPicker2.Value & "#"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\rptpurchaseperiod.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, ""
    Print #1, Space(14) & "Sri Annamalaiyar Pattasu Kadai"
    Print #1, Space(16) & "No 5A, RAILWAY STATION ROAD"
    Print #1, Space(19) & "METTUPALAYAM - 641301"
    Print #1, Space(13) & "CELL NO - 98437 44144, 93449 22273"
    Print #1, ""
    Print #1, Space(20) & "Purchase Period Wise Report Details"
    Print #1, "Report for the date from " & Format(DTPicker1.Value, "dd/mm/yyyy") & " to " & Format(DTPicker2.Value, "dd/mm/yyyy")
    Print #1, "--------------------------------------------------------------------------------"    ' 80 Characters
    Print #1, "ICode" & Space(2) & "Item Name" & Space(20) & Space(2) & "Item Type" & Space(2) & "Item Rate" & Space(2) & "Quantity" & Space(2) & " Total Amt"
    Print #1, "--------------------------------------------------------------------------------"
    While Not rs.EOF
        iicode = 5 - Len(rs.Fields("itemcode"))
        iiname = 29 - Len(rs.Fields("itemname"))
        iitype = 9 - Len(rs.Fields("itemtype"))
        iirate = 9 - Len(Format(rs.Fields("itemrate"), "0.00"))
        iiqty = 8 - Len(Format(rs.Fields("quantity"), "0.00"))
        iiamt = 10 - Len(Format(rs.Fields("itemamt"), "0.00"))
        
        Print #1, UCase(rs.Fields("itemcode")) & Space(iicode) & Space(2) & UCase(rs.Fields("itemname")) & Space(iiname) & Space(2) & UCase(rs.Fields("itemtype")) & Space(iitype) & Space(2) & Space(iirate) & Format(rs.Fields("itemrate"), "0.00") & Space(2) & Space(iiqty) & Format(rs.Fields("quantity"), "0.00") & Space(2) & Space(iiamt) & Format(rs.Fields("itemamt"), "0.00")
        rs.MoveNext
    Wend
    Print #1, "--------------------------------------------------------------------------------"
    
    igamt = 10 - Len(Format(gamt, "0.00"))
    idis = 10 - Len(Format(dis, "0.00"))
    itpa = 10 - Len(Format(Val(gamt) - Val(dis), "0.00"))
    
    Print #1, Space(51) & "Total Amount (Rs): " & Space(igamt) & Format(gamt, "0.00")
    Print #1, Space(51) & "Discount Amt (Rs): " & Space(idis) & Format(dis, "0.00")
    Print #1, Space(42) & "Total Purchase Amount (Rs): " & Space(itpa) & Format(Val(gamt) - Val(dis), "0.00")
    Close #1
    retval = Shell("notepad.exe rptpurchaseperiod.txt", vbMaximizedFocus)
End If

Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
Print #1, "TYPE rptpurchaseperiod.txt>PRN"
Print #1, "EXIT"
Close #1
retval = Shell(App.Path & "\print.bat", vbHide)

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Me.BackColor = RGB(238, 6, 147)
Call connect
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
