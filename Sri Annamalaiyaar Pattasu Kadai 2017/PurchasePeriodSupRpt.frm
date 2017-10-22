VERSION 5.00
Begin VB.Form PurchasePeriodSupRpt 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Purchase Supplier Wise Report"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbsupname 
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
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
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
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3000
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
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE SUPPLIER WISE REPORT"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   7200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Supplier Name"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   3990
   End
End
Attribute VB_Name = "PurchasePeriodSupRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdReport_Click()
If rs.State = 1 Then rs.Close
rs.Open "Select sum(itemamt) from tbl_purchase where supname='" & cmbsupname.Text & "'", db, adOpenDynamic, adLockOptimistic
gamt = rs.Fields(0)

If rs.State = 1 Then rs.Close
rs.Open "Select distinct pid,discount from tbl_purchase where supname='" & cmbsupname.Text & "'", db, adOpenDynamic, adLockOptimistic
dis = 0
While Not rs.EOF
    dis = Val(dis) + Val(rs.Fields("discount"))
    rs.MoveNext
Wend

stmt = "select itemcode,itemname,itemtype,itemrate,quantity,itemamt,gridtotamt,discount,totamt from tbl_purchase where supname='" & cmbsupname.Text & "'"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\rptpurchasesupplier.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, ""
    Print #1, Space(21) & "Sri Annamalaiyar Pattasu Kadai"
    Print #1, Space(23) & "No 5A, RAILWAY STATION ROAD"
    Print #1, Space(26) & "METTUPALAYAM - 641301"
    Print #1, Space(20) & "CELL NO - 98437 44144, 93449 22273"
    Print #1, ""
    Print #1, Space(20) & "Purchase Supplier Wise Report Details"
    Print #1, "Report is based on the supplier " & cmbsupname.Text
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
    retval = Shell("notepad.exe rptpurchasesupplier.txt", vbMaximizedFocus)
End If

Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
Print #1, "TYPE rptpurchasesupplier.txt>PRN"
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

If rs.State = 1 Then rs.Close
rs.Open "Select suppliername from tbl_suppliermaster", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    cmbsupname.AddItem rs.Fields("suppliername")
    rs.MoveNext
Wend
rs.Close

End Sub
