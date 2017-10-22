VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmReportSales 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbcustname 
      BeginProperty Font 
         Name            =   "Tamil-Ananthi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   10
      Top             =   1080
      Width           =   4335
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   10200
      TabIndex        =   1
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
      MICON           =   "FrmReportSales.frx":0000
      PICN            =   "FrmReportSales.frx":001C
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
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Bill. No  |Customer Name                                          |Sales Date |Bill Amt        "
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
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "SAVE"
      Top             =   7440
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReportSales.frx":072E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtp_fdate 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   89653251
      CurrentDate     =   41799
   End
   Begin MSComCtl2.DTPicker dtp_tdate 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   89653251
      CurrentDate     =   41799
   End
   Begin Project1.Button BtnClick 
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      ToolTipText     =   "SAVE"
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Click  "
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
      MICON           =   "FrmReportSales.frx":074A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name :"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   7320
      Width           =   10935
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmReportSales.frx":0766
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALES REPORT"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2595
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "FrmReportSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClick_Click()
If rs.State = 1 Then rs.Close
rs.Open "Select * from tbl_sales where sdate between #" & Format(dtp_fdate.Value, "MM/DD/YYYY") & "# and #" & Format(dtp_tdate.Value, "MM/DD/YYYY") & "#", db, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "No Records Check the Date", vbInformation, "Press Management"
    Exit Sub
End If

If rs.State = 1 Then rs.Close
If cmbcustname.Text = "" Then
    rs.Open "select distinct billno,custname,concernname,custtype,sdate,totamt from tbl_sales where sdate between #" & Format(dtp_fdate.Value, "MM/DD/YYYY") & "# and #" & Format(dtp_tdate.Value, "MM/DD/YYYY") & "# order by billno", db, adOpenDynamic, adLockOptimistic
Else
    rs.Open "select distinct billno,custname,concernname,custtype,sdate,totamt from tbl_sales where custname='" & Trim(cmbcustname.Text) & "' and sdate between #" & Format(dtp_fdate.Value, "MM/DD/YYYY") & "# and #" & Format(dtp_tdate.Value, "MM/DD/YYYY") & "# order by billno", db, adOpenDynamic, adLockOptimistic
End If
MSGrid.Rows = 1
While Not rs.EOF
    MSGrid.AddItem rs.Fields("billno") & vbTab & rs.Fields("custname") & vbTab & rs.Fields("concernname") & vbTab & rs.Fields("custtype") & vbTab & rs.Fields("sdate") & vbTab & Format(rs.Fields("totamt"), "0.00")
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub BtnReport_Click()
If rs.State = 1 Then rs.Close
rs.Open "Select distinct sum(totamt) from tbl_sales where sdate between #" & Format(dtp_fdate.Value, "MM/DD/YYYY") & "# and #" & Format(dtp_tdate.Value, "MM/DD/YYYY") & "#", db, adOpenDynamic, adLockOptimistic
gtotamt = rs.Fields(0)

'----------Notepad print------------------
Open App.Path & "\rptsales.txt" For Output As #1

Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
Print #1, ""
'Print #1, Space(14) & "Press Management"
Print #1, ""
Print #1, Space(12) & "Sales Period Wise Report Details"
Print #1, "Report for the date from " & Format(dtp_fdate.Value, "dd/mm/yyyy") & " to " & Format(dtp_tdate.Value, "dd/mm/yyyy")
Print #1, "--------------------------------------------------------------------------------"    ' 80 Characters
Print #1, "Bill No " & Space(1) & "Customer Name" & Space(12) & Space(1) & "Concern Name" & Space(1) & "C.Type    " & Space(1) & "Sales Date" & Space(1) & "  Tot. Amt"
Print #1, "--------------------------------------------------------------------------------"

Set i = Nothing
For i = 1 To MSGrid.Rows - 1
    ibillno = 8 - Len(MSGrid.TextMatrix(i, 0))
    icustname = 25 - Len(Mid(MSGrid.TextMatrix(i, 1), 1, 25))
    iconcname = 12 - Len(Mid(MSGrid.TextMatrix(i, 2), 1, 12))
    ictype = 10 - Len(Mid(MSGrid.TextMatrix(i, 3), 1, 10))
    isalesdate = 10 - Len(MSGrid.TextMatrix(i, 4))
    itotamt = 10 - Len(MSGrid.TextMatrix(i, 5))
    Print #1, MSGrid.TextMatrix(i, 0) & Space(ibillno) & Space(1) & UCase(Mid(MSGrid.TextMatrix(i, 1), 1, 25)) & Space(icustname) & Space(1) & UCase(Mid(MSGrid.TextMatrix(i, 2), 1, 12)) & Space(iconcname) & Space(1) & UCase(Mid(MSGrid.TextMatrix(i, 3), 1, 10)) & Space(ictype) & Space(1) & MSGrid.TextMatrix(i, 4) & Space(isalesdate) & Space(1) & Space(itotamt) & MSGrid.TextMatrix(i, 5)
Next i

Print #1, "--------------------------------------------------------------------------------"
igtotamt = 10 - Len(Format(gtotamt, "0.00"))

Print #1, Space(70) & Space(igtotamt) & Format(gtotamt, "0.00")
Close #1
retval = Shell("notepad.exe rptsales.txt", vbMaximizedFocus)
End Sub

Private Sub Form_Load()
Call connect

dtp_fdate.Value = Date
dtp_tdate.Value = Date

stmt = "select customername from tbl_customer order by cid"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        cmbcustname.AddItem rs.Fields("customername")
        rs.MoveNext
    Loop
End If
rs.Close
End Sub
