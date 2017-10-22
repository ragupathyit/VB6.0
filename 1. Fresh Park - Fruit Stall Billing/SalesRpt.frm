VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SalesRpt 
   Caption         =   "Sales Day Wise Report"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   8400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
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
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "REPORT"
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
      Top             =   2040
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
      Format          =   89260033
      CurrentDate     =   40537
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALES DAY WISE REPORT"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   4080
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   3600
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select the Date to Take Report"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   3450
   End
End
Attribute VB_Name = "SalesRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub CmdReport_Click()

If rs.State = 1 Then rs.Close
rs.Open "Select * from tbl_sales where salesdate=#" & DTPicker1.Value & "#", db, adOpenDynamic, adLockOptimistic
If rs.EOF Then
    MsgBox "No Records Check the Date", vbInformation, "Fresh Park"
    Exit Sub
End If

If rs.State = 1 Then rs.Close
rs.Open "Select sum(itemamt) from tbl_sales where salesdate=#" & DTPicker1.Value & "#", db, adOpenDynamic, adLockOptimistic
bamt = rs.Fields(0)
ibamt = 13 - Len(Format(bamt, "0.00"))

If rs.State = 1 Then rs.Close
stmt = "select distinct billno,salesdate,billamt from tbl_sales where salesdate=#" & DTPicker1.Value & "#"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\report.txt" For Output As #1
    
    Print #1, Space(13) & "FRESH PARK"
    Print #1, Space(10) & "Palamudir Nilayam"
    Print #1, Space(8) & "CBE.Road, Mettupalayam"
    Print #1, Space(14) & "PH-224977"
    Print #1, ""
    Print #1, Space(9) & "Sales Report Details"
    Print #1, ""
    Print #1, "Report for the date of: " & Format(DTPicker1.Value, "DD/MM/YYYY")
    Print #1, ""
    Print #1, "==========================================================="
    Print #1, "Sl. No" & Space(2) & "Bill No" & Space(2) & "Sales Date" & Space(15) & "Total" & Space(5) & "Type"
    Print #1, "==========================================================="
    i = 1
    While Not rs.EOF
        ii = 6 - Len(i)
        ibno = 7 - Len(rs.Fields("billno"))
        isalesd = 20 - Len(Format(rs.Fields("salesdate"), "DD/MM/YYYY"))
        itotal = 13 - Len(Format(rs.Fields("billamt"), "0.00"))
        
        Print #1, i & Space(ii) & Space(2) & rs.Fields("billno") & Space(ibno) & Space(2) & Format(rs.Fields("salesdate"), "DD/MM/YYYY") & Space(isalesd) & Space(itotal) & Format(rs.Fields("billamt"), "0.00") & Space(2) & "Cash"
        
        i = i + 1
        rs.MoveNext
    Wend
    Print #1, "-----------------------------------------------------------"
    Print #1, Space(29) & "Total : " & Space(ibamt) & Format(bamt, "0.00")
    Print #1, Space(38) & "------------"
    Close #1
    retval = Shell("notepad.exe report.txt", vbMaximizedFocus)
End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

DTPicker1.Value = Format(Date, "DD/MM/YYYY")

End Sub
