VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFeesStructure 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Fees Structure"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmFeesStructure.frx":0000
      Left            =   9600
      List            =   "FrmFeesStructure.frx":0002
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmFeesStructure.frx":0004
      Left            =   6600
      List            =   "FrmFeesStructure.frx":0006
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbcustid 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmFeesStructure.frx":0008
      Left            =   1920
      List            =   "FrmFeesStructure.frx":000A
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbcustname 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmFeesStructure.frx":000C
      Left            =   2880
      List            =   "FrmFeesStructure.frx":000E
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   10
      Text            =   "0"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtbillno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   10800
      TabIndex        =   3
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
      MICON           =   "FrmFeesStructure.frx":0010
      PICN            =   "FrmFeesStructure.frx":002C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnSave 
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      ToolTipText     =   "SAVE"
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Save   "
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
      MICON           =   "FrmFeesStructure.frx":073E
      PICN            =   "FrmFeesStructure.frx":075A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnDelete 
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      ToolTipText     =   "DELETE"
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Delete"
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
      MICON           =   "FrmFeesStructure.frx":0E6C
      PICN            =   "FrmFeesStructure.frx":0E88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnClear 
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      ToolTipText     =   "CLEAR"
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Clear  "
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
      MICON           =   "FrmFeesStructure.frx":159A
      PICN            =   "FrmFeesStructure.frx":15B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnPrevious 
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "Previous"
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
      MICON           =   "FrmFeesStructure.frx":1CC8
      PICN            =   "FrmFeesStructure.frx":1CE4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnNext 
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      ToolTipText     =   "Next"
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
      MICON           =   "FrmFeesStructure.frx":23F6
      PICN            =   "FrmFeesStructure.frx":2412
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstpapername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      ItemData        =   "FrmFeesStructure.frx":2B24
      Left            =   120
      List            =   "FrmFeesStructure.frx":2B26
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1800
      Width           =   3825
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5415
      Left            =   3960
      TabIndex        =   13
      Top             =   1800
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Fees Type Name                                                   |Amount        "
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
   Begin Project1.Button BtnModify 
      Height          =   495
      Left            =   4080
      TabIndex        =   21
      ToolTipText     =   "MODIFY"
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Modify"
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
      MICON           =   "FrmFeesStructure.frx":2B28
      PICN            =   "FrmFeesStructure.frx":2B44
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year *"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name *"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Name"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   8760
      TabIndex        =   9
      Top             =   7320
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semester *"
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
      Left            =   8160
      TabIndex        =   6
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees ID"
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
      Left            =   8640
      TabIndex        =   5
      Top             =   0
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmFeesStructure.frx":3256
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FEES STRUCTURE"
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
      TabIndex        =   2
      Top             =   240
      Width           =   3030
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   7800
      Width           =   11415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   8655
      Left            =   0
      Top             =   -120
      Width           =   11415
   End
End
Attribute VB_Name = "FrmFeesStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
'db.Execute "update tbl_sales set billcancel='Y' where billno=" & Val(txtbillno.Text)
'MsgBox "Bill Cancelled Successfully", vbInformation, "Fees Collection"
'Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmFeesStructure.Show
End Sub

Private Sub BtnNext_Click()
'If rs.State = 1 Then rs.Close
'rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) + 1, db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    Call navigation
'
'    BtnSave.Enabled = False
'    BtnBill.Enabled = True
'    BtnDelete.Enabled = True
'Else
'    Call BtnClear_Click
'End If
End Sub

Private Sub BtnPrevious_Click()
'If rs.State = 1 Then rs.Close
'rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    Call navigation
'
'    BtnSave.Enabled = False
'    BtnBill.Enabled = True
'    BtnDelete.Enabled = True
'End If
End Sub

Function navigation()
'txtbillno.Text = ""
'cmbcustname.Text = ""
'txtadvance.Text = ""
'txttotamt.Text = ""
'txtpayamt.Text = ""
'MSGrid.Rows = 2
'MSGrid.TextMatrix(1, 0) = ""
'MSGrid.TextMatrix(1, 1) = ""
'MSGrid.TextMatrix(1, 2) = ""
'
'If rs.Fields("billcancel") = "Y" Then
'    lblcancel.Caption = "CANCEL BILL"
'    GoTo nxt:
'Else
'    lblcancel.Caption = ""
'End If
'
'txtbillno.Text = rs.Fields("billno")
'cmbcustid.Text = IIf(IsNull(rs.Fields("cid")), "", rs.Fields("cid"))
'cmbcustname.Text = rs.Fields("custname")
'dtp_sdate.Value = rs.Fields("sdate")
'txtadvance.Text = rs.Fields("advamt")
'txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
'txtobalance.Text = Format(rs.Fields("obalance"), "0.00")
'txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
'
'i = 1
'While Not rs.EOF
'    MSGrid.TextMatrix(i, 0) = rs.Fields("pid")
'    MSGrid.TextMatrix(i, 1) = rs.Fields("papername")
'    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("prate"), "0.00")
'    MSGrid.Rows = MSGrid.Rows + 1
'    i = i + 1
'    rs.MoveNext
'Wend
'rs.Close
'
'nxt:
'
End Function

Private Sub BtnSave_Click()
''-------------------Validation Starts Here-----------------------------
'If cmbcustname.Text = "" Then
'    MsgBox "Select the Customer Name Properly...", vbInformation, "Fees Collection"
'    cmbcustname.SetFocus
'Else
''-------------------Validation Ends Here-------------------------------
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
'    For i = 1 To MSGrid.Rows - 1
'        rs.AddNew
'        rs.Fields("billno") = Val(txtbillno.Text)
'        rs.Fields("cid") = Val(cmbcustid.Text)
'        rs.Fields("custname") = Trim(cmbcustname.Text)
'        rs.Fields("sdate") = dtp_sdate.Value
'        rs.Fields("advamt") = Val(txtadvance.Text)
'        rs.Fields("pid") = MSGrid.TextMatrix(i, 0)
'        rs.Fields("papername") = MSGrid.TextMatrix(i, 1)
'        rs.Fields("prate") = Format(Val(MSGrid.TextMatrix(i, 2)), "0.00")
'        rs.Fields("totamt") = Format(Round(Val(txttotamt.Text)), "0.00")
'        rs.Fields("obalance") = Format(Round(Val(txtobalance.Text)), "0.00")
'        rs.Fields("payamt") = Format(Round(Val(txtpayamt.Text)), "0.00")
'        rs.Update
'    Next i
'    rs.Close
'
'    If Val(txtpayamt.Text) < Val(txttotamt.Text) Then
'        '====================Sales Balance====================
'        If rs.State = 1 Then rs.Close
'        rs.Open "select * from tbl_salesbalance", db, adOpenDynamic, adLockOptimistic
'        rs.AddNew
'            rs.Fields("billno") = Val(txtbillno.Text)
'            rs.Fields("salesdate") = dtp_sdate.Value
'            rs.Fields("cid") = Val(cmbcustid.Text)
'            rs.Fields("custname") = UCase(cmbcustname.Text)
'            rs.Fields("balamt") = Format(Round(Val(txttotamt.Text) - Val(txtpayamt.Text)), "0.00")
'            rs.Fields("obalance") = Format(Round(Val(txtobalance.Text)), "0.00")
'            rs.Fields("totamt") = Format(Round(Val(txttotamt.Text)), "0.00")
'            rs.Fields("payamt") = Format(Round(Val(txtpayamt.Text)), "0.00")
'            rs.Fields("baldesc") = Format(Round(Val(txttotamt.Text)), "0.00") & "-" & Format(Round(Val(txtpayamt.Text)), "0.00")
'        rs.Update
'        rs.Close
'        '====================Sales Balance====================
'    Else
'        If rs.State = 1 Then rs.Close
'        rs.Open "select * from tbl_salesbalance where custname='" & Trim(cmbcustname.Text) & "' order by billno", db, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
'            rs.MoveLast
'            If rs.Fields("balamt") = 0 Then
'                db.Execute "delete from tbl_salesbalance where custname='" & Trim(cmbcustname.Text) & "'"
'            End If
'        End If
'    End If
'
'    MsgBox "Bill Details Saved Successfully...", vbInformation, "Fees Collection"
'
'    'Call BtnBill_Click
'    Call BtnClear_Click
'End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

'Private Sub cmbcustid_Click()
'cmbcustname.Text = cmbcustname.List(cmbcustid.ListIndex)
'cmbcustname.SetFocus
'End Sub
'
'Private Sub cmbcustname_Click()
'cmbcustid.Text = cmbcustid.List(cmbcustname.ListIndex)
'lstpapername.SetFocus
'End Sub

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select billno from tbl_sales order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = rs.Fields("billno") + 1
Else
    txtbillno.Text = 1
End If
rs.Close

BtnSave.Enabled = True
BtnDelete.Enabled = False
End Sub

Private Function Fill()
'stmt = "select * from tbl_papermaster order by pid"
'If rs.State = 1 Then rs.Close
'rs.Open stmt, db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    rs.MoveFirst
'    Do While Not rs.EOF
'        lstpapername.AddItem rs.Fields("papername")
'        rs.MoveNext
'    Loop
'End If
'rs.Close
'
'rs.Open "select cid,customername from tbl_customer order by cid", db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    rs.MoveFirst
'    Do While Not rs.EOF
'        cmbcustid.AddItem rs.Fields("cid")
'        cmbcustname.AddItem rs.Fields("customername")
'        rs.MoveNext
'    Loop
'End If
End Function

'Private Sub lstpapername_ItemCheck(Item As Integer)
'If lstpapername.Selected(Item) = True Then
'    If MSGrid.TextMatrix(1, 0) = "" Then
'        MSGrid.Rows = 1
'    End If
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from tbl_papermaster where papername='" & Trim(lstpapername.List(Item)) & "'", db, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then
'        MSGrid.AddItem rs.Fields("pid") & vbTab & rs.Fields("papername") & vbTab & Format(rs.Fields("prate"), "0.00")
'    End If
'Else
'    For i = 1 To MSGrid.Rows - 1
'        If MSGrid.TextMatrix(i, 1) = lstpapername.List(Item) Then
'            If MSGrid.Rows = 2 Then
'                MSGrid.TextMatrix(1, 0) = ""
'            Else
'                MSGrid.RemoveItem i
'            End If
'            Exit For
'        End If
'    Next
'End If
''--------------Calculating the total amount
'txttotamt.Text = 0
'For i = 1 To MSGrid.Rows - 1
'    txttotamt.Text = Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 2))
'Next
'txttotamt.Text = Format(Val(txttotamt.Text) + Val(txtobalance.Text), "0.00")
'
'txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
'
'txtpayamt.SetFocus
'txtpayamt.SelStart = 0
'txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
'End Sub
'
'Private Sub cmbcustname_LostFocus()
'If cmbcustname.Text <> "" Then
'    '--------------------------Customer Advance Amount-----------------------------------
'    stmt = "select advance from tbl_customer where cid=" & Trim(Val(cmbcustid.Text)) & " and customername='" & Trim(cmbcustname.Text) & "'"
'    If rs1.State = 1 Then rs1.Close
'    rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
'    If Not rs1.EOF Then
'        txtadvance.Text = Format(Val(rs1.Fields("advance")), "0.00")
'    End If
'    rs1.Close
'
'    '--------------------------Customer Old Balance Amount--------------------------------
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from tbl_salesbalance where cid=" & Trim(Val(cmbcustid.Text)) & " and custname='" & cmbcustname.Text & "' order by billno", db, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then
'        rs.MoveLast
'        txtobalance.Text = Format(Val(rs.Fields("balamt")), "0.00")
'    Else
'        txtobalance.Text = "0.00"
'    End If
'    rs.Close
'End If
'End Sub
'
'Private Sub txtpayamt_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    BtnSave.SetFocus
'End If
'End Sub
'
'Private Sub txtpayamt_LostFocus()
'txtpayamt.Text = Format(Val(txtpayamt.Text), "0.00")
'End Sub
