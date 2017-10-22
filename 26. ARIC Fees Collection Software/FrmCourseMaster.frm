VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCourseMaster 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Course Master"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14130
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtsname 
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
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtccode 
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
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtcname 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   12600
      TabIndex        =   10
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
      MICON           =   "FrmCourseMaster.frx":0000
      PICN            =   "FrmCourseMaster.frx":001C
      PICH            =   "FrmCourseMaster.frx":072E
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
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "SAVE"
      Top             =   5040
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
      MICON           =   "FrmCourseMaster.frx":0E40
      PICN            =   "FrmCourseMaster.frx":0E5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnModify 
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "MODIFY"
      Top             =   5040
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
      MICON           =   "FrmCourseMaster.frx":156E
      PICN            =   "FrmCourseMaster.frx":158A
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
      Left            =   8520
      TabIndex        =   7
      ToolTipText     =   "DELETE"
      Top             =   5040
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
      MICON           =   "FrmCourseMaster.frx":1C9C
      PICN            =   "FrmCourseMaster.frx":1CB8
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
      Left            =   6480
      TabIndex        =   6
      ToolTipText     =   "CLEAR"
      Top             =   5040
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
      MICON           =   "FrmCourseMaster.frx":23CA
      PICN            =   "FrmCourseMaster.frx":23E6
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
      Height          =   4095
      Left            =   6360
      TabIndex        =   8
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "C. Code  |Course Name                                                      "
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
   Begin MSForms.ComboBox cmbduration 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3413;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name *"
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
      Left            =   600
      TabIndex        =   14
      Top             =   3960
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duration *"
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
      Left            =   600
      TabIndex        =   13
      Top             =   3120
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code *"
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
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
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
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmCourseMaster.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COURSE MASTER"
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
      TabIndex        =   9
      Top             =   240
      Width           =   2955
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   4920
      Width           =   13215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "FrmCourseMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_coursemaster where ccode='" & Trim(UCase(txtccode.Text)) & "'"
MsgBox "Course Deleted Successfully", vbInformation, "Fees Collection"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmCourseMaster.Show
End Sub

Private Sub BtnModify_Click()
db.Execute "delete from tbl_coursemaster where ccode='" & Trim(UCase(txtccode.Text)) & "'"
'-------------------Validation Starts Here-----------------------------
If txtccode.Text = "" Then
    MsgBox "Enter the Course Code Properly...", vbInformation, "Fees Collection"
    txtccode.SetFocus
ElseIf txtcname.Text = "" Then
    MsgBox "Enter the Course Name Properly...", vbInformation, "Fees Collection"
    txtcname.SetFocus
ElseIf cmbduration.Text = "" Then
    MsgBox "Select the Course Duration Properly...", vbInformation, "Fees Collection"
    cmbduration.SetFocus
ElseIf txtsname.Text = "" Then
    MsgBox "Enter the Course Short Name Properly...", vbInformation, "Fees Collection"
    txtsname.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_coursemaster where ccode='" & Trim(UCase(txtccode.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Course is Already Saved, Type Different Course Code", vbInformation, "Fees Collection"
    Else
        rs.AddNew
        rs.Fields("ccode") = Trim(UCase(txtccode.Text))
        rs.Fields("cname") = Trim(UCase(txtcname.Text))
        rs.Fields("duration") = Trim(UCase(cmbduration.Text))
        rs.Fields("sname") = Trim(UCase(txtsname.Text))
        rs.Update
        rs.Close
        MsgBox "Course Modified Successfully", vbInformation, "Fees Collection"
    End If
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtccode.Text = "" Then
    MsgBox "Enter the Course Code Properly...", vbInformation, "Fees Collection"
    txtccode.SetFocus
ElseIf txtcname.Text = "" Then
    MsgBox "Enter the Course Name Properly...", vbInformation, "Fees Collection"
    txtcname.SetFocus
ElseIf cmbduration.Text = "" Then
    MsgBox "Select the Course Duration Properly...", vbInformation, "Fees Collection"
    cmbduration.SetFocus
ElseIf txtsname.Text = "" Then
    MsgBox "Enter the Course Short Name Properly...", vbInformation, "Fees Collection"
    txtsname.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_coursemaster where ccode='" & Trim(UCase(txtccode.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Course is Already Saved, Type Different Course Code", vbInformation, "Fees Collection"
    Else
        rs.AddNew
        rs.Fields("ccode") = Trim(UCase(txtccode.Text))
        rs.Fields("cname") = Trim(UCase(txtcname.Text))
        rs.Fields("duration") = Trim(UCase(cmbduration.Text))
        rs.Fields("sname") = Trim(UCase(txtsname.Text))
        rs.Update
        rs.Close
        MsgBox "Course Saved Successfully", vbInformation, "Fees Collection"
    End If
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
stmt = "select ccode,cname from tbl_coursemaster order by ccode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("ccode") & vbTab & rs.Fields("cname")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub Form_Load()
Call connect
Call Fill

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

cmbduration.AddItem "1 Year"
cmbduration.AddItem "2 Year"
cmbduration.AddItem "3 Year"
cmbduration.AddItem "4 Year"
cmbduration.AddItem "5 Year"

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_coursemaster where ccode='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtccode.Text = rs1.Fields("ccode")
    txtcname.Text = rs1.Fields("cname")
    cmbduration.Text = rs1.Fields("duration")
    txtsname.Text = rs1.Fields("sname")
End If

txtcname.SetFocus
BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtccode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcname.SetFocus
    txtcname.SelStart = 0
    txtcname.SelLength = Len(txtcname.Text)    'select the text
End If
End Sub

Private Sub txtcname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbduration.SetFocus
    cmbduration.SelStart = 0
    cmbduration.SelLength = Len(cmbduration.Text)    'select the text
End If
End Sub

Private Sub cmbduration_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    txtsname.SetFocus
    txtsname.SelStart = 0
    txtsname.SelLength = Len(txtsname.Text)    'select the text
End If
End Sub

Private Sub txtsname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
