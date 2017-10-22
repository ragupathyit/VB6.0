VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmStudentMaster 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Student Master"
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
   Begin VB.TextBox txtregno 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   12600
      TabIndex        =   12
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
      MICON           =   "FrmStudentMaster.frx":0000
      PICN            =   "FrmStudentMaster.frx":001C
      PICH            =   "FrmStudentMaster.frx":072E
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
      TabIndex        =   6
      ToolTipText     =   "SAVE"
      Top             =   5880
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
      MICON           =   "FrmStudentMaster.frx":0E40
      PICN            =   "FrmStudentMaster.frx":0E5C
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
      TabIndex        =   7
      ToolTipText     =   "MODIFY"
      Top             =   5880
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
      MICON           =   "FrmStudentMaster.frx":156E
      PICN            =   "FrmStudentMaster.frx":158A
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
      TabIndex        =   9
      ToolTipText     =   "DELETE"
      Top             =   5880
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
      MICON           =   "FrmStudentMaster.frx":1C9C
      PICN            =   "FrmStudentMaster.frx":1CB8
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
      TabIndex        =   8
      ToolTipText     =   "CLEAR"
      Top             =   5880
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
      MICON           =   "FrmStudentMaster.frx":23CA
      PICN            =   "FrmStudentMaster.frx":23E6
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
      Height          =   4935
      Left            =   6360
      TabIndex        =   10
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8705
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Reg. No        |Student Name                                               "
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
   Begin MSForms.ComboBox cmbcsem 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4920
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
   Begin MSForms.ComboBox cmbcyear 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4200
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
   Begin VB.Label Label7 
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
      Left            =   480
      TabIndex        =   18
      Top             =   5040
      Width           =   1245
   End
   Begin MSForms.ComboBox cmbcname 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   3495
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6165;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
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
      Left            =   480
      TabIndex        =   17
      Top             =   2880
      Width           =   1575
   End
   Begin MSForms.ComboBox cmbccode 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
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
      Left            =   480
      TabIndex        =   16
      Top             =   4320
      Width           =   705
   End
   Begin VB.Label Label4 
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
      Left            =   480
      TabIndex        =   15
      Top             =   3600
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register No *"
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
      Left            =   480
      TabIndex        =   14
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name *"
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
      Left            =   480
      TabIndex        =   13
      Top             =   2160
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmStudentMaster.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT MASTER"
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
      TabIndex        =   11
      Top             =   240
      Width           =   3165
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
      Top             =   5760
      Width           =   13215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   6495
      Left            =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "FrmStudentMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_studentmaster where regno='" & Trim(UCase(txtregno.Text)) & "'"
MsgBox "Student Deleted Successfully", vbInformation, "Fees Collection"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmStudentMaster.Show
End Sub

Private Sub BtnModify_Click()
db.Execute "delete from tbl_studentmaster where regno='" & Trim(UCase(txtregno.Text)) & "'"
'-------------------Validation Starts Here-----------------------------
If txtregno.Text = "" Then
    MsgBox "Enter the Register No. Properly...", vbInformation, "Fees Collection"
    txtregno.SetFocus
ElseIf txtsname.Text = "" Then
    MsgBox "Enter the Student Name Properly...", vbInformation, "Fees Collection"
    txtsname.SetFocus
ElseIf cmbccode.Text = "" Then
    MsgBox "Select the Course Code Properly...", vbInformation, "Fees Collection"
    cmbccode.SetFocus
ElseIf cmbcname.Text = "" Then
    MsgBox "Select the Course Name Properly...", vbInformation, "Fees Collection"
    cmbcname.SetFocus
ElseIf cmbcyear.Text = "" Then
    MsgBox "Select the Student Year Properly...", vbInformation, "Fees Collection"
    cmbcyear.SetFocus
ElseIf cmbcsem.Text = "" Then
    MsgBox "Select the Student Semester Properly...", vbInformation, "Fees Collection"
    cmbcsem.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_studentmaster where regno='" & Trim(UCase(txtregno.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Student is Already Saved, Type Different Register No", vbInformation, "Fees Collection"
    Else
        rs.AddNew
        rs.Fields("regno") = Trim(UCase(txtregno.Text))
        rs.Fields("sname") = Trim(UCase(txtsname.Text))
        rs.Fields("ccode") = Trim(UCase(cmbccode.Text))
        rs.Fields("cname") = Trim(UCase(cmbcname.Text))
        rs.Fields("cyear") = Trim(UCase(cmbcyear.Text))
        rs.Fields("csem") = Trim(UCase(cmbcsem.Text))
        rs.Update
        rs.Close
        MsgBox "Student Modified Successfully", vbInformation, "Fees Collection"
    End If
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtregno.Text = "" Then
    MsgBox "Enter the Register No. Properly...", vbInformation, "Fees Collection"
    txtregno.SetFocus
ElseIf txtsname.Text = "" Then
    MsgBox "Enter the Student Name Properly...", vbInformation, "Fees Collection"
    txtsname.SetFocus
ElseIf cmbccode.Text = "" Then
    MsgBox "Select the Course Code Properly...", vbInformation, "Fees Collection"
    cmbccode.SetFocus
ElseIf cmbcname.Text = "" Then
    MsgBox "Select the Course Name Properly...", vbInformation, "Fees Collection"
    cmbcname.SetFocus
ElseIf cmbcyear.Text = "" Then
    MsgBox "Select the Student Year Properly...", vbInformation, "Fees Collection"
    cmbcyear.SetFocus
ElseIf cmbcsem.Text = "" Then
    MsgBox "Select the Student Semester Properly...", vbInformation, "Fees Collection"
    cmbcsem.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_studentmaster where regno='" & Trim(UCase(txtregno.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Student is Already Saved, Type Different Register No", vbInformation, "Fees Collection"
    Else
        rs.AddNew
        rs.Fields("regno") = Trim(UCase(txtregno.Text))
        rs.Fields("sname") = Trim(UCase(txtsname.Text))
        rs.Fields("ccode") = Trim(UCase(cmbccode.Text))
        rs.Fields("cname") = Trim(UCase(cmbcname.Text))
        rs.Fields("cyear") = Trim(UCase(cmbcyear.Text))
        rs.Fields("csem") = Trim(UCase(cmbcsem.Text))
        rs.Update
        rs.Close
        MsgBox "Student Saved Successfully", vbInformation, "Fees Collection"
    End If
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
stmt = "select regno,sname from tbl_studentmaster order by regno"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("regno") & vbTab & rs.Fields("sname")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub cmbccode_Click()
cmbcname.Text = cmbcname.List(cmbccode.ListIndex)
End Sub

Private Sub cmbcname_Click()
cmbccode.Text = cmbccode.List(cmbcname.ListIndex)
End Sub

Private Sub cmbccode_Change()
'cmbcname.Text = cmbcname.List(cmbccode.ListIndex)
End Sub

Private Sub cmbcname_Change()
'cmbccode.Text = cmbccode.List(cmbcname.ListIndex)
End Sub

Private Sub cmbcyear_Change()
cmbcsem.Clear
If cmbcyear.Text = "I" Then
    cmbcsem.AddItem "I"
    cmbcsem.AddItem "II"
ElseIf cmbcyear.Text = "II" Then
    cmbcsem.AddItem "III"
    cmbcsem.AddItem "IV"
End If
End Sub

Private Sub Form_Load()
Call connect
Call Fill
Call fillcombo

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

cmbcyear.AddItem "I"
cmbcyear.AddItem "II"

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Public Sub fillcombo()
If rs.State = 1 Then rs.Close
rs.Open "select ccode,sname from tbl_coursemaster order by ccode", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    While Not rs.EOF
        cmbccode.AddItem Trim(rs.Fields("ccode"))
        cmbcname.AddItem Trim(rs.Fields("sname"))
        rs.MoveNext
    Wend
End If
rs.Close
End Sub

Private Sub MsGrid_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_studentmaster where regno='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtregno.Text = rs1.Fields("regno")
    txtsname.Text = rs1.Fields("sname")
    cmbccode.Text = rs1.Fields("ccode")
    cmbcname.Text = rs1.Fields("cname")
    cmbcyear.Text = rs1.Fields("cyear")
    cmbcsem.Text = rs1.Fields("csem")
End If

txtregno.SetFocus
BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtregno_Change()
stmt = "select regno,sname from tbl_studentmaster where regno like '" & Trim(txtregno.Text) & "%' order by regno"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("regno") & vbTab & rs.Fields("sname")
        rs.MoveNext
    Loop
End If
rs.Close
End Sub

Private Sub txtregno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsname.SetFocus
    txtsname.SelStart = 0
    txtsname.SelLength = Len(txtsname.Text)    'select the text
End If
End Sub

Private Sub txtsname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbccode.SetFocus
    cmbccode.SelStart = 0
    cmbccode.SelLength = Len(cmbccode.Text)    'select the text
End If
End Sub

Private Sub cmbccode_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    cmbcname.SetFocus
    cmbcname.SelStart = 0
    cmbcname.SelLength = Len(cmbcname.Text)    'select the text
End If
End Sub

Private Sub cmbcname_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    cmbcyear.SetFocus
    cmbcyear.SelStart = 0
    cmbcyear.SelLength = Len(cmbcyear.Text)    'select the text
End If
End Sub

Private Sub cmbcyear_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    cmbcsem.SetFocus
    cmbcsem.SelStart = 0
    cmbcsem.SelLength = Len(cmbcsem.Text)    'select the text
End If
End Sub

Private Sub cmbcsem_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
