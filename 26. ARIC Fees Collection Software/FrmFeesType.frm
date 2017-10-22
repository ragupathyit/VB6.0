VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFeesType 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Fees Name Master"
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
   Begin VB.TextBox txtfamt 
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
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtfcode 
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
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtfname 
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
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   12360
      TabIndex        =   8
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
      MICON           =   "FrmFeesType.frx":0000
      PICN            =   "FrmFeesType.frx":001C
      PICH            =   "FrmFeesType.frx":072E
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
      TabIndex        =   1
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
      MICON           =   "FrmFeesType.frx":0E40
      PICN            =   "FrmFeesType.frx":0E5C
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
      TabIndex        =   2
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
      MICON           =   "FrmFeesType.frx":156E
      PICN            =   "FrmFeesType.frx":158A
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
      TabIndex        =   4
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
      MICON           =   "FrmFeesType.frx":1C9C
      PICN            =   "FrmFeesType.frx":1CB8
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
      TabIndex        =   3
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
      MICON           =   "FrmFeesType.frx":23CA
      PICN            =   "FrmFeesType.frx":23E6
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
      TabIndex        =   5
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "F. Code  |Fees Name                                                      "
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Amount *"
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
      Top             =   3600
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Code"
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
      TabIndex        =   10
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Name *"
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmFeesType.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FEES NAME MASTER"
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
      TabIndex        =   7
      Top             =   240
      Width           =   3510
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   12975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   4920
      Width           =   12975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "FrmFeesType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_feesname where fcode=" & Trim(Val(txtfcode.Text))
MsgBox "Fees Name Deleted Successfully", vbInformation, "Fees Collection"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmFeesType.Show
End Sub

Private Sub BtnModify_Click()
db.Execute "delete from tbl_feesname where fcode=" & Trim(Val(txtfcode.Text))
'-------------------Validation Starts Here-----------------------------
If txtfname.Text = "" Then
    MsgBox "Enter the Fees Name Properly...", vbInformation, "Fees Collection"
    txtfname.SetFocus
ElseIf txtfamt.Text = "" Then
    MsgBox "Enter the Fees Amount Properly...", vbInformation, "Fees Collection"
    txtfamt.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_feesname", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("fcode") = Trim(Val(txtfcode.Text))
        rs.Fields("fname") = Trim(UCase(txtfname.Text))
        rs.Fields("famt") = Trim(Val(txtfamt.Text))
    rs.Update
    rs.Close
    MsgBox "Fees Name Modified Successfully", vbInformation, "Fees Collection"
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtfname.Text = "" Then
    MsgBox "Enter the Fees Name Properly...", vbInformation, "Fees Collection"
    txtfname.SetFocus
ElseIf txtfamt.Text = "" Then
    MsgBox "Enter the Fees Amount Properly...", vbInformation, "Fees Collection"
    txtfamt.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_feesname", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("fcode") = Trim(Val(txtfcode.Text))
        rs.Fields("fname") = Trim(UCase(txtfname.Text))
        rs.Fields("famt") = Trim(Val(txtfamt.Text))
    rs.Update
    rs.Close
    MsgBox "Fees Name Saved Successfully", vbInformation, "Fees Collection"
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
stmt = "select * from tbl_feesname order by fcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("fcode") & vbTab & rs.Fields("fname")
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

If rs.State = 1 Then rs.Close
rs.Open "select fcode from tbl_feesname order by fcode", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtfcode.Text = Val(rs.Fields("fcode")) + 1
Else
    txtfcode.Text = 1
End If
rs.Close

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_feesname where fcode=" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtfcode.Text = rs1.Fields("fcode")
    txtfname.Text = rs1.Fields("fname")
    txtfamt.Text = rs1.Fields("famt")
End If

txtfname.SetFocus
txtfname.SelStart = 0
txtfname.SelLength = Len(txtfname.Text)    'select the text

BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtfamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub

Private Sub txtfname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfamt.SetFocus
    txtfamt.SelStart = 0
    txtfamt.SelLength = Len(txtfamt.Text)    'select the text
End If
End Sub
