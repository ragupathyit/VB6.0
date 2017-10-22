VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmTableMaster 
   BackColor       =   &H00400040&
   Caption         =   "Dining Table Master"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
   Begin VB.TextBox txttableno 
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
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txttableid 
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   6960
      TabIndex        =   6
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
      MICON           =   "FrmTableMaster.frx":0000
      PICN            =   "FrmTableMaster.frx":001C
      PICH            =   "FrmTableMaster.frx":072E
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
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "SAVE"
      Top             =   5640
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
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTableMaster.frx":0E40
      PICN            =   "FrmTableMaster.frx":0E5C
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
      Left            =   2400
      TabIndex        =   8
      ToolTipText     =   "MODIFY"
      Top             =   5640
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
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTableMaster.frx":156E
      PICN            =   "FrmTableMaster.frx":158A
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
      Left            =   4080
      TabIndex        =   9
      ToolTipText     =   "CLEAR"
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Refresh"
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
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTableMaster.frx":1C9C
      PICN            =   "FrmTableMaster.frx":1CB8
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
      Left            =   5760
      TabIndex        =   10
      ToolTipText     =   "DELETE"
      Top             =   5640
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
      BCOL            =   12583104
      BCOLO           =   12583104
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTableMaster.frx":23CA
      PICN            =   "FrmTableMaster.frx":23E6
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
      Height          =   6375
      Left            =   7560
      TabIndex        =   2
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Table Number            "
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   5400
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DINING TABLE  MASTER"
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
      TabIndex        =   5
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmTableMaster.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table No.  *"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table Id"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "FrmTableMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClear_Click()
Unload Me
FrmTableMaster.Show
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub BtnDelete_Click()
db.Execute "delete from tbl_tablemaster where tableid=" & Val(txttableid.Text)
db.Execute "delete from tbl_runorder where tableid=" & Val(txttableid.Text)
MsgBox "Record is Deleted Successfully", vbInformation, "Sri Saravana Bhavan"
Call BtnClear_Click
End Sub

Private Sub BtnModify_Click()
Call update
MsgBox "Record Modified Successfully", vbInformation, "Sri Saravana Bhavan"
Call BtnClear_Click
End Sub

Private Sub BtnSave_Click()
Call update
MsgBox "Record Saved Successfully", vbInformation, "Sri Saravana Bhavan"
Call BtnClear_Click
End Sub

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select tableid from tbl_tablemaster order by tableid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txttableid.Text = rs.Fields("tableid") + 1
Else
    txttableid.Text = 1
End If
rs.Close

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

Me.Show
txttableno.SetFocus

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Function Fill()
stmt = "select * from tbl_tablemaster order by tableno"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        'MSGrid.AddItem rs.Fields("tableid") & vbTab & rs.Fields("tableno")
        MSGrid.AddItem rs.Fields("tableno")
        rs.MoveNext
    Wend
End If
rs.Close
End Function

Function update()
'-------------------Validation Starts Here-----------------------------
If txttableno.Text = "" Then
    MsgBox "Enter the Table Number Properly...", vbInformation, "Sri Saravana Bhavan"
    txttableno.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    db.Execute "delete from tbl_tablemaster where tableid=" & Val(txttableid.Text)
    'db.Execute "delete from tbl_runorder where tableid=" & Val(txttableid.Text)
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_tablemaster", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields("tableid") = Val(Trim(txttableid.Text))
    rs.Fields("tableno") = UCase(Trim(txttableno.Text))
    rs.update
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_runorder where tableid=" & Val(txttableid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        
    Else
        rs.AddNew
    End If
    rs.Fields("tableid") = Val(Trim(txttableid.Text))
    rs.Fields("tableno") = UCase(Trim(txttableno.Text))
    rs.update
    rs.Close
    
    db.Execute "update tbl_order set tableno='" & UCase(Trim(txttableno.Text)) & "' where tableid=" & Val(Trim(txttableid.Text)) & " and iscomplete=false"
End If
End Function

Private Sub MSGrid_Click()
If MSGrid.Rows > 1 Then
    If rs2.State = 1 Then rs2.Close
    rs2.Open "select * from tbl_tablemaster where tableno='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs2.EOF Then
        txttableid.Text = rs2.Fields("tableid")
        txttableno.Text = rs2.Fields("tableno")
        txttableno.SetFocus
        txttableno.SelStart = 0
        txttableno.SelLength = Len(txttableno.Text)
        'txttableid.Enabled = False
    End If
    
    BtnSave.Enabled = False
    BtnModify.Enabled = True
    BtnDelete.Enabled = True
End If
End Sub

Private Sub txttableno_Change()
stmt = "select * from tbl_tablemaster where tableno like '" & Trim(txttableno.Text) & "%' order by tableno"
If rs3.State = 1 Then rs3.Close
rs3.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs3.EOF Then
    rs3.MoveFirst
    Do While Not rs3.EOF
        MSGrid.AddItem rs3.Fields("tableno")
        rs3.MoveNext
    Loop
End If
rs3.Close
End Sub

Private Sub txttableno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
