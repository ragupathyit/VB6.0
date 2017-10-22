VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPaperMaster 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   7695
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P. ID"
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
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Name"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Rate"
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
         Left            =   6240
         TabIndex        =   15
         Top             =   0
         Width           =   1185
      End
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   8040
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
      MICON           =   "FrmPaperMaster.frx":0000
      PICN            =   "FrmPaperMaster.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FrameForm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   8055
      Begin VB.TextBox txtpaperrate 
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtpapername 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox txtpid 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Rate *"
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
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Name *"
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
         Left            =   0
         TabIndex        =   6
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper ID"
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
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   930
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   3975
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "         .|                                                                                         .|                         ."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tamil-Ananthi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Button BtnSave 
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      ToolTipText     =   "SAVE"
      Top             =   6600
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
      MICON           =   "FrmPaperMaster.frx":072E
      PICN            =   "FrmPaperMaster.frx":074A
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
      Left            =   2880
      TabIndex        =   11
      ToolTipText     =   "MODIFY"
      Top             =   6600
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
      MICON           =   "FrmPaperMaster.frx":0E5C
      PICN            =   "FrmPaperMaster.frx":0E78
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
      Left            =   6480
      TabIndex        =   12
      ToolTipText     =   "DELETE"
      Top             =   6600
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
      MICON           =   "FrmPaperMaster.frx":158A
      PICN            =   "FrmPaperMaster.frx":15A6
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
      Left            =   4680
      TabIndex        =   13
      ToolTipText     =   "CLEAR"
      Top             =   6600
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
      MICON           =   "FrmPaperMaster.frx":1CB8
      PICN            =   "FrmPaperMaster.frx":1CD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmPaperMaster.frx":23E6
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAPER MASTER"
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
      TabIndex        =   3
      Top             =   240
      Width           =   2670
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   6480
      Width           =   9255
   End
End
Attribute VB_Name = "FrmPaperMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_papermaster where pid=" & Val(txtpid.Text)
MsgBox "Paper Deleted Successfully", vbInformation, "Press Management"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmPaperMaster.Show
End Sub

Private Sub BtnModify_Click()
'-------------------Validation Starts Here-----------------------------
If txtpapername.Text = "" Then
    MsgBox "Enter the Paper Name Properly...", vbInformation, "Press Management"
    txtitemname.SetFocus
ElseIf Not IsNumeric(Val(txtpaperrate.Text)) Then
    MsgBox "Enter the Paper Rate Properly...", vbInformation, "Press Management"
    txtpaperrate.Text = ""
    txtpaperrate.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_papermaster where pid=" & Val(Trim(txtpid.Text)), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.Fields("pid") = Val(Trim(txtpid.Text))
        rs.Fields("papername") = Trim(txtpapername.Text)
        rs.Fields("prate") = Val(Trim(txtpaperrate.Text))
        rs.Update
    End If
    rs.Close
    MsgBox "Paper Modified Successfully", vbInformation, "Press Management"

    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtpapername.Text = "" Then
    MsgBox "Enter the Paper Name Properly...", vbInformation, "Press Management"
    txtpapername.SetFocus
ElseIf Not IsNumeric(Val(txtpaperrate.Text)) Then
    MsgBox "Enter the Paper Rate Properly...", vbInformation, "Press Management"
    txtpaperrate.Text = ""
    txtpaperrate.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_papermaster", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields("pid") = Val(Trim(txtpid.Text))
    rs.Fields("papername") = Trim(txtpapername.Text)
    rs.Fields("prate") = Val(Trim(txtpaperrate.Text))
    rs.Update
    rs.Close
    MsgBox "Paper Saved Successfully", vbInformation, "Press Management"

    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
stmt = "select * from tbl_papermaster order by pid"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("pid") & vbTab & rs.Fields("papername") & vbTab & Format(rs.Fields("prate"), "0.00")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select pid from tbl_papermaster order by pid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtpid.Text = rs.Fields("pid") + 1
Else
    txtpid.Text = 1
End If
rs.Close

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_papermaster where pid=" & Val(Trim(MSGrid.TextMatrix(MSGrid.Row, 0))), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtpid.Text = rs1.Fields("pid")
    txtpapername.Text = rs1.Fields("papername")
    txtpaperrate.Text = Format(rs1.Fields("prate"), "0.00")
End If

txtpapername.SetFocus
BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtpapername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpaperrate.SetFocus
    txtpaperrate.SelStart = 0
    txtpaperrate.SelLength = Len(txtpaperrate.Text)    'select the text
End If
End Sub

Private Sub txtpaperrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
