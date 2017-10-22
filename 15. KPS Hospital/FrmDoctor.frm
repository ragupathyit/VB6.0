VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDoctor 
   BackColor       =   &H00FF0000&
   Caption         =   "Doctor Details"
   ClientHeight    =   7290
   ClientLeft      =   225
   ClientTop       =   450
   ClientWidth     =   15135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   49775.87
   ScaleMode       =   0  'User
   ScaleWidth      =   45227.06
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdesig 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   13
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtquali 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "S&AVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "&MODIFY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   8
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "C&LEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtdname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   10095
      Left            =   9360
      TabIndex        =   3
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   17806
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16711680
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "D Code |Doctor Name                                        |Qualification              "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
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
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2160
      TabIndex        =   12
      Top             =   4800
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2160
      TabIndex        =   11
      Top             =   3600
      Width           =   1875
   End
   Begin VB.Label lblhide 
      Caption         =   "Label2"
      Height          =   495
      Left            =   10440
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   4545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   2280
      Width           =   1965
   End
End
Attribute VB_Name = "FrmDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclear_Click()
Unload Me
Load Me
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
db.Execute "delete from tbl_doctormaster where dcode=" & Val(lblhide.Caption)
MsgBox "Successfully Deleted...", vbInformation, "KPS Hospital"
Call cmdclear_Click
End Sub

Private Sub CmdModify_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_doctormaster where dcode=" & Val(lblhide.Caption), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.Fields("doctorname") = UCase(txtdname.Text)
    rs.Fields("quali") = UCase(txtquali.Text)
    rs.Fields("desig") = UCase(txtdesig.Text)
    rs.Update
End If
rs.Close
MsgBox "Successfully Modified", vbInformation, "KPS Hospital"
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtdname.Text = "" Then
    MsgBox "Enter the Doctor Name Properly...", vbInformation, "KPS Hospital"
    txtdname.SetFocus
'ElseIf txtquali.Text = "" Then
'    MsgBox "Enter the Qualification Properly...", vbInformation, "KPS Hospital"
'    txtquali.SetFocus
'ElseIf txtdesig.Text = "" Then
'    MsgBox "Enter the Designation Properly...", vbInformation, "KPS Hospital"
'    txtdesig.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_doctormaster order by dcode", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveLast
        s = Val(rs.Fields("dcode")) + 1
    Else
        s = 1
    End If
    
    rs.AddNew
    rs.Fields("dcode") = Val(s)
    rs.Fields("doctorname") = UCase(txtdname.Text)
    rs.Fields("quali") = UCase(txtquali.Text)
    rs.Fields("desig") = UCase(txtdesig.Text)
    rs.Update

    MsgBox "Doctor Details Saved Successfully", vbInformation, "KPS Hospital"
    rs.Close

    Call cmdclear_Click
End If
End Sub

Private Sub Form_Load()
Call connect
Call Fill

CmdSave.Enabled = True
CmdModify.Enabled = False
cmdclear.Enabled = True
CmdDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
    lblhide.Caption = MSGrid.TextMatrix(MSGrid.Row, 0)
    If rs.State = 1 Then rs.Close
    rs.Open "Select * from tbl_doctormaster where dcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 0)), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txtdname.Text = rs.Fields("doctorname")
        txtquali.Text = rs.Fields("quali")
        txtdesig.Text = rs.Fields("desig")
    End If

    CmdSave.Enabled = False
    CmdModify.Enabled = True
    cmdclear.Enabled = True
    CmdDelete.Enabled = True
    txtdname.SetFocus
End Sub

Private Function Fill()
stmt = "select * from tbl_doctormaster order by dcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("dcode") & vbTab & rs.Fields("doctorname") & vbTab & rs.Fields("quali")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub txtdname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtquali.SetFocus
End If
End Sub

Private Sub txtquali_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdesig.SetFocus
End If
End Sub

Private Sub txtdesig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub
