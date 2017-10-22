VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFee 
   BackColor       =   &H00FF0000&
   Caption         =   "Fees Details"
   ClientHeight    =   7290
   ClientLeft      =   225
   ClientTop       =   450
   ClientWidth     =   15135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   74766.22
   ScaleMode       =   0  'User
   ScaleWidth      =   60511.92
   WindowState     =   2  'Maximized
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
      Left            =   480
      TabIndex        =   2
      Top             =   6240
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
      Left            =   2280
      TabIndex        =   10
      Top             =   6240
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
      Left            =   6000
      TabIndex        =   9
      Top             =   6240
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
      Left            =   7920
      TabIndex        =   8
      Top             =   6240
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
      Left            =   4080
      TabIndex        =   7
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtfcharges 
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
      Left            =   4680
      TabIndex        =   1
      Top             =   4080
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   10215
      Left            =   9840
      TabIndex        =   5
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   18018
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16711680
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "F Code |Fees Name                                        |Charges    "
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
   Begin VB.TextBox txtfname 
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
      Left            =   4680
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblhide 
      Caption         =   "Label2"
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "FEES DETAILS"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   360
      Width           =   3675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Name"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Charges"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   4200
      Width           =   2100
   End
End
Attribute VB_Name = "FrmFee"
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
db.Execute "delete from tbl_feesmaster where fcode=" & Val(lblhide.Caption)
MsgBox "Successfully Deleted...", vbInformation, "KPS Hospital"
Call cmdclear_Click
End Sub

Private Sub CmdModify_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_feesmaster where fcode=" & Val(lblhide.Caption), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.Fields("feename") = UCase(txtfname.Text)
    rs.Fields("charges") = Val(txtfcharges.Text)
    rs.Update
End If
rs.Close
MsgBox "Successfully Modified", vbInformation, "KPS Hospital"
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtfname.Text = "" Then
    MsgBox "Enter the Fees Name Properly...", vbInformation, "KPS Hospital"
    txtfname.SetFocus
ElseIf txtfcharges.Text = "" Then
    MsgBox "Enter the Fees Charges Properly...", vbInformation, "KPS Hospital"
    txtfcharges.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_feesmaster order by fcode", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveLast
        s = Val(rs.Fields("fcode")) + 1
    Else
        s = 1
    End If
    
    rs.AddNew
    rs.Fields("fcode") = Val(s)
    rs.Fields("feename") = UCase(txtfname.Text)
    rs.Fields("charges") = Val(txtfcharges.Text)
    rs.Update
    MsgBox "Fees Details Saved Successfully", vbInformation, "KPS Hospital"
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
    rs.Open "Select * from tbl_feesmaster where fcode=" & Val(MSGrid.TextMatrix(MSGrid.Row, 0)), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txtfname.Text = rs.Fields("feename")
        txtfcharges.Text = rs.Fields("charges")
    End If

    CmdSave.Enabled = False
    CmdModify.Enabled = True
    cmdclear.Enabled = True
    CmdDelete.Enabled = True
    txtfname.SetFocus
End Sub

Private Function Fill()
stmt = "select * from tbl_feesmaster order by fcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("fcode") & vbTab & rs.Fields("feename") & vbTab & rs.Fields("charges")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub txtfname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfcharges.SetFocus
End If
End Sub

Private Sub txtfcharges_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub
