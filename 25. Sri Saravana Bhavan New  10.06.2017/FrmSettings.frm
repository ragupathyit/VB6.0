VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSettings 
   BackColor       =   &H00400040&
   Caption         =   "Settings"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   12750
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
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
      Height          =   2460
      ItemData        =   "FrmSettings.frx":0000
      Left            =   3240
      List            =   "FrmSettings.frx":0016
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtcnpwd 
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
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox txtnpwd 
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
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox txtcpwd 
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
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   11520
      TabIndex        =   0
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
      MICON           =   "FrmSettings.frx":0066
      PICN            =   "FrmSettings.frx":0082
      PICH            =   "FrmSettings.frx":0794
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
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   "SAVE"
      Top             =   6720
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
      MICON           =   "FrmSettings.frx":0EA6
      PICN            =   "FrmSettings.frx":0EC2
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
      Left            =   3960
      TabIndex        =   11
      ToolTipText     =   "CLEAR"
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Refresh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      MICON           =   "FrmSettings.frx":15D4
      PICN            =   "FrmSettings.frx":15F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnSBackup 
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      ToolTipText     =   "BACKUP"
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "S&oftware Backup        "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      MICON           =   "FrmSettings.frx":1D02
      PICN            =   "FrmSettings.frx":1D1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnOptimize 
      Height          =   495
      Left            =   8280
      TabIndex        =   16
      ToolTipText     =   "BACKUP"
      Top             =   1320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "O&ptimize Today's Bill "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      MICON           =   "FrmSettings.frx":2430
      PICN            =   "FrmSettings.frx":244C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAST UPDATED DATE: 14/06/2017"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   5160
      TabIndex        =   15
      Top             =   240
      Width           =   5700
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accessing Forms *"
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
      Left            =   720
      TabIndex        =   13
      Top             =   3720
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   7320
      X2              =   7320
      Y1              =   1080
      Y2              =   6120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   6480
      Width           =   12135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SETTINGS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password *"
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
      Left            =   720
      TabIndex        =   8
      Top             =   3120
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password  *"
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
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   1830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password "
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
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1995
   End
   Begin MSForms.ComboBox cmbusername 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5530;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name *"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmSettings.frx":2B5E
      Top             =   240
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   12135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   7455
      Left            =   0
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClear_Click()
Unload Me
FrmSettings.Show
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub BtnOptimize_Click()
If rs.State = 1 Then rs.Close
rs.Open "select distinct orderid from tbl_order where orderdate=#" & Format(Date, "mm/dd/yyyy") & "# order by orderid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    bill = 1
    db.Execute "Delete from tbl_tempbill"
    While Not rs.EOF
        db.Execute "update tbl_order set billno=" & Trim(bill) & " where orderid=" & Val(rs.Fields("orderid"))
        db.Execute "insert into tbl_tempbill values ('" & Format(Date, "mm/dd/yyyy") & "'," & bill & ")"
        bill = Val(bill) + 1
        rs.MoveNext
    Wend
End If
MsgBox "Successfully Optimized for Today's Sales", vbInformation
End Sub

Private Sub BtnSave_Click()
If cmbusername.Text = "" Then
    MsgBox "Select the User Name Properly...", vbInformation, "Sri Saravana Bhavan"
    cmbusername.SetFocus
ElseIf txtnpwd.Text = "" Then
    MsgBox "Enter the New Password Properly...", vbInformation, "Sri Saravana Bhavan"
    txtnpwd.SetFocus
ElseIf txtcnpwd.Text = "" Then
    MsgBox "Enter the Confirm Password Properly...", vbInformation, "Sri Saravana Bhavan"
    txtcnpwd.SetFocus
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_login where username='" & Trim(cmbusername.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        'rs.Fields("password") = Trim(txtnpwd.Text)
        'rs.update
        
        db.Execute "delete from tbl_login where username='" & Trim(cmbusername.Text) & "'"
        
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) = True Then
                rs.AddNew
                rs.Fields("username") = Trim(cmbusername.Text)
                rs.Fields("password") = Trim(txtnpwd.Text)
                rs.Fields("forms") = List1.List(i)
                rs.update
            End If
        Next i
    Else
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) = True Then
                rs.AddNew
                rs.Fields("username") = Trim(cmbusername.Text)
                rs.Fields("password") = Trim(txtnpwd.Text)
                rs.Fields("forms") = List1.List(i)
                rs.update
            End If
        Next i
    End If
    MsgBox "Password is Successfully Updated", vbInformation, "Sri Saravana Bhavan"
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSBackup_Click()
a = MsgBox("Are you sure to take backup and start from billno 1", vbYesNo)
If a = vbYes Then
'    Dim fso As Object
'    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
'    Call fso.CopyFile(App.Path & "\master.mdb", App.Path & "\master" & Year(Now) & ".mdb")
'
'    If Now = "03/31/" & Year(Now) Then
'        db.Execute "delete from tbl_order where orderdate<=#03/31/" & Year(Now) & "#"
'    End If
    
    MsgBox "Software Backup is Created Successfully", vbInformation, "Sri Saravana Bhavan"
    BtnSBackup.Enabled = False
End If
End Sub

Private Sub cmbusername_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_login where username='" & Trim(cmbusername.Text) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtcpwd.Text = rs.Fields("password")
    While Not rs.EOF
        For i = 0 To List1.ListCount - 1
            If List1.List(i) = rs.Fields("forms") Then
                List1.Selected(i) = True
            End If
        Next i
        rs.MoveNext
    Wend
End If
rs.Close

txtnpwd.SetFocus
txtnpwd.SelStart = 0
txtnpwd.SelLength = Len(txtnpwd.Text)   'select the text
End Sub

Private Sub Form_Load()
Call connect
cmbusername.Clear
If rs.State = 1 Then rs.Close
rs.Open "select distinct username from tbl_login order by username", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    cmbusername.AddItem rs.Fields("username")
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub txtnpwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcnpwd.SetFocus
    txtcnpwd.SelStart = 0
    txtcnpwd.SelLength = Len(txtcnpwd.Text)
End If
End Sub

Private Sub txtcnpwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(txtnpwd.Text) <> Trim(txtcnpwd.Text) Then
        MsgBox "Password is not match, give the correct password!", vbInformation, "Sri Saravana Bhavan"
        txtnpwd.Text = ""
        txtcnpwd.Text = ""
        txtnpwd.SetFocus
    Else
        BtnSave.SetFocus
    End If
End If
End Sub
