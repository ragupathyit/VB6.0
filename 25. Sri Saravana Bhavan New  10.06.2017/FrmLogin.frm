VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtpwd 
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtuname 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin Project1.Button Button1 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "LOGIN"
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
      MICON           =   "FrmLogin.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button Button2 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "CLOSE"
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
      MICON           =   "FrmLogin.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1620
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   3120
      Width           =   5895
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim flag As Integer

Private Sub Button1_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_login where username='" & Trim(txtuname.Text) & "' and password='" & Trim(txtpwd.Text) & "'", db, adOpenDynamic, adLockOptimistic

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_waitermaster where username='" & Trim(txtuname.Text) & "' and password='" & Trim(txtpwd.Text) & "'", db, adOpenDynamic, adLockOptimistic

flag = 0
If Not rs.EOF Then
    uname = Trim(txtuname.Text)
    flag = 1
ElseIf Not rs1.EOF Then
    uname = Trim(txtuname.Text)
    flag = 1
End If
If flag = 1 Then
    'MsgBox "Your Login is valid!", vbInformation
    Unload Me
    MDIForm1.Show
Else
    MsgBox "Username and Password Incorrect!", vbInformation, "Sri Saravana Bhavan"
End If
End Sub

Private Sub Button2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

'strs = "D8-50-E6-F1-90-FF" 'Ragupathy mac address
'strs = "C8-3A-35-D1-7F-BD" ' Sri Saravana Bhavan Mac Address
'If Not GetMACs_AdaptInfo() = strs Then
'    MsgBox "Please contact the software vendor to use this software...", vbInformation, "Sri Saravana Bhavan"
'    End
'End If

'---------------Software Maintanence-------------------------------------
'If Date = #5/26/2016# Then
'    MsgBox "This software needs maintenance. It will work for next 4 days only. Please contact the software vendor", vbInformation, "Sri Saravana Bhavan"
'End If
'If Date = #5/27/2016# Then
'    MsgBox "This software needs maintenance. It will work for next 3 days only. Please contact the software vendor", vbInformation, "Sri Saravana Bhavan"
'End If
'If Date = #5/28/2016# Then
'    MsgBox "This software needs maintenance. It will work for next 2 days only. Please contact the software vendor", vbInformation, "Sri Saravana Bhavan"
'End If
'If Date = #5/29/2016# Then
'    MsgBox "This software needs maintenance. It will work for next 1 days only. Please contact the software vendor", vbInformation, "Sri Saravana Bhavan"
'End If
'If Date = #5/30/2016# Then
'    MsgBox "Please contact the software vendor. This software is going to out of date from tomarrow onwords", vbInformation, "Sri Saravana Bhavan"
'End If
'If Date > #5/30/2016# Then
'    MsgBox "Please contact the software vendor to use this software...", vbInformation, "Sri Saravana Bhavan"
'    End
'End If

End Sub

Private Sub txtpwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Button1_Click
End If
End Sub

Private Sub txtuname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpwd.SetFocus
End If
End Sub
