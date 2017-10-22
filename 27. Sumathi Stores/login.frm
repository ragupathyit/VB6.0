VERSION 5.00
Begin VB.Form LoginFrm1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "  Login Form"
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   6420
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Proj_1.Button cmdlogin 
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":DBF5
      PICN            =   "login.frx":DC11
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox txtuname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   3240
      Width           =   3015
   End
   Begin Proj_1.Button cmdcancel 
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":E323
      PICN            =   "login.frx":E33F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "LoginFrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim flag As Integer

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub cmdlogin_Click()
flag = 0
rs.MoveFirst
If Not rs.EOF Then
      Do While Not rs.EOF
        If (txtuname.Text = rs.Fields(0)) And (txtpwd.Text = rs.Fields(1)) Then
            flag = 1
        End If
        rs.MoveNext
    Loop
End If
If flag = 1 Then
    'MsgBox "Your Login is valid!", vbInformation
    Unload Me
    MDIForm1.Show
Else
    MsgBox "Username and Password Incorrect!", vbInformation, "Sumathi Stores"
End If
End Sub

Private Sub Form_Load()

db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

rs.Open "select * from tbl_login", db, adOpenDynamic, adLockOptimistic

'strs = "00-1F-16-BE-02-84" 'Ragupathy mac address

'strs = "50-E5-49-90-3B-71" ' Sumathi StoresMac Address
'If Not GetMACs_AdaptInfo() = strs Then
'    MsgBox "Please contact the software vendor Prabhu to use this software...", vbInformation, "Sumathi Stores"
'    End
'End If

'---------------Software Maintanence-------------------------------------
'If Date = #3/28/2016# Then
'    MsgBox "This software needs maintenance. It will work for next 3 days only. Please contact the software vendor", vbInformation, "Sumathi Stores"
'End If
'If Date = #3/29/2016# Then
'    MsgBox "This software needs maintenance. It will work for next 2 days only. Please contact the software vendor", vbInformation, "Sumathi Stores"
'End If
If Date = #4/28/2017# Then
    MsgBox "This software needs maintenance. It will work for next 1 days only. Please contact the software vendor", vbInformation, "Sumathi Stores"
End If
If Date = #4/29/2017# Then
    MsgBox "Please contact the software vendor. This software is going to out of date from tomarrow onwords", vbInformation, "Sumathi Stores"
End If
If Date > #4/30/2017# Then
    MsgBox "Please contact the software vendor to use this software...", vbInformation, "Sumathi Stores"
    End
End If

End Sub

Private Sub txtpwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdlogin_Click
End If
End Sub

Private Sub txtuname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpwd.SetFocus
End If
End Sub
