VERSION 5.00
Begin VB.Form LoginFrm1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "  Login Form"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4465
      Width           =   3495
   End
   Begin VB.TextBox txtuname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3840
      TabIndex        =   0
      Top             =   3650
      Width           =   3495
   End
   Begin Proj_1.Button cmdlogin 
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   5280
      Width           =   4095
      _ExtentX        =   7223
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
      BCOL            =   16576
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":CEE5
      PICN            =   "login.frx":CF01
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
      Height          =   9000
      Left            =   0
      Top             =   0
      Width           =   10500
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
    MsgBox "Username and Password Incorrect!", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
End If
End Sub

Private Sub Form_Load()

db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

rs.Open "select * from tbl_login", db, adOpenDynamic, adLockOptimistic

'strs = "00-1F-16-BE-02-84" 'Ragupathy mac address
'strs = "50-E5-49-90-3B-71" ' Sri Annamalaiyar Pattasu KadaiMac Address
'If Not GetMACs_AdaptInfo() = strs Then
    'MsgBox "Please contact the software vendor to use this software...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    'End
'End If

'---------------Software Maintanence-------------------------------------
If Date = #11/22/2017# Then
    MsgBox "This software needs maintenance. It will work for next 3 days only. Please contact the software vendor", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
End If
If Date = #11/23/2017# Then
    MsgBox "This software needs maintenance. It will work for next 2 days only. Please contact the software vendor", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
End If
If Date = #11/24/2017# Then
    MsgBox "This software needs maintenance. It will work for next 1 days only. Please contact the software vendor", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
End If
If Date = #11/25/2017# Then
    MsgBox "Please contact the software vendor. This software is going to out of date from tomarrow onwords", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
End If
If Date > #11/25/2017# Then
    MsgBox "Please contact the software vendor to use this software...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
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
