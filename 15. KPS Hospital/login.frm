VERSION 5.00
Begin VB.Form LoginFrm 
   BackColor       =   &H00FF0000&
   Caption         =   "Login"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
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
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtpwd 
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
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtuname 
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
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1605
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   645
      Width           =   1530
   End
End
Attribute VB_Name = "LoginFrm"
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
    If txtuname.Text = "bill" Then
        MDIForm1.mnudoctor.Enabled = False
        MDIForm1.mnufee.Enabled = False
        MDIForm1.mnubackup.Enabled = False
    End If
    
    Unload Me
    MDIForm1.Show
Else
    MsgBox "Username and Password Incorrect!", vbInformation, "RMV Medical"
End If
End Sub

Private Sub Form_Load()
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

rs.Open "select * from tbl_login", db, adOpenDynamic, adLockOptimistic

'strs = "00-1C-C0-31-20-2C"  'Prabhu Shop Machine mac address
'strs = "00-1C-C0-31-20-2C"
'If Not (GetMACs_AdaptInfo() = strs) Then
'    MsgBox "Please contact the software vendor to use this software...", vbInformation, "RMV Medical"
'    End
'End If

'---------------Software Maintanence-------------------------------------
'If Date = #12/20/2014# Then
'    MsgBox "This software needs maintenance. It will work for next 3 days only. Please contact the software vendor", vbInformation, "RMV Medical"
'End If
'If Date = #12/21/2014# Then
'    MsgBox "This software needs maintenance. It will work for next 2 days only. Please contact the software vendor", vbInformation, "RMV Medical"
'End If
'If Date = #12/22/2014# Then
'    MsgBox "This software needs maintenance. It will work for next 1 days only. Please contact the software vendor", vbInformation, "RMV Medical"
'End If
'If Date = #12/23/2014# Then
'    MsgBox "Please contact the software vendor. This software is going to out of date from tomarrow onwords", vbInformation, "RMV Medical"
'End If
'If Date > #12/23/2014# Then
'    MsgBox "Please contact the software vendor to use this software...", vbInformation, "RMV Medical"
'    End
'End If
'---------------Software Maintanence-------------------------------------
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
