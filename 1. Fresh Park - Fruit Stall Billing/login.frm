VERSION 5.00
Begin VB.Form LoginFrm 
   Caption         =   "Login Form"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtuname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   2325
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Enter User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1605
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   3120
      Width           =   6735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "LOGIN FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1965
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   6735
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
    'MsgBox "Your Login is valid!", vbInformation
    Unload Me
    MDIForm1.Show
Else
    MsgBox "Username and Password Incorrect!", vbInformation, "Fresh Park"
End If
End Sub

Private Sub Form_Load()

db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

rs.Open "select * from tbl_login", db, adOpenDynamic, adLockOptimistic

'strs = "00-1F-16-BE-02-84" 'Ragu Lap acer
'strs = "00-80-48-63-23-62" 'Fresh Park mac address
'If Not GetMACs_AdaptInfo() = strs Then
    'MsgBox "Please contact the software vendor to use this software...", vbInformation, "Fresh Park"
   ' End
'End If

'---------------Software Maintanence-------------------------------------
'If Date = #3/28/2012# Then
 '   MsgBox "This software needs maintenance. It will work for next 3 days only. Please contact the software vendor", vbInformation, "Fresh Park"
'End If
'If Date = #3/29/2012# Then
'    MsgBox "This software needs maintenance. It will work for next 2 days only. Please contact the software vendor", vbInformation, "Fresh Park"
'End If
'If Date = #3/30/2012# Then
 '   MsgBox "This software needs maintenance. It will work for next 1 days only. Please contact the software vendor", vbInformation, "Fresh Park"
'End If
'If Date = #3/31/2012# Then
 '   MsgBox "Please contact the software vendor. This software is going to out of date from tomarrow onwords", vbInformation, "Fresh Park"
'End If
'If Date = #4/1/2012# Then
 '   MsgBox "Please contact the software vendor to use this software...", vbInformation, "Fresh Park"
  '  End
'End If

End Sub

'Private Sub txtpwd_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
 '   Call cmdlogin_Click
'End If
'End Sub

'Private Sub txtuname_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
 '   txtpwd.SetFocus
'End If
'End Sub
