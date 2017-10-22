VERSION 5.00
Begin VB.Form LoginFrm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "  Login Form"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0E0FF&
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
      ToolTipText     =   "Click to exit or cancel"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00C0E0FF&
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
      ToolTipText     =   "Click to login"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtpwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtuname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   3240
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2325
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1605
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   6735
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
    MsgBox "Username and Password Incorrect!", vbInformation, "Rotary Club of Mettupalayam"
End If
End Sub

Private Sub Form_Load()

'Me.BackColor = RGB(255, 204, 203)

db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

rs.Open "select * from tbl_login", db, adOpenDynamic, adLockOptimistic

'strs = "00-1F-16-BE-02-84" 'Ragupathy mac address
'strs = "00-1C-C0-31-20-2C" ' Sri Rotary Club of Mettupalayam Machine 1 Mac Address
'strs = "00-1E-90-E7-7D-EC" ' Sri Rotary Club of Mettupalayam Machine 2 Mac Address
'If Not GetMACs_AdaptInfo() = strs Then
    'MsgBox "Please contact the software vendor to use this software...", vbInformation, "Rotary Club of Mettupalayam"
    'End
'End If

'---------------Software Maintanence-------------------------------------
If Date = #12/28/2017# Then
    MsgBox "This software needs maintenance. It will work for next 3 days only. Please contact the software vendor", vbInformation, "Rotary Club of Mettupalayam"
End If
If Date = #12/29/2017# Then
    MsgBox "This software needs maintenance. It will work for next 2 days only. Please contact the software vendor", vbInformation, "Rotary Club of Mettupalayam"
End If
If Date = #12/30/2017# Then
    MsgBox "This software needs maintenance. It will work for next 1 days only. Please contact the software vendor", vbInformation, "Rotary Club of Mettupalayam"
End If
If Date = #12/31/2017# Then
    MsgBox "Please contact the software vendor. This software is going to out of date from tomarrow onwords", vbInformation, "Rotary Club of Mettupalayam"
End If
If Date > #12/31/2017# Then
    MsgBox "Please contact the software vendor to use this software...", vbInformation, "Rotary Club of Mettupalayam"
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
