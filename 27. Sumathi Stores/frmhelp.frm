VERSION 5.00
Begin VB.Form frmhelp 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   8505
   ClientLeft      =   75
   ClientTop       =   60
   ClientWidth     =   13470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2.   F3 = It Opens Retail Sales Form and F4 = It Opens Whole Sales Form"
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
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   11895
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8.   D = This key is used to change the cursor from continue button to discount textbox"
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
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   11895
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7.   A = This key is used to change the cursor from continue button to save button"
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
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   11895
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5.   F9 = Move the cursor from the main list to the date picker"
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
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   11895
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "19. Alt + X = It closes the Application"
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
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   11895
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "18. Alt + H = It opens the Help Window in the Main Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   11895
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "17. Alt + U = It opens the Calculator Window in the Main Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6600
      Width           =   11895
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "16. Alt + R = It opens the Report Window in the Main Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   11895
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "15. Alt + S = It opens the Sales Window in the Main Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   11895
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "14. Alt + I = It opens the Item Master Window in the Main Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   11895
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "13. Alt + E = Cursor moves to the Exit Button in the Sales Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   11895
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12. Alt + L = Cursor moves to the Clear Button in the Sales Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   11895
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "11. Alt + B = Cursor moves to the Bill Button in the Sales Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   11895
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10. Alt + A = Cursor moves to the Save Button in the Sales Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   11895
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9.   Alt + C = Cursor moves to the Continue Button in the Sales Window"
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   11895
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6.   F11 = This key is allows you to change the current row item rate in the main list itself"
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   11895
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4.   F8 = Move the cursor from the main list to the customer name textbox"
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   11895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3.   F6 = Delete the current row in the main list"
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1.   F1 = Cursor moves to the right side list for selecting the item and press enter to add in the main list"
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   13335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(35, 29, 29)
End Sub
