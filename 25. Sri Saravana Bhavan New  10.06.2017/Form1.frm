VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Since 1987"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   2100
      Left            =   4080
      Picture         =   "Form1.frx":0000
      Top             =   600
      Width           =   1890
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sri Saravana Bhavan"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KOT Billing System 1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9944537774,9865555042"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Adonis Technologies"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "Form1.frx":B50A
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Interval = 15
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 2
If ProgressBar1.Value >= ProgressBar1.Max Then
    Timer1.Enabled = False
    Unload Me
    FrmLogin.Show
End If
End Sub
