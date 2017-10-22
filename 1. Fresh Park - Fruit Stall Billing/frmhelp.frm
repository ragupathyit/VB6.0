VERSION 5.00
Begin VB.Form frmhelp 
   Caption         =   "Help"
   ClientHeight    =   9870
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   13500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   13500
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdClose 
      Caption         =   "CLOSE"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "3.   (< ) =  Previous Bill,     (>)  =  Next Bill"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   1920
      Width           =   11895
   End
   Begin VB.Label Label17 
      Caption         =   "2.   F2 = Sales Form 1,  F3 = Sales Form 2,  F4 = Sales Form 3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   1440
      Width           =   11895
   End
   Begin VB.Label Label16 
      Caption         =   "17. Alt + U = It opens the Calculator Window in the Main Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   8640
      Width           =   11895
   End
   Begin VB.Label Label15 
      Caption         =   "16. Alt + R = It opens the Report Window in the Main Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   8160
      Width           =   11895
   End
   Begin VB.Label Label14 
      Caption         =   "15. Alt + S = It opens the Sales Window in the Main Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   7680
      Width           =   11895
   End
   Begin VB.Label Label13 
      Caption         =   "14. Alt + I = It opens the Item Master Window in the Main Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   7200
      Width           =   11895
   End
   Begin VB.Label Label12 
      Caption         =   "13. Alt + E = Cursor moves to the Exit Button in the Sales Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   6720
      Width           =   11895
   End
   Begin VB.Label Label11 
      Caption         =   "12. Alt + L = Cursor moves to the Clear Button in the Sales Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   6240
      Width           =   11895
   End
   Begin VB.Label Label10 
      Caption         =   "11. Alt + B = Cursor moves to the Bill Button in the Sales Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   5760
      Width           =   11895
   End
   Begin VB.Shape Shape1 
      Height          =   9375
      Left            =   480
      Top             =   360
      Width           =   12135
   End
   Begin VB.Label Label9 
      Caption         =   "10. S (or) Alt + A = Cursor moves to the Save Button in the Sales Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   5280
      Width           =   11895
   End
   Begin VB.Label Label8 
      Caption         =   "9.   P (or) Alt + P = Cursor moves to the Print Button in the Sales Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4800
      Width           =   11895
   End
   Begin VB.Label Label7 
      Caption         =   "8.   C (or) Alt + C = Cursor moves to the Continue Button in the Sales Window"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   4320
      Width           =   11895
   End
   Begin VB.Label Label5 
      Caption         =   "7.   F11 = This key is allows you to change the current row item rate in the main list itself"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3840
      Width           =   11895
   End
   Begin VB.Label Label4 
      Caption         =   "6.   F8 = Move the cursor from the main list to the customer name textbox"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   11895
   End
   Begin VB.Label Label3 
      Caption         =   "5.   F7 = Move the cursor from the main list to the bill no textbox"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   11895
   End
   Begin VB.Label Label2 
      Caption         =   "4.   F6 = Delete the current row in the main list"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   11895
   End
   Begin VB.Label Label1 
      Caption         =   "1.   F1 = Cursor moves to the right side list for selecting the item and press enter to add in the main list"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   11895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   6120
      TabIndex        =   0
      Top             =   480
      Width           =   825
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   480
      Top             =   360
      Width           =   12135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   480
      Top             =   9120
      Width           =   12135
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
