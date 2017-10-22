VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Fresh Park Billing System"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7950
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuitemmaster 
      Caption         =   "&Item Master"
   End
   Begin VB.Menu mnusales 
      Caption         =   "&Sales"
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnusalesreport 
         Caption         =   "Sales Day Wise"
      End
      Begin VB.Menu mnusalesperiodreport 
         Caption         =   "Sales Period Wise"
      End
   End
   Begin VB.Menu mnucalc 
      Caption         =   "Calc&ulator"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub mnucalc_Click()
Call Shell("calc.exe", vbNormalFocus)
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuhelp_Click()
Unload Me
frmhelp.Show
End Sub

Private Sub mnuitemmaster_Click()
Unload Me
ItemFrm.Show
End Sub

Private Sub mnusales_Click()
SalesFrm1.Show
End Sub

Private Sub mnusalesperiodreport_Click()
Unload Me
SalesPeriodRpt.Show
End Sub

Private Sub mnusalesreport_Click()
Unload Me
SalesRpt.Show
End Sub
