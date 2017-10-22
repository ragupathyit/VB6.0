VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FF0000&
   Caption         =   "KSA and Sons"
   ClientHeight    =   7260
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14685
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   14625
      TabIndex        =   0
      Top             =   0
      Width           =   14685
   End
   Begin VB.Menu mnumaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuitemmaster 
         Caption         =   "Item Master"
      End
      Begin VB.Menu mnusupplier 
         Caption         =   "Supplier Master"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucustmaster 
         Caption         =   "Customer Master"
      End
   End
   Begin VB.Menu mnudaybook 
      Caption         =   "&Day Book"
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnustockreport 
         Caption         =   "Current Stock"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnupurperiodreport 
         Caption         =   "Purchase Period Wise"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusalesperiodreport 
         Caption         =   "Sales Period Wise"
      End
      Begin VB.Menu mnusalesstatement 
         Caption         =   "Sales Statement"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudbreport 
         Caption         =   "Daybook Report"
      End
   End
   Begin VB.Menu mnucalculator 
      Caption         =   "Calc&ulator"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

'===============Removing Border and Title Bar===============================
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
 
Const WS_CAPTION = &HC00000
Const WS_SYSMENU = &H80000
'Const WS_MINIMIZEBOX = &H20000
'Const WS_MAXIMIZEBOX = &H10000
Const WS_THICKFRAME = &H40000
Const GWL_STYLE = (-16)
'===============Removing Border and Title Bar===============================

Private Sub MDIForm_Load()
    Me.BackColor = RGB(84, 96, 254)
    Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
    'L = L And Not (WS_MINIMIZEBOX)
    'L = L And Not (WS_MAXIMIZEBOX)
    L = L And Not (WS_THICKFRAME)
    L = L Xor WS_CAPTION
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)
End Sub

Private Sub mnucalculator_Click()
Call Shell("calc.exe", vbNormalFocus)
End Sub

Private Sub mnucustmaster_Click()
CustomerFrm.Show
End Sub

Private Sub mnudaybook_Click()
DBdaybookfrm.Show
End Sub

Private Sub mnudbreport_Click()
daybookRpt.Show
End Sub

Private Sub mnuhelp_Click()
frmhelp.Show
End Sub

Private Sub mnuitemmaster_Click()
AdminPassword.Show
End Sub

Private Sub mnusalesperiodreport_Click()
SalesPeriodRpt.Show
End Sub

Private Sub mnusalesstatement_Click()
SalesStatementRpt.Show
End Sub

Private Sub mnusupplier_Click()
SupplierFrm.Show
End Sub

Private Sub mnuitemsales1_Click()
SalesFrm.Show
End Sub

Private Sub mnupurperiodreport_Click()
PurchasePeriodRpt.Show
End Sub

Private Sub mnupursupwise_Click()
PurchasePeriodSupRpt.Show
End Sub

Private Sub mnustockreport_Click()
If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"

stmt = "select * from tbl_stock"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    '----------Notepad print------------------
    Open App.Path & "\rptcurrentstock.txt" For Output As #1
    
    Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
    Print #1, ""
    Print #1, Space(17) & "Sumathi Stores"
    Print #1, Space(13) & "New Vegitable Market"
    Print #1, Space(9) & "Vegitable Supplier and Commission Agency"
    Print #1, Space(13) & "METTUPALAYAM - 641301"
    Print #1, Space(6) & "CELL NO - 93641 33333, 90034 00000"
    Print #1, ""
    Print #1, "Current Stock as on " & Format(Date, "DD/MM/YYYY") & Space(15) & "Time: " & Time()
    Print #1, "--------------------------------------------------------------"           ' 62
    Print #1, "Item Code" & Space(2) & "Item Name " & Space(31) & Space(2) & "Quantity"
    Print #1, "--------------------------------------------------------------"
    tqty = 0
    X = 10
    While Not rs.EOF
        If Val(rs.Fields("qty")) <> 0 Then
            tqty = Val(tqty) + Val(rs.Fields("qty"))
        
            icode = 9 - Len(rs.Fields("itemcode"))
            iname = 41 - Len(Mid(rs.Fields("itemname"), 1, 41))
            iqty = 8 - Len(Format(rs.Fields("qty"), "0.00"))
        
            Print #1, UCase(rs.Fields("itemcode")) & Space(icode) & Space(2) & UCase(Mid(rs.Fields("itemname"), 1, 41)) & Space(iname) & Space(2) & Space(iqty) & Format(rs.Fields("qty"), "0.00")
        
            X = X + 1
            If X = 60 Then
                X = 10
                Print #1, Chr(12)
                Print #1, ""
                Print #1, Space(16) & "Sumathi Stores"
                Print #1, Space(16) & "No 5A, RAILWAY STATION ROAD"
                Print #1, Space(19) & "METTUPALAYAM - 641301"
                Print #1, Space(19) & "CELL NO - 90431 01082"
                Print #1, ""
                Print #1, "Current Stock as on " & Format(Date, "DD/MM/YYYY") & Space(24) & "     Time: " & Time()
                Print #1, "--------------------------------------------------------------"           ' 62
                Print #1, "Item Code" & Space(2) & "Item Name " & Space(31) & Space(2) & "Quantity"
                Print #1, "--------------------------------------------------------------"
            End If
        End If
        rs.MoveNext
    Wend
    Print #1, "--------------------------------------------------------------"
    Print #1, Space(47) & "Total: " & Space(iqty) & Format(tqty, "0.00")
    Print #1, "--------------------------------------------------------------"
    Close #1
    retval = Shell("notepad.exe rptcurrentstock.txt", vbMaximizedFocus)
    
    Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
    Print #1, "TYPE rptcurrentstock.txt>PRN"
    Print #1, "EXIT"
    Close #1
    retval = Shell(App.Path & "\print.bat", vbHide)
    
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Call CloseForms
If Button.Caption = "&Purchase" Then
    PurchaseFrm.Show
End If
If Button.Caption = "&Billing" Then
    SalesFrm.Show
End If
If Button.Caption = "S&ales Voucher" Then
    VoucherFrm.Show
End If
If Button.Caption = "P&urchase Voucher" Then
    VoucherPurchaseFrm.Show
End If
If Button.Caption = "E&xit" Then
    End
End If
End Sub

Public Function CloseForms()
For Each frm In Forms
    If Not frm Is MDIForm1 Then
        Unload frm
    End If
Next
Set frm = Nothing
End Function
