VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DBdaybookfrm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Day Book"
   ClientHeight    =   8235
   ClientLeft      =   -60
   ClientTop       =   -75
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttotamtcr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   6960
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   97386497
      CurrentDate     =   40597
   End
   Begin VB.CommandButton CmdReport 
      BackColor       =   &H00C0E0FF&
      Caption         =   "RE&PORT"
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
      Left            =   6600
      TabIndex        =   6
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txtcbalance 
      Alignment       =   1  'Right Justify
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
      Left            =   2640
      TabIndex        =   9
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtobalance 
      Alignment       =   1  'Right Justify
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txttotamtdr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&SAVE"
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
      Left            =   960
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0E0FF&
      Caption         =   "D&ELETE"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "E&XIT"
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
      Left            =   8520
      TabIndex        =   7
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C&LEAR"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16711680
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Particulars                                                                              | Cr Amount       | Dr Amount       "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblsalesamt 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-- Nil --"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5640
      TabIndex        =   18
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Todays Sales   :"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3000
      TabIndex        =   17
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Balance"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   0
      TabIndex        =   15
      Top             =   7080
      Width           =   2550
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   2685
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amt"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   4935
      TabIndex        =   13
      Top             =   7080
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   7440
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAL DAY BOOK EXPENCES"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   0
      Width           =   6720
   End
End
Attribute VB_Name = "DBdaybookfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
Unload Me
DBdaybookfrm.Show
End Sub

Private Sub CmdDelete_Click()
If Not txtid.Text = "" Then
    db.Execute "delete from tbl_DBbalance where did=" & Val(txtid.Text)
    db.Execute "delete from tbl_DBdebit where did=" & Val(txtid.Text)
    db.Execute "delete from tbl_DBcredit where did=" & Val(txtid.Text)
    MsgBox "Successfully Deleted", vbInformation, "KPS Hospital"
    Call cmdclear_Click
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub CmdReport_Click()
If Not txtid.Text = "" Then
    '----------Notepad print------------------
    Open App.Path & "\DBDayreport.txt" For Output As #1
    Print #1, Space(13) & "K.P.S HOSPITALS (P) LTD"
    Print #1, Space(8) & "ANNUR ROAD, METTUPALAYAM - 641 301"
    Print #1, Space(14) & "04254-224314, 224315"
    Print #1, ""
    Print #1, ""
    Print #1, " Time  =" & Time & Space(26) & "Date  =" & Format(Date, "DD/MM/YYYY")
    Print #1, ""
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBbalance where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    Print #1, Space(1) & "Day book as on " & Format(rs.Fields("pdate"), "DD/MM/YYYY")
    Print #1, Space(1) & "============================================================"
    Print #1, Space(1) & " Particulars" & Space(30) & "Cr Amt" & Space(4) & "Dr Amt"
    Print #1, Space(1) & "============================================================"
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBbalance where did=" & Val(txtid.Text) - 1, db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        iamt = 10 - Len(Format(rs.Fields("cbalance"), "0.00"))
        Print #1, Space(1) & " Opening Balance" & Space(23) & Space(iamt) & Format(rs.Fields("cbalance"), "0.00")
    Else
        iamt = 10 - Len("0.00")
        Print #1, Space(1) & " Opening Balance" & Space(23) & Space(iamt) & "0.00"
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBcredit where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        iparticulars = 38 - Len(rs.Fields("particulars"))
        iamt = 10 - Len(Format(rs.Fields("amt"), "0.00"))
        Print #1, Space(2) & UCase(rs.Fields("particulars")) & Space(iparticulars) & Space(iamt) & Format(rs.Fields("amt"), "0.00")
        i = i + 1
        rs.MoveNext
    Wend
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBdebit where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        iparticulars = 48 - Len(rs.Fields("particulars"))
        iamt = 10 - Len(Format(rs.Fields("amt"), "0.00"))
        Print #1, Space(2) & UCase(rs.Fields("particulars")) & Space(iparticulars) & Space(iamt) & Format(rs.Fields("amt"), "0.00")
        i = i + 1
        rs.MoveNext
    Wend
    
    Print #1, Space(1) & "------------------------------------------------------------"
    
    If rs.State = 1 Then rs.Close
    If rs1.State = 1 Then rs1.Close
    rs.Open "select * from tbl_DBcredit where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    
    If rs1.State = 1 Then rs1.Close
    If Val(txtid.Text) = 1 Then
        cb = 0
    Else
        rs1.Open "select * from tbl_DBbalance where did=" & Val(txtid.Text) - 1, db, adOpenDynamic, adLockOptimistic
        cb = Val(rs1.Fields("cbalance"))
    End If
    totcr = Val(rs.Fields("totamt")) + Val(cb)
    
    iamt = 10 - Len(Format(totcr, "0.00"))
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBdebit where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    iamt1 = 10 - Len(Format(rs.Fields("totamt"), "0.00"))
    
    Print #1, Space(2) & "Total" & Space(33) & Space(iamt) & Format(totcr, "0.00") & Space(iamt1) & Format(rs.Fields("totamt"), "0.00")
    Print #1, Space(41) & "--------------------"
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBbalance where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    iamt = 10 - Len(Format(rs.Fields("cbalance"), "0.00"))
    Print #1, Space(1) & " Closing Balance" & Space(23) & Space(iamt) & Format(rs.Fields("cbalance"), "0.00")
    
    Print #1, ""
    Print #1, Space(15) & "Thank You! Visit Again..."
    Close #1
    retval = Shell("notepad.exe DBDayreport.txt", vbMaximizedFocus)
    '-----------------------------------------------------------------------------------------------------
    s = MsgBox("Do You Want Print", vbYesNo)
    If s = vbYes Then
        Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
        Print #1, "TYPE DBDayreport.txt>PRN"
        Print #1, "EXIT"
        Close #1
        retval = Shell(App.Path & "\print.bat", vbHide)
    End If
End If
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_DBbalance where pdate=#" & DTPicker1.Value & "#", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    MsgBox "Day book was already saved in this date", vbInformation, "KPS Hospital"
Else
    rs.AddNew
    rs.Fields("did") = txtid.Text
    rs.Fields("pdate") = DTPicker1.Value
    rs.Fields("obalance") = Val(txtobalance.Text)
    rs.Fields("cbalance") = Val(txtcbalance.Text)
    rs.Update

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBcredit", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        If Not MSGrid.TextMatrix(i, 1) = "0" Then
            rs.AddNew
            rs.Fields("did") = txtid.Text
            rs.Fields("cdate") = DTPicker1.Value
            rs.Fields("particulars") = MSGrid.TextMatrix(i, 0)
            rs.Fields("amt") = Val(MSGrid.TextMatrix(i, 1))
            rs.Fields("totamt") = Val(txttotamtcr.Text) + Val(lblsalesamt.Caption)
            rs.Update
        End If
    Next i
    rs.AddNew
    rs.Fields("did") = txtid.Text
    rs.Fields("cdate") = DTPicker1.Value
    rs.Fields("particulars") = "Todays Sales"
    rs.Fields("amt") = Val(lblsalesamt.Caption)
    rs.Fields("totamt") = Val(txttotamtcr.Text) + Val(lblsalesamt.Caption)
    rs.Update
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_DBdebit", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        If Not MSGrid.TextMatrix(i, 2) = "0" Then
            rs.AddNew
            rs.Fields("did") = txtid.Text
            rs.Fields("ddate") = DTPicker1.Value
            rs.Fields("particulars") = MSGrid.TextMatrix(i, 0)
            rs.Fields("amt") = Val(MSGrid.TextMatrix(i, 2))
            rs.Fields("totamt") = Val(txttotamtdr.Text)
            rs.Update
        End If
    Next i

    MsgBox "Successfully Saved", vbInformation, "KPS Hospital"

    CmdReport.Enabled = True
    CmdReport.SetFocus

End If
End Sub

Private Sub DTPicker1_Change()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_DBbalance where pdate=#" & DTPicker1.Value & "#", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtid.Text = rs.Fields("did")
End If
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_DBbalance where did=" & Val(txtid.Text) & " order by did", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtcbalance.Text = rs.Fields("cbalance")
    txtobalance.Text = rs.Fields("obalance")
Else
    txtobalance.Text = 0
End If
        
MSGrid.Rows = 2
MSGrid.TextMatrix(1, 0) = ""
MSGrid.TextMatrix(1, 1) = ""
MSGrid.TextMatrix(1, 2) = ""
        
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_DBcredit where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    DTPicker1.Value = rs.Fields("cdate")
    txttotamtcr.Text = rs.Fields("totamt")
    i = 1
    
    If rs.Fields("particulars") = "Todays Sales" Then
        lblsalesamt.Caption = rs.Fields("amt")
    End If
    rs.MoveNext
    
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("particulars")
        MSGrid.TextMatrix(i, 1) = rs.Fields("amt")
        MSGrid.Rows = MSGrid.Rows + 1
        i = i + 1
        rs.MoveNext
    Wend
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from tbl_DBdebit where did=" & Val(txtid.Text), db, adOpenDynamic, adLockOptimistic
    txttotamtdr.Text = rs1.Fields("totamt")
    While Not rs1.EOF
        MSGrid.TextMatrix(i, 0) = rs1.Fields("particulars")
        MSGrid.TextMatrix(i, 2) = rs1.Fields("amt")
        MSGrid.Rows = MSGrid.Rows + 1
        i = i + 1
        rs1.MoveNext
    Wend
    rs1.Close
    
Else
    MsgBox "No Records Found", vbInformation, "KPS Hospital"
End If
rs.Close
    
CmdSave.Enabled = False
cmdclear.Enabled = True
CmdDelete.Enabled = True
CmdReport.Enabled = True
CmdReport.SetFocus

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.Col = 1
    MSGrid.Row = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub Form_Load()
'Me.BackColor = RGB(84, 96, 254)
'Label2.BackColor = RGB(84, 96, 254)
'Label3.BackColor = RGB(84, 96, 254)
'Label5.BackColor = RGB(84, 96, 254)
'Label6.BackColor = RGB(84, 96, 254)
'Label7.BackColor = RGB(84, 96, 254)
'Label8.BackColor = RGB(84, 96, 254)
'lblsalesamt.BackColor = RGB(84, 96, 254)
'MSGrid.BackColorBkg = RGB(84, 96, 254)

Call connect
DTPicker1.Value = Date

rs.Open "select * from tbl_DBbalance order by did", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtid.Text = Val(rs.Fields("did")) + 1
    txtobalance.Text = rs.Fields("cbalance")
Else
    txtid.Text = 1
    txtobalance.Text = 0
End If
rs.Close

MSGrid.Row = 1
MSGrid.Col = 0
MSGrid.CellBackColor = RGB(117, 145, 233)

CmdSave.Enabled = True
cmdclear.Enabled = True
CmdDelete.Enabled = False
CmdReport.Enabled = False

End Sub

Private Sub MSGrid_EnterCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub MSGrid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If

End Sub

Private Sub MSGrid_LeaveCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = vbWhite
End Sub

Private Sub MSGrid_KeyPress(KeyAscii As Integer)
If MSGrid.Col = 0 Or MSGrid.Col = 1 Or MSGrid.Col = 2 Then ' first and second grid coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
    Case 32   ' 32 keyascii is for Space Bar space
        If MSGrid.Col = 0 Then
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 46         ' 46 keyascii is for dot symbol
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
        If MSGrid.Col = 0 Then
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 97 To 122  ' 97-122 keyascii is for small a to z
        If MSGrid.Col = 0 Then
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 0 Then
            If Not MSGrid.TextMatrix(MSGrid.Row, 0) = "" Then
                MSGrid.Col = MSGrid.Col + 1  ' Grid entry was changed to 2nd grid coloumn
            Else
                txtcbalance.SetFocus
            End If
        End If
        
        If MSGrid.Col = 1 Then
            If Not MSGrid.TextMatrix(MSGrid.Row, 1) = "" Then
                
                txttotamtcr.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotamtcr.Text = Val(txttotamtcr.Text) + Val(MSGrid.TextMatrix(i, 1))   'Total bill amount calculation
                Next i
                
                MSGrid.Col = MSGrid.Col + 1                'cursor position changed to the first coloumn of that newly created row
            Else
                MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = 0
                MSGrid.Col = MSGrid.Col + 1                'cursor position changed to the first coloumn of that newly created row
            End If
        End If
        
        If MSGrid.Col = 2 Then
            If Not MSGrid.TextMatrix(MSGrid.Row, 2) = "" Then
                
                txttotamtdr.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotamtdr.Text = Val(txttotamtdr.Text) + Val(MSGrid.TextMatrix(i, 2))   'Total bill amount calculation
                Next i
                
                If MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = "" Then
                    MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                End If
                
                MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                MSGrid.Row = MSGrid.Row + 1     'cursor position changed to the newlly created row
                MSGrid.Col = 0                  'cursor position changed to the first coloumn of that newly created row
                
            Else
                MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = 0
            End If
        End If
    End Select
End If
End Sub

Private Sub txtcbalance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    a = Val(txtobalance.Text) + Val(txttotamtcr.Text)
    b = Val(txtcbalance.Text) + Val(txttotamtdr.Text)
    lblsalesamt.Caption = Val(b) - Val(a)
    
    If CmdSave.Enabled = False Then
        CmdReport.SetFocus
    Else
        CmdSave.SetFocus
    End If
End If
End Sub
