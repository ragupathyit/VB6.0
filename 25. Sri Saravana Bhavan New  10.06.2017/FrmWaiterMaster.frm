VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmWaiterMaster 
   BackColor       =   &H00400040&
   Caption         =   "Waiter Master"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   15405
   WindowState     =   2  'Maximized
   Begin VB.ListBox lst_table 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   3240
      Style           =   1  'Checkbox
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtwaiterid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtwname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtmobileno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtuname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtcpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   8535
      Left            =   7560
      TabIndex        =   6
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   15055
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Waiter Id  |Waiter Name                                   |Mobile No        |Table No    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmWaiterMaster.frx":0000
      PICN            =   "FrmWaiterMaster.frx":001C
      PICH            =   "FrmWaiterMaster.frx":072E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnSave 
      Height          =   495
      Left            =   720
      TabIndex        =   17
      ToolTipText     =   "SAVE"
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Save   "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmWaiterMaster.frx":0E40
      PICN            =   "FrmWaiterMaster.frx":0E5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnModify 
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      ToolTipText     =   "MODIFY"
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Modify"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmWaiterMaster.frx":156E
      PICN            =   "FrmWaiterMaster.frx":158A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnClear 
      Height          =   495
      Left            =   4080
      TabIndex        =   19
      ToolTipText     =   "CLEAR"
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Refresh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmWaiterMaster.frx":1C9C
      PICN            =   "FrmWaiterMaster.frx":1CB8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnDelete 
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      ToolTipText     =   "DELETE"
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12583104
      BCOLO           =   12583104
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmWaiterMaster.frx":23CA
      PICN            =   "FrmWaiterMaster.frx":23E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table No *"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   14
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waiter Id"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   13
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waiter Name *"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   12
      Top             =   1680
      Width           =   1605
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   11
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name *"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password *"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password *"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   7560
      Width           =   7575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WAITER MASTER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   240
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmWaiterMaster.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   8535
      Left            =   0
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "FrmWaiterMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_waitermaster where waiterid=" & Val(txtwaiterid.Text)
MsgBox "Record is Deleted Successfully", vbInformation, "Sri Saravana Bhavan"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmWaiterMaster.Show
End Sub

Private Sub BtnModify_Click()
Call update("Modified Successfully")
End Sub

Function update(str)
'-------------------Validation Starts Here-----------------------------
If txtwname.Text = "" Then
    MsgBox "Enter the Waiter Name Properly...", vbInformation, "Sri Saravana Bhavan"
    txtwname.SetFocus
ElseIf lst_table.SelCount = 0 Then
    MsgBox "Select the Dining Table Properly...", vbInformation, "Sri Saravana Bhavan"
    lst_table.SetFocus
'ElseIf txtuname.Text = "" Then
'    MsgBox "Enter the User Name Properly...", vbInformation, "Sri Saravana Bhavan"
'    txtuname.SetFocus
'ElseIf txtpassword.Text = "" Then
'    MsgBox "Enter the Password Properly...", vbInformation, "Sri Saravana Bhavan"
'    txtpassword.SetFocus
'ElseIf txtcpassword.Text = "" Then
'    MsgBox "Enter the Confirm Password Properly...", vbInformation, "Sri Saravana Bhavan"
'    txtcpassword.SetFocus
'ElseIf txtcpassword.Text <> txtpassword.Text Then
'    MsgBox "Confirm Password Not Matching With Password...", vbInformation, "Sri Saravana Bhavan"
'    txtcpassword.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    
    db.Execute "delete from tbl_waitermaster where waiterid=" & Val(txtwaiterid.Text)
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_waitermaster", db, adOpenDynamic, adLockOptimistic
    For i = 0 To lst_table.ListCount - 1
        If lst_table.Selected(i) = True Then
            rs.AddNew
            rs.Fields("waiterid") = Val(Trim(txtwaiterid.Text))
            rs.Fields("waitername") = UCase(Trim(txtwname.Text))
            rs.Fields("mobileno") = Trim(txtmobileno.Text)
            rs.Fields("tableno") = lst_table.List(i)
            'rs.Fields("username") = Trim(txtuname.Text)
            'rs.Fields("password") = Trim(txtpassword.Text)
            rs.update
        End If
    Next i
    rs.Close
    
    MsgBox "Record " & str, vbInformation, "Sri Saravana Bhavan"
    Call BtnClear_Click

End If
End Function

Private Sub BtnSave_Click()
Call update("Saved Successfully")
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub


Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select waiterid from tbl_waitermaster order by waiterid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtwaiterid.Text = rs.Fields("waiterid") + 1
Else
    txtwaiterid.Text = 1
End If
rs.Close

'------------------adding the tableno in the lst_table listbox
rs.Open "select tableno from tbl_tablemaster order by tableno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    While Not rs.EOF
        lst_table.AddItem Trim(rs.Fields("tableno"))
        rs.MoveNext
    Wend
End If
rs.Close

'------------------removing the tableno from the lst_table listbox, because it is already maintained by other waiter
rs.Open "select tableno from tbl_waitermaster", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    If rs.Fields("tableno") <> "" Then
        For i = 0 To lst_table.ListCount - 1
            If lst_table.List(i) = rs.Fields("tableno") Then
                lst_table.RemoveItem i
            End If
        Next i
    End If
    rs.MoveNext
Wend
rs.Close

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

Me.Show
txtwname.SetFocus

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub
Private Function Fill()
stmt = "select * from tbl_waitermaster order by waiterid"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("waiterid") & vbTab & rs.Fields("waitername") & vbTab & rs.Fields("mobileno") & vbTab & rs.Fields("tableno")
        rs.MoveNext
    Wend
End If
rs.Close
End Function

Private Sub lst_table_ItemCheck(Item As Integer)
'If lst_table.SelCount > 3 Then
'    lst_table.Selected(Item) = False
'    MsgBox "Per Waiter Maximum 3 Tables Only Allowed", vbInformation, "Sri Saravana Bhavan"
'End If
End Sub

Private Sub lst_table_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'txtuname.SetFocus
    'txtuname.SelStart = 0
    'txtuname.SelLength = Len(txtuname.Text)    'select the text
    
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub

Private Sub MSGrid_Click()
If MSGrid.Rows > 1 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_waitermaster where waiterid=" & Val(MSGrid.TextMatrix(MSGrid.Row, 0)) & " order by tableno", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txtwaiterid.Text = rs.Fields("waiterid")
        txtwname.Text = rs.Fields("waitername")
        txtmobileno.Text = rs.Fields("mobileno")
        'txtuname.Text = rs.Fields("username")
        'txtpassword.Text = rs.Fields("password")
        'txtcpassword.Text = rs.Fields("password")
        
        '------------------removing the tableno from the lst_table listbox, because it is already maintained by other waiter
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select tableno from tbl_waitermaster order by tableno", db, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
            If rs1.Fields("tableno") <> "" Then
                For i = 0 To lst_table.ListCount - 1
                    If lst_table.List(i) = rs1.Fields("tableno") Then
                        lst_table.RemoveItem i
                    End If
                Next i
            End If
            rs1.MoveNext
        Wend
        rs1.Close
        
        '-----------------------------added the allready selected table name based on the waiter
        While Not rs.EOF
            If rs.Fields("tableno") <> "" Then
                lst_table.AddItem rs.Fields("tableno")
                For i = 0 To lst_table.ListCount - 1
                    If lst_table.List(i) = rs.Fields("tableno") Then
                        lst_table.Selected(i) = True
                    End If
                Next i
            End If
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    txtwname.SetFocus
    txtwname.SelStart = 0
    txtwname.SelLength = Len(txtwname.Text)
    'txtwaiterid.Enabled = False
    
    BtnSave.Enabled = False
    BtnModify.Enabled = True
    BtnDelete.Enabled = True
End If
End Sub

Private Sub txtwname_Change()
stmt = "select * from tbl_waitermaster where waitername like '" & Trim(txtwname.Text) & "%' order by tableno"
If rs3.State = 1 Then rs3.Close
rs3.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs3.EOF Then
    rs3.MoveFirst
    Do While Not rs3.EOF
        MSGrid.AddItem rs3.Fields("waiterid") & vbTab & rs3.Fields("waitername") & vbTab & rs3.Fields("mobileno") & vbTab & rs3.Fields("tableno")
        rs3.MoveNext
    Loop
End If
rs3.Close
End Sub

Private Sub txtwname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtmobileno.SetFocus
    txtmobileno.SelStart = 0
    txtmobileno.SelLength = Len(txtmobileno.Text)    'select the text
End If
End Sub
Private Sub txtmobileno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lst_table.SetFocus
End If
End Sub
'Private Sub txtuname_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtpassword.SetFocus
'    txtpassword.SelStart = 0
'    txtpassword.SelLength = Len(txtpassword.Text)    'select the text
'End If
'End Sub
'Private Sub txtpassword_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtcpassword.SetFocus
'    txtcpassword.SelStart = 0
'    txtcpassword.SelLength = Len(txtcpassword.Text)    'select the text
'End If
'End Sub
'Private Sub txtcpassword_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If BtnSave.Enabled = True Then
'        BtnSave.SetFocus
'    Else
'        BtnModify.SetFocus
'    End If
'End If
'End Sub
