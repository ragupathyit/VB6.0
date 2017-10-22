VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCustomerMaster 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Customer Master"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14385
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtmobno 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox txtpincode 
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
      Left            =   2520
      TabIndex        =   8
      Top             =   5280
      Width           =   3495
   End
   Begin VB.TextBox txtstate 
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
      Left            =   2520
      TabIndex        =   7
      Top             =   4680
      Width           =   3495
   End
   Begin VB.TextBox txtcity 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox txtaddress2 
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
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtaddress1 
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
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   3495
   End
   Begin VB.OptionButton optctyped 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dealer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton optctypen 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtcid 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtcname 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   13200
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
      MICON           =   "FrmCustomerMaster.frx":0000
      PICN            =   "FrmCustomerMaster.frx":001C
      PICH            =   "FrmCustomerMaster.frx":072E
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
      Left            =   2640
      TabIndex        =   10
      ToolTipText     =   "SAVE"
      Top             =   6600
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCustomerMaster.frx":0E40
      PICN            =   "FrmCustomerMaster.frx":0E5C
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
      Left            =   4920
      TabIndex        =   11
      ToolTipText     =   "MODIFY"
      Top             =   6600
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCustomerMaster.frx":156E
      PICN            =   "FrmCustomerMaster.frx":158A
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
      Left            =   9480
      TabIndex        =   13
      ToolTipText     =   "DELETE"
      Top             =   6600
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCustomerMaster.frx":1C9C
      PICN            =   "FrmCustomerMaster.frx":1CB8
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
      Left            =   7200
      TabIndex        =   12
      ToolTipText     =   "CLEAR"
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Clear  "
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
      MICON           =   "FrmCustomerMaster.frx":23CA
      PICN            =   "FrmCustomerMaster.frx":23E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5655
      Left            =   6360
      TabIndex        =   14
      Top             =   840
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "C. Id  |Customer Name                                               |C. Type    "
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No *"
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
      Left            =   360
      TabIndex        =   25
      Top             =   6000
      Width           =   1245
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode"
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
      Left            =   360
      TabIndex        =   24
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer State"
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
      Left            =   360
      TabIndex        =   23
      Top             =   4800
      Width           =   1710
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer City"
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
      Left            =   360
      TabIndex        =   22
      Top             =   4200
      Width           =   1560
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type"
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
      Left            =   360
      TabIndex        =   21
      Top             =   2400
      Width           =   1650
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
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
      Left            =   360
      TabIndex        =   20
      Top             =   3600
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1 *"
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
      Left            =   360
      TabIndex        =   19
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id"
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
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name *"
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
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmCustomerMaster.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER MASTER"
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
      TabIndex        =   15
      Top             =   240
      Width           =   3465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   6480
      Width           =   13815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "FrmCustomerMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_custmaster where cid=" & Trim(Val(txtcid.Text))
MsgBox "Customer Deleted Successfully", vbInformation, "Kovai Cards"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmCustomerMaster.Show
End Sub

Private Sub BtnModify_Click()
db.Execute "delete from tbl_custmaster where cid=" & Trim(Val(txtcid.Text))
'-------------------Validation Starts Here-----------------------------
If txtcname.Text = "" Then
    MsgBox "Enter the Customer Name Properly...", vbInformation, "Kovai Cards"
    txtcname.SetFocus
ElseIf txtaddress1.Text = "" Then
    MsgBox "Enter the Address Line 1 Properly...", vbInformation, "Kovai Cards"
    txtaddress1.SetFocus
ElseIf txtmobno.Text = "" Then
    MsgBox "Enter the Mobile No Properly...", vbInformation, "Kovai Cards"
    txtmobno.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_custmaster", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("cid") = Trim(Val(txtcid.Text))
        rs.Fields("cname") = Trim(UCase(txtcname.Text))
        If optctypen.Value = True Then
            rs.Fields("ctype") = "Customer"
        Else
            rs.Fields("ctype") = "Dealer"
        End If
        rs.Fields("address1") = Trim(UCase(txtaddress1.Text))
        rs.Fields("address2") = Trim(UCase(txtaddress2.Text))
        rs.Fields("city") = Trim(UCase(txtcity.Text))
        rs.Fields("state") = Trim(UCase(txtstate.Text))
        rs.Fields("pincode") = Trim(txtpincode.Text)
        rs.Fields("mobno") = Trim(txtmobno.Text)
    rs.Update
    rs.Close
    MsgBox "Customer Modified Successfully", vbInformation, "Kovai Cards"
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtcname.Text = "" Then
    MsgBox "Enter the Customer Name Properly...", vbInformation, "Kovai Cards"
    txtcname.SetFocus
ElseIf txtaddress1.Text = "" Then
    MsgBox "Enter the Address Line 1 Properly...", vbInformation, "Kovai Cards"
    txtaddress1.SetFocus
ElseIf txtmobno.Text = "" Then
    MsgBox "Enter the Mobile No Properly...", vbInformation, "Kovai Cards"
    txtmobno.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_custmaster", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("cid") = Trim(Val(txtcid.Text))
        rs.Fields("cname") = Trim(UCase(txtcname.Text))
        If optctypen.Value = True Then
            rs.Fields("ctype") = "Customer"
        Else
            rs.Fields("ctype") = "Dealer"
        End If
        rs.Fields("address1") = Trim(UCase(txtaddress1.Text))
        rs.Fields("address2") = Trim(UCase(txtaddress2.Text))
        rs.Fields("city") = Trim(UCase(txtcity.Text))
        rs.Fields("state") = Trim(UCase(txtstate.Text))
        rs.Fields("pincode") = Trim(txtpincode.Text)
        rs.Fields("mobno") = Trim(txtmobno.Text)
    rs.Update
    rs.Close
    MsgBox "Customer Saved Successfully", vbInformation, "Kovai Cards"
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
If rs.State = 1 Then rs.Close
rs.Open "select cid,cname,ctype from tbl_custmaster order by cid", db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("cid") & vbTab & rs.Fields("cname") & vbTab & rs.Fields("ctype")
        rs.MoveNext
    Loop
Else
    MSGrid.Rows = 2
End If
rs.Close
End Function

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "Select cid from tbl_custmaster order by cid desc", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtcid.Text = Val(rs.Fields("cid")) + 1
Else
    txtcid.Text = 1
End If

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_custmaster where cid=" & Trim(Val(MSGrid.TextMatrix(MSGrid.Row, 0))), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtcid.Text = rs1.Fields("cid")
    txtcname.Text = rs1.Fields("cname")
    If rs1.Fields("ctype") = "Customer" Then
        optctypen.Value = True
    Else
        optctyped.Value = True
    End If
    txtaddress1.Text = rs1.Fields("address1")
    txtaddress2.Text = rs1.Fields("address2")
    txtcity.Text = rs1.Fields("city")
    txtstate.Text = rs1.Fields("state")
    txtpincode.Text = rs1.Fields("pincode")
    txtmobno.Text = rs1.Fields("mobno")
End If

txtcname.SetFocus
txtcname.SelStart = 0
txtcname.SelLength = Len(txtcname.Text)    'select the text

BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtcname_Change()
If rs.State = 1 Then rs.Close
rs.Open "select cid,cname,ctype from tbl_custmaster where cname like '" & Trim(txtcname.Text) & "%' order by cid", db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("cid") & vbTab & rs.Fields("cname") & vbTab & rs.Fields("ctype")
        rs.MoveNext
    Loop
End If
rs.Close
End Sub

Private Sub txtcname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress1.SetFocus
    txtaddress1.SelStart = 0
    txtaddress1.SelLength = Len(txtaddress1.Text)    'select the text
End If
End Sub

Private Sub txtaddress1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress2.SetFocus
    txtaddress2.SelStart = 0
    txtaddress2.SelLength = Len(txtaddress2.Text)    'select the text
End If
End Sub

Private Sub txtaddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcity.SetFocus
    txtcity.SelStart = 0
    txtcity.SelLength = Len(txtcity.Text)    'select the text
End If
End Sub

Private Sub txtcity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtstate.SetFocus
    txtstate.SelStart = 0
    txtstate.SelLength = Len(txtstate.Text)    'select the text
End If
End Sub

Private Sub txtstate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpincode.SetFocus
    txtpincode.SelStart = 0
    txtpincode.SelLength = Len(txtpincode.Text)    'select the text
End If
End Sub

Private Sub txtpincode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtmobno.SetFocus
    txtmobno.SelStart = 0
    txtmobno.SelLength = Len(txtmobno.Text)    'select the text
End If
End Sub

Private Sub txtmobno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
