VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmItemMaster 
   BackColor       =   &H00400040&
   Caption         =   "Item Master"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12750
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtitemid 
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
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   4320
      Width           =   2895
      Begin VB.OptionButton optchinese 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chinese"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optindian 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Indian"
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
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtitemcode 
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtitemname 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtitemprice 
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
      Left            =   3480
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6375
      Left            =   7800
      TabIndex        =   4
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "I|Item Name                                             |Item Id   |Item Price    |Item Type    "
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
      Left            =   7200
      TabIndex        =   14
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
      MICON           =   "FrmItemMaster.frx":0000
      PICN            =   "FrmItemMaster.frx":001C
      PICH            =   "FrmItemMaster.frx":072E
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
      Left            =   840
      TabIndex        =   15
      ToolTipText     =   "SAVE"
      Top             =   5640
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
      MICON           =   "FrmItemMaster.frx":0E40
      PICN            =   "FrmItemMaster.frx":0E5C
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
      Left            =   2520
      TabIndex        =   16
      ToolTipText     =   "MODIFY"
      Top             =   5640
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
      MICON           =   "FrmItemMaster.frx":156E
      PICN            =   "FrmItemMaster.frx":158A
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
      Left            =   4200
      TabIndex        =   17
      ToolTipText     =   "CLEAR"
      Top             =   5640
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
      MICON           =   "FrmItemMaster.frx":1C9C
      PICN            =   "FrmItemMaster.frx":1CB8
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
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "DELETE"
      Top             =   5640
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
      MICON           =   "FrmItemMaster.frx":23CA
      PICN            =   "FrmItemMaster.frx":23E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Id"
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
      Left            =   1200
      TabIndex        =   19
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   4440
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Left            =   3120
      TabIndex        =   9
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name *"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   1410
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Price *"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmItemMaster.frx":2AF8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM MASTER"
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
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   5400
      Width           =   8655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "FrmItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnDelete_Click()
db.Execute "delete from tbl_itemmaster where itemcode=" & Val(txtitemcode.Text)
MsgBox "Item is Deleted Successfully", vbInformation, "Sri Saravana Bhavan"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmItemMaster.Show
End Sub

Private Sub BtnModify_Click()
Call update
Call BtnClear_Click
End Sub

Function update()
'-------------------Validation Starts Here-----------------------------
If txtitemname.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Sri Saravana Bhavan"
    txtitemname.SetFocus
ElseIf txtitemprice.Text = "" Then
    MsgBox "Enter the Item Price Properly...", vbInformation, "Sri Saravana Bhavan"
    txtitemprice.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from tbl_itemmaster where itemid='" & UCase(Trim(txtitemid.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        MsgBox "This Item Id is already taken. Please Choose Different Item Id", vbInformation, "Sri Saravana Bhavan"
        txtitemid.Text = ""
        txtitemid.SetFocus
    Else
        db.Execute "delete from tbl_itemmaster where itemcode=" & Val(txtitemcode.Text)
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_itemmaster", db, adOpenDynamic, adLockOptimistic
        rs.AddNew
        rs.Fields("itemcode") = Val(Trim(txtitemcode.Text))
        rs.Fields("itemname") = UCase(Trim(txtitemname.Text))
        rs.Fields("itemid") = UCase(Trim(txtitemid.Text))
        rs.Fields("price") = Format(Val(txtitemprice.Text), "0.00")
        If optindian.Value = True Then
            rs.Fields("itemtype") = "Indian"
        Else
            rs.Fields("itemtype") = "Chinese"
        End If
        rs.update
        rs.Close
        MsgBox "Item Modified Successfully", vbInformation, "Sri Saravana Bhavan"
    End If
    rs1.Close
End If
End Function

Private Sub BtnSave_Click()
Call update
MsgBox "Item Saved Successfully", vbInformation, "Sri Saravana Bhavan"
Call BtnClear_Click
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select itemcode from tbl_itemmaster order by itemcode", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtitemcode.Text = Val(rs.Fields("itemcode")) + 1
Else
    txtitemcode.Text = 1
End If
rs.Close

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

Me.Show
txtitemname.SetFocus

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemname"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    MSGrid.Rows = 1
    rs.MoveFirst
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & IIf(IsNull(rs.Fields("itemid")), "", rs.Fields("itemid")) & vbTab & Format(rs.Fields("price"), "0.00") & vbTab & rs.Fields("itemtype")
        rs.MoveNext
    Wend
End If
rs.Close
End Function

Private Sub MSGrid_Click()
If MSGrid.Rows > 1 Then
    If rs3.State = 1 Then rs3.Close
    rs3.Open "select * from tbl_itemmaster where itemcode=" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)), db, adOpenDynamic, adLockOptimistic
    If Not rs3.EOF Then
        txtitemcode.Text = rs3.Fields("itemcode")
        txtitemname.Text = rs3.Fields("itemname")
        txtitemid.Text = IIf(IsNull(rs3.Fields("itemid")), "", rs3.Fields("itemid"))
        txtitemprice.Text = rs3.Fields("price")
        If rs3.Fields("itemtype") = "Indian" Then
            optindian.Value = True
        Else
            optchinese.Value = True
        End If
        
        txtitemname.SetFocus
        txtitemname.SelStart = 0
        txtitemname.SelLength = Len(txtitemname.Text)
    End If
    
    BtnSave.Enabled = False
    BtnModify.Enabled = True
    BtnDelete.Enabled = True
End If
End Sub

Private Sub txtitemcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtitemname.SetFocus
    txtitemname.SelStart = 0
    txtitemname.SelLength = Len(txtitemname.Text)    'select the text
End If
End Sub

Private Sub txtitemname_Change()
stmt = "select * from tbl_itemmaster where itemname like'" & Trim(txtitemname.Text) & "%' order by itemname"
If rs1.State = 1 Then rs1.Close
rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs1.EOF Then
    rs1.MoveFirst
    While Not rs1.EOF
        MSGrid.AddItem rs1.Fields("itemcode") & vbTab & rs1.Fields("itemname") & vbTab & rs1.Fields("itemid") & vbTab & Format(rs1.Fields("price"), "0.00") & vbTab & rs1.Fields("itemtype")
        rs1.MoveNext
    Wend
End If
rs1.Close
End Sub

Private Sub txtitemname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtitemid.SetFocus
    txtitemid.SelStart = 0
    txtitemid.SelLength = Len(txtitemid.Text)    'select the text
End If
End Sub

Private Sub txtitemid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtitemprice.SetFocus
    txtitemprice.SelStart = 0
    txtitemprice.SelLength = Len(txtitemprice.Text)    'select the text
End If
End Sub

Private Sub txtitemprice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
