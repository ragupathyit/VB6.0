VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmItemMaster 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Item Master"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14130
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtcrate 
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
      Left            =   2280
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtdrate 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtiid 
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtiname 
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
      Left            =   2280
      TabIndex        =   1
      Top             =   2160
      Width           =   3735
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   12600
      TabIndex        =   9
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
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "SAVE"
      Top             =   5040
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
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "MODIFY"
      Top             =   5040
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
   Begin Project1.Button BtnDelete 
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      ToolTipText     =   "DELETE"
      Top             =   5040
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
   Begin Project1.Button BtnClear 
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      ToolTipText     =   "CLEAR"
      Top             =   5040
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
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   4095
      Left            =   6360
      TabIndex        =   7
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Item Id  |Item Name                                                          "
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Rate *"
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
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Rate *"
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
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Id *"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   1005
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
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   1410
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
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   4920
      Width           =   13215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "FrmItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_itemmaster where iid='" & Trim(UCase(txtiid.Text)) & "'"
MsgBox "Item Deleted Successfully", vbInformation, "Kovai Cards"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmItemMaster.Show
End Sub

Private Sub BtnModify_Click()
db.Execute "delete from tbl_itemmaster where iid='" & Trim(UCase(txtiid.Text)) & "'"
'-------------------Validation Starts Here-----------------------------
If txtiid.Text = "" Then
    MsgBox "Enter the Item Id Properly...", vbInformation, "Kovai Cards"
    txtiid.SetFocus
ElseIf txtiname.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Kovai Cards"
    txtiname.SetFocus
ElseIf txtcrate.Text = "" Then
    MsgBox "Enter the Customer Rate Properly...", vbInformation, "Kovai Cards"
    txtcrate.SetFocus
'ElseIf IsNumeric(Val(txtcrate.Text)) Then
'    MsgBox "Enter the Customer Rate Properly...", vbInformation, "Kovai Cards"
'    txtcrate.SetFocus
'    txtcrate.SelStart = 0
'    txtcrate.SelLength = Len(txtcrate.Text)    'select the text
ElseIf txtdrate.Text = "" Then
    MsgBox "Enter the Dealer Rate Properly...", vbInformation, "Kovai Cards"
    txtdrate.SetFocus
'ElseIf IsNumeric(Val(txtdrate.Text)) Then
'    MsgBox "Enter the Dealer Rate Properly...", vbInformation, "Kovai Cards"
'    txtdrate.SetFocus
'    txtdrate.SelStart = 0
'    txtdrate.SelLength = Len(txtdrate.Text)    'select the text
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where iid='" & Trim(UCase(txtiid.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Item is Already Saved, Type Different Item Id", vbInformation, "Kovai Cards"
    Else
        rs.AddNew
        rs.Fields("iid") = Trim(UCase(txtiid.Text))
        rs.Fields("iname") = Trim(UCase(txtiname.Text))
        rs.Fields("crate") = Format(Trim(txtcrate.Text), "0.00")
        rs.Fields("drate") = Format(Trim(txtdrate.Text), "0.00")
        rs.Update
        rs.Close
        MsgBox "Item Modified Successfully", vbInformation, "Kovai Cards"
    End If
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtiid.Text = "" Then
    MsgBox "Enter the Item Id Properly...", vbInformation, "Kovai Cards"
    txtiid.SetFocus
ElseIf txtiname.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Kovai Cards"
    txtiname.SetFocus
ElseIf txtcrate.Text = "" Then
    MsgBox "Enter the Customer Rate Properly...", vbInformation, "Kovai Cards"
    txtcrate.SetFocus
'ElseIf IsNumeric(Val(txtcrate.Text)) Then
'    MsgBox "Enter the Customer Rate Properly...", vbInformation, "Kovai Cards"
'    txtcrate.SetFocus
'    txtcrate.SelStart = 0
'    txtcrate.SelLength = Len(txtcrate.Text)    'select the text
ElseIf txtdrate.Text = "" Then
    MsgBox "Enter the Dealer Rate Properly...", vbInformation, "Kovai Cards"
    txtdrate.SetFocus
'ElseIf IsNumeric(Val(txtdrate.Text)) Then
'    MsgBox "Enter the Dealer Rate Properly...", vbInformation, "Kovai Cards"
'    txtdrate.SetFocus
'    txtdrate.SelStart = 0
'    txtdrate.SelLength = Len(txtdrate.Text)    'select the text
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where iid='" & Trim(UCase(txtiid.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Item is Already Saved, Type Different Item Id", vbInformation, "Kovai Cards"
    Else
        rs.AddNew
        rs.Fields("iid") = Trim(UCase(txtiid.Text))
        rs.Fields("iname") = Trim(UCase(txtiname.Text))
        rs.Fields("crate") = Format(Trim(txtcrate.Text), "0.00")
        rs.Fields("drate") = Format(Trim(txtdrate.Text), "0.00")
        rs.Update
        rs.Close
        MsgBox "Item Saved Successfully", vbInformation, "Kovai Cards"
    End If
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
stmt = "select iid,iname from tbl_itemmaster order by iid"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("iid") & vbTab & rs.Fields("iname")
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
rs1.Open "select * from tbl_itemmaster where iid='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtiid.Text = rs1.Fields("iid")
    txtiname.Text = rs1.Fields("iname")
    txtcrate.Text = Format(rs1.Fields("crate"), "0.00")
    txtdrate.Text = Format(rs1.Fields("drate"), "0.00")
End If

txtiname.SetFocus
BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtiid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtiname.SetFocus
    txtiname.SelStart = 0
    txtiname.SelLength = Len(txtiname.Text)    'select the text
End If
End Sub

Private Sub txtiname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcrate.SetFocus
    txtcrate.SelStart = 0
    txtcrate.SelLength = Len(txtcrate.Text)    'select the text
End If
End Sub

Private Sub txtcrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdrate.SetFocus
    txtdrate.SelStart = 0
    txtdrate.SelLength = Len(txtdrate.Text)    'select the text
End If
End Sub

Private Sub txtdrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub
