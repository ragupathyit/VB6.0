VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ItemFrm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Item Details"
   ClientHeight    =   6630
   ClientLeft      =   105
   ClientTop       =   -15
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   16755.03
   ScaleMode       =   0  'User
   ScaleWidth      =   26419.7
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttamilname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tamil-Aiswarya"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   16
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtitemrate 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "S&AVE"
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
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdModify 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&MODIFY"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&DELETE"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
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
      Left            =   2880
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
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
      Left            =   3600
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtitemname 
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
      Height          =   375
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6615
      Left            =   5400
      TabIndex        =   7
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483643
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "I Code   |Item Name                                                              |I.Rate      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtitemcode 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tamil Name"
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
      Left            =   480
      TabIndex        =   15
      Top             =   2760
      Width           =   1875
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Rate"
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
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   1530
   End
   Begin MSForms.ComboBox cmbqtytype 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   3600
      Width           =   2295
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4048;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Type"
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
      Left            =   480
      TabIndex        =   11
      Top             =   3600
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM DETAILS"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   240
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   1710
   End
End
Attribute VB_Name = "ItemFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
Unload Me
ItemFrm.Show
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
db.Execute "delete from tbl_itemmaster where itemcode='" & Trim(txtitemcode.Text) & "'"
'db.Execute "delete from tbl_stock where itemcode='" & Trim(txtitemcode.Text) & "'"
MsgBox "Successfully Deleted...", vbInformation, "Sumathi Stores"
Call cmdclear_Click
End Sub

Private Sub CmdModify_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_itemmaster where itemcode='" & Trim(txtitemcode.Text) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.Fields("itemcode") = UCase(Trim(txtitemcode.Text))
    rs.Fields("itemname") = UCase(Trim(txtitemname.Text))
    rs.Fields("tamilname") = Trim(txttamilname.Text)
    rs.Fields("qty_type") = UCase(Trim(cmbqtytype.Text))
    rs.Fields("rate") = Format(Val(txtitemrate.Text), "0.00")
    rs.Update
End If
rs.Close

'If rs1.State = 1 Then rs1.Close
'rs1.Open "select * from tbl_stock where itemcode='" & Trim(txtitemcode.Text) & "'", db, adOpenDynamic, adLockOptimistic
'If Not rs1.EOF Then
'    rs1.Fields("itemcode") = UCase(Trim(txtitemcode.Text))
'    rs1.Fields("itemname") = UCase(Trim(txtitemname.Text))
'    rs1.Update
'End If
'rs1.Close

MsgBox "Successfully Modified", vbInformation, "Sumathi Stores"
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtitemcode.Text = "" Then
    MsgBox "Enter the Item Code Properly...", vbInformation, "Sumathi Stores"
    txtitemcode.SetFocus
ElseIf txtitemname.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Sumathi Stores"
    txtitemname.SetFocus
ElseIf cmbqtytype.Text = "" Then
    MsgBox "Select the Quantity Type Properly...", vbInformation, "Sumathi Stores"
    cmbqtytype.SetFocus
ElseIf txtitemrate.Text = "" Then
    MsgBox "Enter the Item rate Properly...", vbInformation, "Sumathi Stores"
    txtitemrate.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemcode='" & Trim(txtitemcode.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Itemcode is Already Saved. Please Type Different Itemcode", vbInformation, "Sumathi Stores"
    Else
        rs.AddNew
        rs.Fields("itemcode") = UCase(Trim(txtitemcode.Text))
        rs.Fields("itemname") = UCase(Trim(txtitemname.Text))
        rs.Fields("tamilname") = Trim(txttamilname.Text)
        rs.Fields("qty_type") = UCase(Trim(cmbqtytype.Text))
        rs.Fields("rate") = Format(Val(txtitemrate.Text), "0.00")
        rs.Update
        
'        '-------------------Stock Update----------------------
'        If rs1.State = 1 Then rs1.Close
'        rs1.Open "select * from tbl_stock", db, adOpenDynamic, adLockOptimistic
'        rs1.AddNew
'            rs1.Fields("itemcode") = UCase(Trim(txtitemcode.Text))
'            rs1.Fields("itemname") = UCase(Trim(txtitemname.Text))
'            rs1.Fields("qty") = 0
'        rs1.Update
'        rs1.Close
'        '-------------------Stock Update----------------------
        
        MsgBox "Item Saved Successfully", vbInformation, "Sumathi Stores"
    End If
    rs.Close

    Call cmdclear_Click
End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(35, 29, 29)
MSGrid.BackColorBkg = RGB(35, 29, 29)

Call connect
Call Fill

cmbqtytype.AddItem "BAGS"
cmbqtytype.AddItem "BOX"
cmbqtytype.AddItem "GRMS"
cmbqtytype.AddItem "JAR"
cmbqtytype.AddItem "KGS"
cmbqtytype.AddItem "LTS"
cmbqtytype.AddItem "ML"
cmbqtytype.AddItem "NOS"
cmbqtytype.AddItem "PKT"

CmdSave.Enabled = True
CmdModify.Enabled = False
CmdDelete.Enabled = False
End Sub

Private Sub MSGrid_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_itemmaster where itemcode='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtitemcode.Text = rs.Fields("itemcode")
    txtitemname.Text = rs.Fields("itemname")
    txttamilname.Text = IIf(IsNull(rs.Fields("tamilname")), "", rs.Fields("tamilname"))
    cmbqtytype.Text = rs.Fields("qty_type")
    txtitemrate.Text = rs.Fields("rate")
    
    txtitemname.SetFocus
    txtitemname.SelStart = 0
    txtitemname.SelLength = Len(txtitemname.Text)
    
    txtitemcode.Enabled = False
    
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True
End If
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & Format(Val(rs.Fields("rate")), "0.00")
        rs.MoveNext
    Wend
End If
rs.Close
End Function

Private Sub txtitemcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemcode='" & Trim(txtitemcode.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txtitemcode.Text = rs.Fields("itemcode")
        txtitemname.Text = rs.Fields("itemname")
        txttamilname.Text = rs.Fields("tamilname")
        cmbqtytype.Text = rs.Fields("qty_type")
        txtitemrate.Text = rs.Fields("rate")
        txtitemcode.Enabled = False
        
        CmdSave.Enabled = False
        CmdModify.Enabled = True
        CmdDelete.Enabled = True
    Else
        txtitemname.SetFocus
        txtitemname.SelStart = 0
        txtitemname.SelLength = Len(txtitemname.Text)    'select the text
    End If
End If
End Sub

Private Sub txtitemname_Change()
stmt = "select * from tbl_itemmaster where itemname like'" & Trim(txtitemname.Text) & "%' order by itemcode"
If rs1.State = 1 Then rs1.Close
rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs1.EOF Then
    rs1.MoveFirst
    While Not rs1.EOF
        MSGrid.AddItem rs1.Fields("itemcode") & vbTab & rs1.Fields("itemname")
        rs1.MoveNext
    Wend
End If
rs1.Close
End Sub

Private Sub txtitemname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttamilname.SetFocus
    txttamilname.SelStart = 0
    txttamilname.SelLength = Len(txttamilname.Text)
End If
End Sub

Private Sub txttamilname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbqtytype.SetFocus
    cmbqtytype.SelStart = 0
    cmbqtytype.SelLength = Len(cmbqtytype.Text)    'select the text
End If
End Sub

Private Sub cmbqtytype_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    txtitemrate.SetFocus
    txtitemrate.SelStart = 0
    txtitemrate.SelLength = Len(txtitemrate.Text)    'select the text
End If
End Sub

Private Sub txtitemrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub
