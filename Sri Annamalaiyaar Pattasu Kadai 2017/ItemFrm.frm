VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ItemFrm 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Item Details"
   ClientHeight    =   8010
   ClientLeft      =   105
   ClientTop       =   -15
   ClientWidth     =   13320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   20242.51
   ScaleMode       =   0  'User
   ScaleWidth      =   28130.33
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtretailrate 
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
      TabIndex        =   6
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtwholesalerate 
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
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
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
      Left            =   1920
      TabIndex        =   7
      Top             =   7200
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
      Left            =   3960
      TabIndex        =   8
      Top             =   7200
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
      Left            =   8160
      TabIndex        =   10
      Top             =   7200
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
      Left            =   10320
      TabIndex        =   11
      Top             =   7200
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
      Left            =   6000
      TabIndex        =   9
      Top             =   7200
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
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6135
      Left            =   5400
      TabIndex        =   12
      Top             =   720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   33023
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483643
      BackColorBkg    =   16576
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "I Code|Item Name                                             |I Type |Pur Rate       |Retail Rate   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtpurchaserate 
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
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
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
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2520
      TabIndex        =   20
      Top             =   3360
      Width           =   2055
      Begin VB.OptionButton Optpkt 
         BackColor       =   &H000040C0&
         Caption         =   "Packets"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Optbox 
         BackColor       =   &H000040C0&
         Caption         =   "Boxes"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Retail Rate"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Wholesale Rate"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   0
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   15
      Top             =   1560
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Rate"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   2310
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
db.Execute "delete from tbl_itemmaster where itemcode=" & Val(txtitemcode.Text)
db.Execute "delete from tbl_stock where itemcode=" & Val(txtitemcode.Text)
MsgBox "Successfully Deleted...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
Call cmdclear_Click
End Sub

Private Sub CmdModify_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_itemmaster where itemcode=" & Val(txtitemcode.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.Fields("itemcode") = Val(txtitemcode.Text)
    rs.Fields("itemname") = UCase(txtitemname.Text)
    If Optbox.Value = True Then
        rs.Fields("itemtype") = "Boxes"
    Else
        rs.Fields("itemtype") = "Packets"
    End If
    rs.Fields("purchaserate") = Val(Format(txtpurchaserate.Text, "0.00"))
    'rs.Fields("wholesalerate") = Val(Format(txtwholesalerate.Text, "0.00"))
    rs.Fields("retailrate") = Val(Format(txtretailrate.Text, "0.00"))
    rs.Update
End If
rs.Close

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_stock where itemcode=" & Val(txtitemcode.Text), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    rs1.Fields("itemcode") = Val(txtitemcode.Text)
    rs1.Fields("itemname") = UCase(txtitemname.Text)
    If Optbox.Value = True Then
        rs1.Fields("itemtype") = "Boxes"
    Else
        rs1.Fields("itemtype") = "Packets"
    End If
    rs1.Update
End If
rs1.Close

MsgBox "Successfully Modified", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtitemcode.Text = "" Then
    MsgBox "Enter the Item Code Properly...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    txtitemcode.SetFocus
ElseIf txtitemname.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    txtitemname.SetFocus
ElseIf txtpurchaserate.Text = "" Then
    MsgBox "Enter the Purchase Rate Properly...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    txtpurchaserate.SetFocus
'ElseIf txtwholesalerate.Text = "" Then
'    MsgBox "Enter the Wholesale Rate Properly...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
'    txtwholesalerate.SetFocus
ElseIf txtretailrate.Text = "" Then
    MsgBox "Enter the Retail Rate Properly...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    txtretailrate.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemcode=" & txtitemcode.Text, db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Itemcode is Already Saved. Please Type Different Itemcode", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    Else
        rs.AddNew
        rs.Fields("itemcode") = Val(txtitemcode.Text)
        rs.Fields("itemname") = UCase(txtitemname.Text)
        If Optbox.Value = True Then
            rs.Fields("itemtype") = "Boxes"
        Else
            rs.Fields("itemtype") = "Packets"
        End If
        rs.Fields("purchaserate") = Val(Format(txtpurchaserate.Text, "0.00"))
        'rs.Fields("wholesalerate") = Val(Format(txtwholesalerate.Text, "0.00"))
        rs.Fields("retailrate") = Val(Format(txtretailrate.Text, "0.00"))
        rs.Update
        
        '-------------------Stock Update----------------------
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select * from tbl_stock", db, adOpenDynamic, adLockOptimistic
        rs1.AddNew
            rs1.Fields("itemcode") = Val(txtitemcode.Text)
            rs1.Fields("itemname") = UCase(txtitemname.Text)
            If Optbox.Value = True Then
                rs1.Fields("itemtype") = "Boxes"
            Else
                rs1.Fields("itemtype") = "Packets"
            End If
            rs1.Fields("qty") = 0
        rs1.Update
        rs1.Close
        '-------------------Stock Update----------------------
        
        MsgBox "Item Saved Successfully", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    End If
    rs.Close

    Call cmdclear_Click
End If
End Sub

Private Sub Form_Load()
'Me.BackColor = RGB(238, 6, 147)
'Label1.BackColor = RGB(238, 6, 147)
'Label2.BackColor = RGB(238, 6, 147)
'Label3.BackColor = RGB(238, 6, 147)
'Label4.BackColor = RGB(238, 6, 147)
'Label6.BackColor = RGB(238, 6, 147)
'Label7.BackColor = RGB(238, 6, 147)
'Label9.BackColor = RGB(238, 6, 147)
'Frame2.BackColor = RGB(238, 6, 147)
'Optbox.BackColor = RGB(238, 6, 147)
'Optpkt.BackColor = RGB(238, 6, 147)
'MSGrid.BackColorBkg = RGB(238, 6, 147)


Call connect
Call Fill

CmdSave.Enabled = True
CmdModify.Enabled = False
CmdDelete.Enabled = False
End Sub

Private Sub MSGrid_Click()
    txtitemcode.Text = MSGrid.TextMatrix(MSGrid.Row, 0)
    txtitemname.Text = MSGrid.TextMatrix(MSGrid.Row, 1)
    If MSGrid.TextMatrix(MSGrid.Row, 2) = "Boxes" Then
        Optbox.Value = True
        Optpkt.Value = False
    Else
        Optbox.Value = False
        Optpkt.Value = True
    End If
    txtpurchaserate.Text = MSGrid.TextMatrix(MSGrid.Row, 3)
    'txtwholesalerate.Text = MSGrid.TextMatrix(MSGrid.Row, 4)
    txtretailrate.Text = MSGrid.TextMatrix(MSGrid.Row, 4)
    txtitemname.SetFocus
    txtitemname.SelStart = 0
    txtitemname.SelLength = Len(txtitemname.Text)
    txtitemcode.Enabled = False
    
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("itemtype") & vbTab & rs.Fields("purchaserate") & vbTab & rs.Fields("retailrate")
        rs.MoveNext
    Wend
End If
rs.Close
End Function

Private Sub txtitemcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemcode=" & Val(txtitemcode.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txtitemcode.Text = rs.Fields("itemcode")
        txtitemname.Text = rs.Fields("itemname")
        If Trim(rs.Fields("itemtype")) = "Boxes" Then
            Optbox.Value = True
            Optpkt.Value = False
        Else
            Optbox.Value = False
            Optpkt.Value = True
        End If
        txtpurchaserate.Text = rs.Fields("purchaserate")
        'txtwholesalerate.Text = rs.Fields("wholesalerate")
        txtretailrate.Text = rs.Fields("retailrate")
        txtpurchaserate.SetFocus
        txtpurchaserate.SelStart = 0
        txtpurchaserate.SelLength = Len(txtpurchaserate.Text)
        txtitemcode.Enabled = False
        
        CmdSave.Enabled = False
        CmdModify.Enabled = True
        CmdDelete.Enabled = True
    Else
        txtitemname.SetFocus
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
        MSGrid.AddItem rs1.Fields("itemcode") & vbTab & rs1.Fields("itemname") & vbTab & rs1.Fields("itemtype") & vbTab & rs1.Fields("purchaserate") & vbTab & rs1.Fields("retailrate")
        rs1.MoveNext
    Wend
End If
rs1.Close
End Sub

Private Sub txtitemname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Optbox.SetFocus
End If
End Sub

Private Sub Optbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpurchaserate.SetFocus
    txtpurchaserate.SelStart = 0
    txtpurchaserate.SelLength = Len(txtpurchaserate.Text)
End If
End Sub

Private Sub Optpkt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpurchaserate.SetFocus
    txtpurchaserate.SelStart = 0
    txtpurchaserate.SelLength = Len(txtpurchaserate.Text)
End If
End Sub

Private Sub txtpurchaserate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtretailrate.SetFocus
    txtretailrate.SelStart = 0
    txtretailrate.SelLength = Len(txtretailrate.Text)
End If
End Sub

'Private Sub txtwholesalerate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtretailrate.SetFocus
'    txtretailrate.SelStart = 0
'    txtretailrate.SelLength = Len(txtretailrate.Text)
'End If
'End Sub

Private Sub txtretailrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub
