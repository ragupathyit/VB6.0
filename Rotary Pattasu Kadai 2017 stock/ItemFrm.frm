VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ItemFrm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Item Details"
   ClientHeight    =   8280
   ClientLeft      =   105
   ClientTop       =   -15
   ClientWidth     =   14040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   20924.84
   ScaleMode       =   0  'User
   ScaleWidth      =   29650.89
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsalesrate 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
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
      Left            =   2400
      TabIndex        =   5
      Top             =   7680
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
      Left            =   4320
      TabIndex        =   6
      Top             =   7680
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
      Left            =   8400
      TabIndex        =   8
      Top             =   7680
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
      Left            =   10440
      TabIndex        =   9
      Top             =   7680
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
      Left            =   6240
      TabIndex        =   7
      Top             =   7680
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
      Left            =   2400
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6375
      Left            =   5400
      TabIndex        =   10
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16761024
      BackColorBkg    =   16744576
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Item Code |Item Name                                                 |Item Type |Quantity|Sales Rate       "
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
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   3960
      Width           =   2895
      Begin VB.OptionButton Optpkt 
         BackColor       =   &H00FF8080&
         Caption         =   "Packets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton Optbox 
         BackColor       =   &H00FF8080&
         Caption         =   "Boxes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox txtquantity 
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
      Left            =   2400
      TabIndex        =   17
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   7320
      Width           =   14175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM DETAILS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6000
      TabIndex        =   13
      Top             =   240
      Width           =   2460
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   14175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Per Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      TabIndex        =   19
      Top             =   6000
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   18
      Top             =   5040
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Width           =   1200
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
MsgBox "Successfully Deleted...", vbInformation, "Rotary Club of Mettupalayam"
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
    rs.Fields("quantity") = Val(txtquantity.Text)
    rs.Fields("salesrate") = Val(Format(txtsalesrate.Text, "0.00"))
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
    rs1.Fields("qty") = Val(txtquantity.Text)
    rs1.Update
End If
rs1.Close

MsgBox "Successfully Modified", vbInformation, "Rotary Club of Mettupalayam"
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtitemcode.Text = "" Then
    MsgBox "Enter the Item Code Properly...", vbInformation, "Rotary Club of Mettupalayam"
    txtitemcode.SetFocus
ElseIf txtitemname.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Rotary Club of Mettupalayam"
    txtitemname.SetFocus
ElseIf txtquantity.Text = "" Then
    MsgBox "Enter the Quantity Properly...", vbInformation, "Rotary Club of Mettupalayam"
    txtquantity.SetFocus
ElseIf txtsalesrate.Text = "" Then
    MsgBox "Enter the Sales Rate Properly...", vbInformation, "Rotary Club of Mettupalayam"
    txtsalesrate.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemcode=" & Val(txtitemcode.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Itemcode is Already Saved. Please Type Different Itemcode", vbInformation, "Rotary Club of Mettupalayam"
    Else
        rs.AddNew
        rs.Fields("itemcode") = Val(txtitemcode.Text)
        rs.Fields("itemname") = UCase(txtitemname.Text)
        If Optbox.Value = True Then
            rs.Fields("itemtype") = "Boxes"
        Else
            rs.Fields("itemtype") = "Packets"
        End If
        rs.Fields("quantity") = Val(txtquantity.Text)
        rs.Fields("salesrate") = Val(Format(txtsalesrate.Text, "0.00"))
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
            rs1.Fields("qty") = Val(txtquantity.Text)
        rs1.Update
        rs1.Close
        '-------------------Stock Update----------------------
        
        MsgBox "Item Saved Successfully", vbInformation, "Rotary Club of Mettupalayam"
    End If
    rs.Close

    Call cmdclear_Click
End If
End Sub

Private Sub Form_Load()

'Me.BackColor = RGB(255, 204, 203)
'MSGrid.BackColorBkg = RGB(255, 204, 203)
'Frame2.BackColor = RGB(255, 204, 203)
'Optbox.BackColor = RGB(255, 204, 203)
'Optpkt.BackColor = RGB(255, 204, 203)

Call connect
Call Fill

CmdSave.Enabled = True
CmdModify.Enabled = False
cmddelete.Enabled = False
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
txtquantity.Text = MSGrid.TextMatrix(MSGrid.Row, 3)
txtsalesrate.Text = MSGrid.TextMatrix(MSGrid.Row, 4)
txtitemname.SetFocus
txtitemname.SelStart = 0
txtitemname.SelLength = Len(txtitemname.Text)
txtitemcode.Enabled = False

CmdSave.Enabled = False
CmdModify.Enabled = True
cmddelete.Enabled = True
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("itemtype") & vbTab & rs.Fields("quantity") & vbTab & rs.Fields("salesrate")
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
        txtquantity.Text = rs.Fields("quantity")
        txtsalesrate.Text = rs.Fields("salesrate")
        txtitemcode.Enabled = False

        CmdSave.Enabled = False
        CmdModify.Enabled = True
        cmddelete.Enabled = True
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
        MSGrid.AddItem rs1.Fields("itemcode") & vbTab & rs1.Fields("itemname") & vbTab & rs1.Fields("itemtype") & vbTab & rs1.Fields("quantity") & vbTab & rs1.Fields("salesrate")
        rs1.MoveNext
    Wend
End If
rs1.Close
End Sub

Private Sub txtitemname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub txtitemname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Optbox.SetFocus
End If
End Sub

Private Sub Optbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtquantity.SetFocus
    txtquantity.SelStart = 0
    txtquantity.SelLength = Len(txtquantity.Text)
End If
End Sub

Private Sub Optpkt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtquantity.SetFocus
    txtquantity.SelStart = 0
    txtquantity.SelLength = Len(txtquantity.Text)
End If
End Sub

Private Sub txtquantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsalesrate.SetFocus
    txtsalesrate.SelStart = 0
    txtsalesrate.SelLength = Len(txtsalesrate.Text)
End If
End Sub

Private Sub txtsalesrate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub
