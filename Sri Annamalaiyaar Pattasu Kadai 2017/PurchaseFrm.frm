VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Item Purchase"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   17580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   17580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprevious 
      BackColor       =   &H00C0E0FF&
      Caption         =   "<< <<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C0E0FF&
      Caption         =   ">> >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox List1 
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
      Height          =   2130
      ItemData        =   "PurchaseFrm.frx":0000
      Left            =   5280
      List            =   "PurchaseFrm.frx":0002
      TabIndex        =   25
      Top             =   1245
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtsupname 
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
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtfind 
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
      Left            =   11760
      TabIndex        =   24
      Top             =   0
      Width           =   5775
   End
   Begin VB.TextBox txtpayamt 
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
      Height          =   525
      Left            =   4080
      TabIndex        =   22
      Text            =   "0"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox txtgridtotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox txtquantity 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdcontinue 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   5
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox txtpid 
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
      TabIndex        =   17
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox txtdiscount 
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
      Height          =   525
      Left            =   480
      TabIndex        =   14
      Text            =   "0"
      Top             =   7560
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   9720
      TabIndex        =   2
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   741
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
      Format          =   100204545
      CurrentDate     =   40537
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "S&ave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   6
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7200
      TabIndex        =   4
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9240
      TabIndex        =   3
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00C0E0FF&
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5160
      TabIndex        =   9
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox txtinvno 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   6255
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   33023
      ForeColorFixed  =   16777215
      BackColorBkg    =   16576
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   $"PurchaseFrm.frx":0004
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
   Begin MSFlexGridLib.MSFlexGrid MSGrid1 
      Height          =   8895
      Left            =   11760
      TabIndex        =   18
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15690
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   33023
      ForeColorFixed  =   16777215
      BackColorBkg    =   16576
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "I Code |Item Name                                                |I Type   "
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Payment Amt"
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
      Left            =   3960
      TabIndex        =   23
      Top             =   8040
      Width           =   2100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   7560
      TabIndex        =   21
      Top             =   7560
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2040
      TabIndex        =   15
      Top             =   8040
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Discount"
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
      Left            =   360
      TabIndex        =   13
      Top             =   8040
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   8880
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASED ITEM DETAILS"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   5445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
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
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   1680
   End
End
Attribute VB_Name = "PurchaseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
Unload Me
PurchaseFrm.Show
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
db.Execute "delete from tbl_purchase where pid=" & Val(txtpid.Text)
'Stock update-------------------------------
If rs.State = 1 Then rs.Close
rs.Open "select itemcode from tbl_itemmaster order by itemcode", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select sum(quantity) from tbl_purchase where itemcode=" & Val(rs.Fields("itemcode")) & " and purchasedate=#" & DTPicker1.Value & "#", db, adOpenDynamic, adLockOptimistic
    quantity = rs1.Fields(0)

    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from tbl_stock where itemcode=" & rs.Fields("itemcode"), db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        rs1.Fields("qty") = Val(quantity)
        rs1.Update
    End If

    rs.MoveNext
Wend
'Stock update------------------------------

MsgBox "Successfully Deleted...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
Call cmdclear_Click
End Sub

Private Sub cmdcontinue_Click()
MSGrid.Row = MSGrid.Rows - 1
MSGrid.Col = 0
MSGrid.SetFocus
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdcontinue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 80 Then     ' p key
    txtpayamt.SetFocus
    txtpayamt.SelStart = 0
    txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
End If

If KeyCode = 68 Then     ' d key
    txtdiscount.SetFocus
    txtdiscount.SelStart = 0
    txtdiscount.SelLength = Len(txtdiscount.Text)    'select the text
End If

If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub cmdnext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtpid.Text = ""
    txtinvno.Text = ""
    txtsupname.Text = ""
    txtgridtotamt.Text = ""
    txtquantity.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""

    txtpid.Text = rs.Fields("pid")
    txtinvno.Text = rs.Fields("invoiceno")
    txtsupname.Text = rs.Fields("supname")
    List1.Visible = False
    DTPicker1.Value = rs.Fields("purchasedate")
    txtquantity.Text = Format(rs.Fields("quantity"), "0.00")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtdiscount.Text = Format(rs.Fields("discount"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemtype")
        MSGrid.TextMatrix(i, 3) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 4) = rs.Fields("quantity")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    cmdcontinue.Enabled = True
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    cmdclear.Enabled = True

    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
Else
    Call cmdclear_Click
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
End If
End Sub

Private Sub cmdprevious_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtpid.Text = ""
    txtinvno.Text = ""
    txtsupname.Text = ""
    txtgridtotamt.Text = ""
    txtquantity.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""

    txtpid.Text = rs.Fields("pid")
    txtinvno.Text = rs.Fields("invoiceno")
    txtsupname.Text = rs.Fields("supname")
    List1.Visible = False
    DTPicker1.Value = rs.Fields("purchasedate")
    txtquantity.Text = Format(rs.Fields("quantity"), "0.00")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtdiscount.Text = Format(rs.Fields("discount"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemtype")
        MSGrid.TextMatrix(i, 3) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 4) = rs.Fields("quantity")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    cmdcontinue.Enabled = True
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    cmdclear.Enabled = True

    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
Else
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
End If
rs.Close
End Sub

Private Sub CmdSave_Click()
If txtinvno.Text = "" Then
    MsgBox "Enter the invoice no properly...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
    txtinvno.SetFocus
Else
    '================Supplier Update====================
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_suppliermaster where suppliername='" & txtsupname.Text & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        s = rs.Fields("sid")
    Else
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select * from tbl_suppliermaster order by sid", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            rs1.MoveLast
            s = rs1.Fields("sid")
            s = s + 1
        End If

        rs1.AddNew
            rs1.Fields("sid") = s
            rs1.Fields("suppliername") = UCase(txtsupname.Text)
        rs1.Update
        rs1.Close
    End If
    rs.Close
    '================Supplier Update====================

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then     '  If the record is allready stored means we delete it and then update it
        'Stock update-------------------------------
        For i = 1 To MSGrid.Rows - 2
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_stock where itemcode=" & MSGrid.TextMatrix(i, 0), db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                rs.Fields("qty") = Format(Val(rs.Fields("qty")) - Val(MSGrid.TextMatrix(i, 4)), "0.00")
                rs.Update
            End If
            rs.Close
        Next i
        'Stock update-------------------------------

        db.Execute "delete from tbl_purchase where pid=" & Val(txtpid.Text)
    End If

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchase", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        rs.AddNew
        rs.Fields("pid") = Val(txtpid.Text)
        rs.Fields("invoiceno") = txtinvno.Text
        rs.Fields("sid") = s
        rs.Fields("supname") = txtsupname.Text
        rs.Fields("purchasedate") = DTPicker1.Value
        rs.Fields("itemcode") = Val(MSGrid.TextMatrix(i, 0))
        rs.Fields("itemname") = MSGrid.TextMatrix(i, 1)
        rs.Fields("itemtype") = MSGrid.TextMatrix(i, 2)
        rs.Fields("itemrate") = Val(MSGrid.TextMatrix(i, 3))
        rs.Fields("quantity") = Val(MSGrid.TextMatrix(i, 4))
        rs.Fields("itemamt") = Val(MSGrid.TextMatrix(i, 5))
        rs.Fields("gridtotamt") = Val(txtgridtotamt.Text)
        rs.Fields("discount") = Val(txtdiscount.Text)
        rs.Fields("totamt") = Val(txttotamt.Text)
        rs.Fields("payamt") = Val(txtpayamt.Text)
        rs.Update
    Next i
    rs.Close

    'Stock update-------------------------------
    For i = 1 To MSGrid.Rows - 2
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select * from tbl_stock where itemcode=" & MSGrid.TextMatrix(i, 0), db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_stock where itemcode=" & MSGrid.TextMatrix(i, 0), db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                rs.Fields("qty") = Val(MSGrid.TextMatrix(i, 4)) + Val(rs.Fields("qty"))
                rs.Update
            End If
        End If
        rs1.Close
    Next i
    'Stock update------------------------------

    MsgBox "Successfully Saved...", vbInformation, "Sri Annamalaiyar Pattasu Kadai"
End If
Call cmdclear_Click
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 80 Then     ' p key
    txtpayamt.SetFocus
    txtpayamt.SelStart = 0
    txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
End If

If KeyCode = 68 Then     ' d key
    txtdiscount.SetFocus
    txtdiscount.SelStart = 0
    txtdiscount.SelLength = Len(txtdiscount.Text)    'select the text
End If

If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub MSGrid1_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_itemmaster where itemcode=" & MSGrid1.TextMatrix(MSGrid1.Row, 0), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    MSGrid.TextMatrix(MSGrid.Row, 0) = rs1.Fields("itemcode")
    MSGrid.TextMatrix(MSGrid.Row, 1) = rs1.Fields("itemname")
    MSGrid.TextMatrix(MSGrid.Row, 2) = rs1.Fields("itemtype")
    MSGrid.TextMatrix(MSGrid.Row, 3) = Format(Val(rs1.Fields("purchaserate")), "0.00")
    MSGrid.Col = MSGrid.Col + 4  ' Grid entry was changed to quantity coloumn
    MSGrid.SetFocus
End If
rs1.Close
End Sub

Private Sub txtdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub txtdiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttotamt.Text = Format(Val(txtgridtotamt.Text) - Val(txtdiscount.Text), "0.00")
    txtpayamt.SetFocus
End If
End Sub

Private Sub txtinvno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub txtinvno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsupname.SetFocus
End If
End Sub

Private Sub txtpayamt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub txtsupname_Change()
List1.Visible = True

stmt = "select suppliername from tbl_suppliermaster where suppliername like'" & Trim(txtsupname.Text) & "%' order by suppliername"
If rs1.State = 1 Then rs1.Close
rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
List1.Clear
If Not rs1.EOF Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        List1.AddItem rs1.Fields("suppliername")
        rs1.MoveNext
    Loop
End If
rs1.Close

If Not List1.ListCount = 0 Then
    List1.ListIndex = 0
End If
End Sub

Private Sub txtsupname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub txtsupname_LostFocus()
List1.Visible = False
End Sub

Private Sub txtsupname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not List1.ListCount = 0 Then
        txtsupname.Text = List1.List(List1.ListIndex)
    End If
    DTPicker1.SetFocus
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.Col = 0
    MSGrid.Row = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub Form_Load()
'Me.BackColor = RGB(238, 6, 147)
'Label1.BackColor = RGB(238, 6, 147)
'Label2.BackColor = RGB(238, 6, 147)
'Label3.BackColor = RGB(238, 6, 147)
'Label4.BackColor = RGB(238, 6, 147)
'Label5.BackColor = RGB(238, 6, 147)
'Label6.BackColor = RGB(238, 6, 147)
'Label7.BackColor = RGB(238, 6, 147)
'Label8.BackColor = RGB(238, 6, 147)
'MSGrid.BackColorBkg = RGB(238, 6, 147)
'MSGrid1.BackColorBkg = RGB(238, 6, 147)

Call connect

DTPicker1.Value = Date

If rs.State = 1 Then rs.Close
rs.Open "select suppliername from tbl_suppliermaster order by suppliername", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    List1.AddItem rs.Fields("suppliername")
    rs.MoveNext
Wend
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchase order by pid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtpid.Text = Val(rs.Fields("pid")) + 1
Else
    txtpid.Text = 1
End If
rs.Close

MSGrid.Row = 1
MSGrid.Col = 0
MSGrid.CellBackColor = RGB(117, 145, 233)

Call Fill

CmdSave.Enabled = True
CmdDelete.Enabled = False
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid1.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("itemtype")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub MSGrid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 112 Then 'F1 Key
    If Not MSGrid.Rows = 1 Then
        MSGrid.CellBackColor = vbWhite
        txtfind.SetFocus
    End If
End If

If KeyCode = 117 Then 'F6 Key for Delete the row
    If Not MSGrid.Rows = 1 Then
        txtquantity.Text = Format(Val(txtquantity.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 4)), "0.00")
        txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 5)), "0.00")
        txttotamt.Text = Format(Val(txtgridtotamt.Text), "0.00")

        MSGrid.Row = MSGrid.Row
        MSGrid.Col = 0
        If MSGrid.Row = 1 Then
            MSGrid.TextMatrix(1, 0) = ""
            MSGrid.TextMatrix(1, 1) = ""
            MSGrid.TextMatrix(1, 2) = ""
            MSGrid.TextMatrix(1, 3) = ""
            MSGrid.TextMatrix(1, 4) = ""
            MSGrid.TextMatrix(1, 5) = ""
        Else
            MSGrid.RemoveItem MSGrid.Row
        End If
        MSGrid.CellBackColor = RGB(117, 145, 233)
    End If
End If

If KeyCode = 118 Then 'F7 Key
    txtinvno.SetFocus
    txtinvno.SelStart = 0
    txtinvno.SelLength = Len(txtinvno.Text)    'select the text
End If

If KeyCode = 119 Then 'F8 Key
    txtsupname.SetFocus
    txtsupname.SelStart = 0
    txtsupname.SelLength = Len(txtsupname.Text)
End If

If KeyCode = 120 Then 'F9 Key
    DTPicker1.SetFocus
End If

'If KeyCode = 122 Then 'F11 Key
'    MSGrid.Row = MSGrid.Row
'    MSGrid.Col = 3                 ' Navigate to Item Rate Coloumn in Grid
'    MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
'    MSGrid.SetFocus
'End If

If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If

End Sub

Private Sub MSGrid_KeyPress(KeyAscii As Integer)

If MSGrid.Col = 0 Or MSGrid.Col = 4 Then      ' (MSGrid.Col = 3 is Itemrate) Itemcode and Quantity gridy coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
'    Case 46         ' 46 keyascii is for dot symbol
'        If MSGrid.Col = 3 Then   'For Itemrate Coloumn
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
'        If MSGrid.Col = 0 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
'    Case 97 To 122  ' 97-122 keyascii is for small a to z
'        If MSGrid.Col = 0 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 0 Then      ' Item Code
            If MSGrid.TextMatrix(MSGrid.Row, 0) <> "" Then
                If rs.State = 1 Then rs.Close
                rs.Open "select * from tbl_itemmaster where itemcode=" & MSGrid.TextMatrix(MSGrid.Row, 0), db, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    MSGrid.TextMatrix(MSGrid.Row, 0) = rs.Fields("itemcode")
                    MSGrid.TextMatrix(MSGrid.Row, 1) = rs.Fields("itemname")
                    MSGrid.TextMatrix(MSGrid.Row, 2) = rs.Fields("itemtype")
                    MSGrid.TextMatrix(MSGrid.Row, 3) = Format(rs.Fields("purchaserate"), "0.00")
                    MSGrid.Col = MSGrid.Col + 4  ' Grid entry was changed to quantity coloumn
                Else
                    MSGrid.TextMatrix(MSGrid.Row, 0) = ""
                    MSGrid.Row = MSGrid.Row
                    MSGrid.Col = 0
                End If
                rs.Close
            Else
                MSGrid.CellBackColor = vbWhite
                cmdcontinue.SetFocus   'cursor navigation to the Continue Button
            End If
        End If

'        If MSGrid.Col = 3 Then      ' Item Rate
'            If MSGrid.TextMatrix(MSGrid.Row, 3) <> "" Then
'                MSGrid.Row = MSGrid.Row     'cursor position maintains the same row
'                MSGrid.Col = 5              'cursor position changed to the quantity coloumn of that same row
'                MSGrid.SetFocus
'            End If
'        End If

        If MSGrid.Col = 4 Then  '   Quantity
            If MSGrid.TextMatrix(MSGrid.Row, 4) <> "" Then
                MSGrid.TextMatrix(MSGrid.Row, 5) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 3)) * Val(MSGrid.TextMatrix(MSGrid.Row, 4)), "0.00")
                txtgridtotamt.Text = 0
                txtquantity.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Grid Total bill amount calculation
                    txttotamt.Text = Format(Val(txtgridtotamt.Text), "0.00")
                    txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
                    txtquantity.Text = Val(txtquantity.Text) + Val(MSGrid.TextMatrix(i, 4))   'Total quantity calculation
                Next i
                If MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = "" Then
                    MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                End If
                MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                MSGrid.Row = MSGrid.Row + 1     'cursor position changed to the newlly created row
                MSGrid.Col = 0                  'cursor position changed to the first coloumn of that newly created row
            End If
        End If
    End Select
End If
End Sub

Private Sub MSGrid_EnterCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub MSGrid_LeaveCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
MSGrid.CellBackColor = vbWhite
End Sub

Private Sub MSGrid1_EnterCell()
MSGrid1.Row = MSGrid1.Row
MSGrid1.Col = MSGrid1.Col
MSGrid1.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub MSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then 'F1 Key
    MSGrid1.CellBackColor = vbWhite
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
    MSGrid.CellBackColor = RGB(117, 145, 233)
End If
End Sub

Private Sub MSGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MSGrid1.CellBackColor = vbWhite

    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from tbl_itemmaster where itemcode=" & MSGrid1.TextMatrix(MSGrid1.Row, 0), db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        MSGrid.TextMatrix(MSGrid.Row, 0) = rs1.Fields("itemcode")
        MSGrid.TextMatrix(MSGrid.Row, 1) = rs1.Fields("itemname")
        MSGrid.TextMatrix(MSGrid.Row, 2) = rs1.Fields("itemtype")
        MSGrid.TextMatrix(MSGrid.Row, 3) = Format(Val(rs1.Fields("purchaserate")), "0.00")
        MSGrid.Col = MSGrid.Col + 4  ' Grid entry was changed to quantity coloumn
        MSGrid.SetFocus
    End If
    rs1.Close
End If
End Sub

Private Sub MSGrid1_LeaveCell()
MSGrid1.Row = MSGrid1.Row
MSGrid1.Col = MSGrid1.Col
MSGrid1.CellBackColor = vbWhite
End Sub

Private Sub txtbillno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsupname.SetFocus
End If
End Sub

Private Sub txtfind_Change()
stmt = "select * from tbl_itemmaster where itemname like'" & Trim(txtfind.Text) & "%' order by itemcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid1.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("itemtype")
        rs.MoveNext
    Loop
End If
rs.Close
End Sub

Private Sub txtfind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then 'F1 Key
    txtfind.BackColor = vbWhite
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
    MSGrid.CellBackColor = RGB(117, 145, 233)
End If
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfind.BackColor = vbWhite
    MSGrid1.Row = 1
    MSGrid1.Col = 1
    MSGrid1.SetFocus
    MSGrid1.CellBackColor = RGB(117, 145, 233)
End If
End Sub

Private Sub txtpayamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdSave.SetFocus
End If
End Sub

Private Sub txtfind_GotFocus()
    txtfind.BackColor = RGB(117, 145, 233)
End Sub

Private Sub txtfind_LostFocus()
    txtfind.BackColor = vbWhite
End Sub

Private Sub cmdcontinue_GotFocus()
    cmdcontinue.BackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdcontinue_LostFocus()
    cmdcontinue.BackColor = RGB(239, 234, 219)
End Sub

Private Sub CmdSave_GotFocus()
    CmdSave.BackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdSave_LostFocus()
    CmdSave.BackColor = RGB(239, 234, 219)
End Sub

Private Sub cmdclear_GotFocus()
    cmdclear.BackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdclear_LostFocus()
    cmdclear.BackColor = RGB(239, 234, 219)
End Sub

Private Sub cmddelete_GotFocus()
    CmdDelete.BackColor = RGB(117, 145, 233)
End Sub

Private Sub cmddelete_LostFocus()
    CmdDelete.BackColor = RGB(239, 234, 219)
End Sub

Private Sub CmdClose_GotFocus()
    CmdClose.BackColor = RGB(117, 145, 233)
End Sub

Private Sub CmdClose_LostFocus()
    CmdClose.BackColor = RGB(239, 234, 219)
End Sub
