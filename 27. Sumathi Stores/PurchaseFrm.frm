VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Item Purchase"
   ClientHeight    =   9210
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   18015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   18015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlorryhire 
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
      Height          =   405
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   6850
      Width           =   1575
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
      Height          =   405
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0"
      Top             =   7995
      Width           =   1575
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
      Height          =   405
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0"
      Top             =   7230
      Width           =   1575
   End
   Begin VB.TextBox txtpayamt 
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
      Height          =   405
      Left            =   10320
      TabIndex        =   26
      Text            =   "0"
      Top             =   8370
      Width           =   1575
   End
   Begin VB.TextBox txtobalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   10320
      TabIndex        =   25
      Text            =   "0"
      Top             =   7605
      Width           =   1575
   End
   Begin VB.TextBox txtcooly 
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
      Height          =   405
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtbalamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   10320
      TabIndex        =   23
      Text            =   "0"
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox txtgridtotqty 
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
      Height          =   315
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton CmdModify 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Modify"
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
      Left            =   4800
      TabIndex        =   18
      Top             =   7440
      Width           =   1215
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
      TabIndex        =   15
      Top             =   1245
      Visible         =   0   'False
      Width           =   3375
   End
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
      Left            =   1800
      TabIndex        =   17
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
      Left            =   9600
      TabIndex        =   16
      Top             =   240
      Width           =   975
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
      Left            =   12240
      TabIndex        =   14
      Top             =   0
      Width           =   5775
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
      Left            =   1080
      TabIndex        =   5
      Top             =   7440
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
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   7440
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
      Left            =   3000
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
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
      Left            =   4800
      TabIndex        =   3
      Top             =   8160
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
      Left            =   6600
      TabIndex        =   8
      Top             =   7440
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
   Begin MSFlexGridLib.MSFlexGrid MSGrid1 
      Height          =   8895
      Left            =   12240
      TabIndex        =   13
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15690
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "I Code |Item Name                                                             "
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
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   97452035
      CurrentDate     =   42430
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5175
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
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
   Begin VB.Label lbl_lastvdate 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      Left            =   7440
      TabIndex        =   37
      Top             =   8400
      Width           =   300
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Lorry Hire"
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
      Left            =   8760
      TabIndex        =   36
      Top             =   6960
      Width           =   1140
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Total Amt"
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
      Left            =   8760
      TabIndex        =   34
      Top             =   8040
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Grid Total"
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
      Left            =   8760
      TabIndex        =   33
      Top             =   7320
      Width           =   1110
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "&Payment Amt"
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
      Left            =   8760
      TabIndex        =   32
      Top             =   8400
      Width           =   1515
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Old Balance"
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
      Left            =   8760
      TabIndex        =   31
      Top             =   7680
      Width           =   1380
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Cooly"
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
      Left            =   8760
      TabIndex        =   30
      Top             =   6600
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Balance"
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
      Left            =   8760
      TabIndex        =   29
      Top             =   8760
      Width           =   930
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Total Quantity"
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
      Left            =   5760
      TabIndex        =   22
      Top             =   6480
      Width           =   1590
   End
   Begin VB.Label lbldeleted 
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE DELETED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   480
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      Left            =   8880
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   5445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   0
      TabIndex        =   7
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
If Format(Date, "DD/MM/YYYY") = Format(DTPicker1.Value, "DD/MM/YYYY") Then
    db.Execute "update tbl_purchase set isdel=false where pid=" & Val(txtpid.Text)
    db.Execute "update tbl_purchasebalance set isdel=false where pid=" & Val(txtpid.Text)
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchasevoucher where pid=" & Val(txtpid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.Fields("pid") = ""
        rs.Update
    End If
    rs.Close
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then     '  If the record is allready stored means we delete it and then update it
        'Stock update-------------------------------
        While Not rs.EOF
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select * from tbl_stock where itemcode=" & Val(Trim(rs.Fields("itemcode"))), db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                rs1.Fields("qty") = Format(Val(rs1.Fields("qty")) - Val(rs.Fields("quantity")), "0.00")
                rs1.Update
            End If
            rs1.Close
            rs.MoveNext
        Wend
        'Stock update-------------------------------
    End If
        
    MsgBox "Purchase Item is Canceled Successfully", vbInformation, "Sumathi Stores"
Else
    MsgBox "Purchase Item is Not Canceled. Contact Software Developer", vbInformation, "Sumathi Stores"
End If
Call cmdclear_Click
    
End Sub

Private Sub cmdcontinue_Click()
MSGrid.Row = MSGrid.Rows - 1
MSGrid.Col = 0
MSGrid.SetFocus
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdcontinue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub CmdModify_Click()
If Format(Date, "DD/MM/YYYY") = Format(DTPicker1.Value, "DD/MM/YYYY") Then
    If txtinvno.Text = "" Then
        MsgBox "Enter the invoice no properly...", vbInformation, "Sumathi Stores"
        txtinvno.SetFocus
    ElseIf txtsupname.Text = "" Then
        MsgBox "Enter the Supplier Name Properly...", vbInformation, "Sumathi Stores"
        txtsupname.SetFocus
    Else
    '    '================Supplier Update====================
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_suppliermaster where suppliername='" & txtsupname.Text & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            s = rs.Fields("sid")
    '    Else
    '        If rs1.State = 1 Then rs1.Close
    '        rs1.Open "select * from tbl_suppliermaster order by sid", db, adOpenDynamic, adLockOptimistic
    '        If Not rs.EOF Then
    '            rs1.MoveLast
    '            s = rs1.Fields("sid")
    '            s = s + 1
    '        End If
    '
    '        rs1.AddNew
    '            rs1.Fields("sid") = s
    '            rs1.Fields("suppliername") = UCase(txtsupname.Text)
    '        rs1.Update
    '        rs1.Close
        End If
        rs.Close
    '    '================Supplier Update====================
    
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text), db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then     '  If the record is allready stored means we delete it and then update it
            'Stock update-------------------------------
            While Not rs.EOF
                If rs1.State = 1 Then rs1.Close
                rs1.Open "select * from tbl_stock where itemcode=" & Val(Trim(rs.Fields("itemcode"))), db, adOpenDynamic, adLockOptimistic
                If Not rs1.EOF Then
                    rs1.Fields("qty") = Format(Val(rs1.Fields("qty")) - Val(rs.Fields("quantity")), "0.00")
                    rs1.Update
                End If
                rs1.Close
                rs.MoveNext
            Wend
            'Stock update-------------------------------
    
            db.Execute "delete from tbl_purchase where pid=" & Val(txtpid.Text)
            db.Execute "delete from tbl_purchasebalance where pid=" & Val(txtpid.Text)
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_purchasevoucher where pid=" & Val(txtpid.Text) & "", db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                rs.Fields("pid") = ""
                rs.Update
            End If
            rs.Close
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
            rs.Fields("bags") = Val(MSGrid.TextMatrix(i, 0))
            rs.Fields("itemcode") = Val(MSGrid.TextMatrix(i, 1))
            rs.Fields("itemname") = MSGrid.TextMatrix(i, 2)
            rs.Fields("quantity") = MSGrid.TextMatrix(i, 3)
            rs.Fields("itemrate") = Format(Val(MSGrid.TextMatrix(i, 4)), "0.00")
            rs.Fields("itemamt") = Format(Val(MSGrid.TextMatrix(i, 5)), "0.00")
            rs.Fields("cooly") = Format(Val(txtcooly.Text), "0.00")
            rs.Fields("lorryhire") = Format(Val(txtlorryhire.Text), "0.00")
            rs.Fields("gridtotqty") = Format(Val(txtgridtotqty.Text), "0.00")
            rs.Fields("gridtotamt") = Format(Val(txtgridtotamt.Text), "0.00")
            rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
            rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
            rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
            rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
            rs.Update
        Next i
        rs.Close
    
        '====================Purchase Balance====================
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_purchasebalance", db, adOpenDynamic, adLockOptimistic
        rs.AddNew
            rs.Fields("pid") = Val(txtpid.Text)
            rs.Fields("purdate") = DTPicker1.Value
            rs.Fields("sid") = Val(s)
            rs.Fields("supname") = Trim(txtsupname.Text)
            rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
            rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
            rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
            rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
            rs.Fields("baldesc") = txttotamt.Text & "-" & txtpayamt.Text
        rs.Update
        rs.Close
        '====================Purchase Balance====================
            
        db.Execute "update tbl_purchasevoucher set pid=" & Val(txtpid.Text) & " where sid=" & Val(s)
        
        'Stock update-------------------------------
        For i = 1 To MSGrid.Rows - 1
            If MSGrid.TextMatrix(i, 0) <> "" Then
                If rs.State = 1 Then rs.Close
                rs.Open "select * from tbl_stock where itemcode=" & MSGrid.TextMatrix(i, 1), db, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    rs.Fields("qty") = Val(rs.Fields("qty")) + Val(MSGrid.TextMatrix(i, 3))
                    rs.Update
                End If
            End If
        Next i
        'Stock update------------------------------
    
        MsgBox "Successfully Modified...", vbInformation, "Sumathi Stores"
    End If
Else
    MsgBox "Not Modified. Contact Software Developer...", vbInformation, "Sumathi Stores"
End If
Call cmdclear_Click
End Sub

Private Sub cmdnext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lbldeleted.Visible = False
    txtpid.Text = ""
    txtinvno.Text = ""
    txtsupname.Text = ""
    txtcooly.Text = ""
    txtlorryhire.Text = ""
    txtgridtotamt.Text = ""
    txtgridtotqty.Text = ""
    txtobalance.Text = ""
    txttotamt.Text = ""
    txtpayamt.Text = ""
    txtbalamt.Text = ""
    
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""

    If rs.Fields("isdel") = False Then
        txtpid.Text = rs.Fields("pid")
        lbldeleted.Visible = True
        GoTo s:
    Else
        txtpid.Text = rs.Fields("pid")
    End If
    
    txtpid.Text = rs.Fields("pid")
    txtinvno.Text = rs.Fields("invoiceno")
    txtsupname.Text = rs.Fields("supname")
    DTPicker1.Value = Format(rs.Fields("purchasedate"), "DD/MM/YYYY")
    List1.Visible = False
    DTPicker1.Value = rs.Fields("purchasedate")
    txtcooly.Text = Format(rs.Fields("cooly"), "0.00")
    txtlorryhire.Text = Format(rs.Fields("lorryhire"), "0.00")
    txtgridtotqty.Text = Format(rs.Fields("gridtotqty"), "0.00")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtobalance.Text = Format(rs.Fields("obalance"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    txtbalamt.Text = Format(rs.Fields("balamt"), "0.00")
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("bags")
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 3) = rs.Fields("quantity")
        MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True

    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
Else
    Call cmdclear_Click
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
End If

s:
    List1.Visible = False
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True
End Sub

Private Sub cmdprevious_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lbldeleted.Visible = False
    txtpid.Text = ""
    txtinvno.Text = ""
    txtsupname.Text = ""
    txtcooly.Text = ""
    txtlorryhire.Text = ""
    txtgridtotamt.Text = ""
    txtgridtotqty.Text = ""
    txtobalance.Text = ""
    txttotamt.Text = ""
    txtpayamt.Text = ""
    txtbalamt.Text = ""
    
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
    MSGrid.TextMatrix(1, 5) = ""
    
    If rs.Fields("isdel") = False Then
        txtpid.Text = rs.Fields("pid")
        lbldeleted.Visible = True
        GoTo s:
    Else
        txtpid.Text = rs.Fields("pid")
    End If
    
    txtpid.Text = rs.Fields("pid")
    txtinvno.Text = rs.Fields("invoiceno")
    txtsupname.Text = rs.Fields("supname")
    DTPicker1.Value = Format(rs.Fields("purchasedate"), "DD/MM/YYYY")
    List1.Visible = False
    txtcooly.Text = Format(rs.Fields("cooly"), "0.00")
    txtlorryhire.Text = Format(rs.Fields("lorryhire"), "0.00")
    txtgridtotqty.Text = Format(rs.Fields("gridtotqty"), "0.00")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtobalance.Text = Format(rs.Fields("obalance"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    txtbalamt.Text = Format(rs.Fields("balamt"), "0.00")
    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("bags")
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemcode")
        MSGrid.TextMatrix(i, 2) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 3) = rs.Fields("quantity")
        MSGrid.TextMatrix(i, 4) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("itemamt"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True

    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
Else
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
End If

s:
    List1.Visible = False
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True
End Sub

Private Sub CmdSave_Click()
If txtinvno.Text = "" Then
    MsgBox "Enter the invoice no properly...", vbInformation, "Sumathi Stores"
    txtinvno.SetFocus
ElseIf txtsupname.Text = "" Then
    MsgBox "Enter the Supplier Name Properly...", vbInformation, "Sumathi Stores"
    txtsupname.SetFocus
Else
'    '================Supplier Update====================
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_suppliermaster where suppliername='" & txtsupname.Text & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        s = rs.Fields("sid")
'    Else
'        If rs1.State = 1 Then rs1.Close
'        rs1.Open "select * from tbl_suppliermaster order by sid", db, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
'            rs1.MoveLast
'            s = rs1.Fields("sid")
'            s = s + 1
'        End If
'
'        rs1.AddNew
'            rs1.Fields("sid") = s
'            rs1.Fields("suppliername") = UCase(txtsupname.Text)
'        rs1.Update
'        rs1.Close
    End If
    rs.Close
'    '================Supplier Update====================

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchase where pid=" & Val(txtpid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then     '  If the record is allready stored means we delete it and then update it
        'Stock update-------------------------------
        While Not rs.EOF
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select * from tbl_stock where itemcode=" & Val(Trim(rs.Fields("itemcode"))), db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                rs1.Fields("qty") = Format(Val(rs1.Fields("qty")) - Val(rs.Fields("quantity")), "0.00")
                rs1.Update
            End If
            rs1.Close
            rs.MoveNext
        Wend
        'Stock update-------------------------------

        db.Execute "delete from tbl_purchase where pid=" & Val(txtpid.Text)
        db.Execute "delete from tbl_purchasebalance where pid=" & Val(txtpid.Text)
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_purchasevoucher where pid=" & Val(txtpid.Text) & "", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            rs.Fields("pid") = ""
            rs.Update
        End If
        rs.Close
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
        rs.Fields("bags") = Val(MSGrid.TextMatrix(i, 0))
        rs.Fields("itemcode") = Val(MSGrid.TextMatrix(i, 1))
        rs.Fields("itemname") = MSGrid.TextMatrix(i, 2)
        rs.Fields("quantity") = MSGrid.TextMatrix(i, 3)
        rs.Fields("itemrate") = Format(Val(MSGrid.TextMatrix(i, 4)), "0.00")
        rs.Fields("itemamt") = Format(Val(MSGrid.TextMatrix(i, 5)), "0.00")
        rs.Fields("cooly") = Format(Val(txtcooly.Text), "0.00")
        rs.Fields("lorryhire") = Format(Val(txtlorryhire.Text), "0.00")
        rs.Fields("gridtotqty") = Format(Val(txtgridtotqty.Text), "0.00")
        rs.Fields("gridtotamt") = Format(Val(txtgridtotamt.Text), "0.00")
        rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
        rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
        rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
        rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
        rs.Update
    Next i
    rs.Close
    
    '====================Purchase Balance====================
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_purchasebalance", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("pid") = Val(txtpid.Text)
        rs.Fields("purdate") = DTPicker1.Value
        rs.Fields("sid") = Val(s)
        rs.Fields("supname") = Trim(txtsupname.Text)
        rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
        rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
        rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
        rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
        rs.Fields("baldesc") = txttotamt.Text & "-" & txtpayamt.Text
    rs.Update
    rs.Close
    '====================Purchase Balance====================
    
    db.Execute "update tbl_purchasevoucher set pid=" & Val(txtpid.Text) & " where sid=" & Val(s)

    'Stock update-------------------------------
    For i = 1 To MSGrid.Rows - 1
        If MSGrid.TextMatrix(i, 0) <> "" Then
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_stock where itemcode=" & MSGrid.TextMatrix(i, 1), db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                rs.Fields("qty") = Val(rs.Fields("qty")) + Val(MSGrid.TextMatrix(i, 3))
                rs.Update
            End If
        End If
    Next i
    'Stock update------------------------------

    MsgBox "Successfully Saved...", vbInformation, "Sumathi Stores"
End If
Call cmdclear_Click
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub MSGrid1_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_itemmaster where itemcode=" & MSGrid1.TextMatrix(MSGrid1.Row, 0), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    MSGrid.TextMatrix(MSGrid.Rows - 1, 1) = rs1.Fields("itemcode")
    MSGrid.TextMatrix(MSGrid.Rows - 1, 2) = rs1.Fields("itemname")
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 3  ' Grid entry was changed to Item Rate coloumn
    MSGrid.SetFocus
End If
rs1.Close
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

Private Sub txtlorryhire_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdcontinue.SetFocus
End If
Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not txtlorryhire.Text = "" Then txtlorryhire.Text = Mid(Trim(txtlorryhire.Text), 1, (Len(txtlorryhire.Text) - 1))
    Case 46         ' 46 keyascii is for dot symbol
        txtlorryhire.Text = txtlorryhire.Text & Chr(KeyAscii)
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        txtlorryhire.Text = txtlorryhire.Text & Chr(KeyAscii)
End Select
End Sub

Private Sub txtlorryhire_LostFocus()
If txtlorryhire.Text <> "" Then
    txtgridtotamt.Text = 0
    For i = 1 To MSGrid.Rows - 1
        txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Grid Total bill amount calculation
    Next i
    txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text) + Val(txtcooly.Text) + Val(txtlorryhire.Text), "0.00")
    txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")
    'txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
    txtcooly.Text = Format(Val(txtcooly.Text), "0.00")
    txtlorryhire.Text = Format(Val(txtlorryhire.Text), "0.00")
End If
cmdcontinue.SetFocus
End Sub

Private Sub txtcooly_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtlorryhire.SetFocus
End If
'If KeyAscii = 13 Then
'    If txtcooly.Text <> "" Then
'        txtgridtotamt.Text = 0
'        For i = 1 To MSGrid.Rows - 1
'            txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Grid Total bill amount calculation
'        Next i
'        txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text) + Val(txtcooly.Text), "0.00")
'        txtbalamt.Text = Format(Val(txttotamt.Text), "0.00")
'        'txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
'        txtcooly.Text = Format(Val(txtcooly.Text), "0.00")
'        cmdcontinue.SetFocus
'    Else
'        MSGrid.Row = MSGrid.Rows - 1
'        MSGrid.Col = 0
'        MSGrid.SetFocus
'        MSGrid.CellBackColor = RGB(117, 145, 233)
'    End If
'End If
Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not txtcooly.Text = "" Then txtcooly.Text = Mid(Trim(txtcooly.Text), 1, (Len(txtcooly.Text) - 1))
    Case 46         ' 46 keyascii is for dot symbol
        txtcooly.Text = txtcooly.Text & Chr(KeyAscii)
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        txtcooly.Text = txtcooly.Text & Chr(KeyAscii)
End Select
End Sub

Private Sub txtcooly_LostFocus()
If txtcooly.Text <> "" Then
    txtgridtotamt.Text = 0
    For i = 1 To MSGrid.Rows - 1
        txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Grid Total bill amount calculation
    Next i
    txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text) + Val(txtcooly.Text), "0.00")
    txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")
    'txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
    txtcooly.Text = Format(Val(txtcooly.Text), "0.00")
    txtlorryhire.SetFocus
Else
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 0
    MSGrid.SetFocus
    MSGrid.CellBackColor = RGB(117, 145, 233)
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
'MsgBox KeyCode
'MsgBox List1.ListIndex
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
If KeyCode = 40 Then    'Down arrow key
    For i = 0 To List1.ListCount - 1
        If List1.ListIndex = i Then
            On Error Resume Next
            List1.ListIndex = i + 1
            Exit For
        End If
    Next i
End If
If KeyCode = 38 Then    'Up arrow key
    For i = 0 To List1.ListCount - 1
        If List1.ListIndex = i Then
            On Error Resume Next
            List1.ListIndex = i - 1
            Exit For
        ElseIf List1.ListIndex = 0 Then
            List1.ListIndex = 0
        End If
    Next i
End If
End Sub

Private Sub txtsupname_LostFocus()
List1.Visible = False

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchasebalance where supname='" & txtsupname.Text & "' and isdel=yes order by pid, purdate", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtobalance.Text = Format(Val(rs.Fields("balamt")), "0.00")
Else
    txtobalance.Text = "0"
End If
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_purchasevoucher where supname='" & txtsupname.Text & "' and pid is null order by id", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    lbl_lastvdate.Caption = Format(rs.Fields("vdate"), "DD/MM/YYYY")
    txtpayamt.Text = Format(Val(rs.Fields("rbalance")), "0.00")
Else
    lbl_lastvdate.Caption = ""
    txtpayamt.Text = "0"
End If
rs.Close
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
Me.BackColor = RGB(35, 29, 29)
Label1.BackColor = RGB(35, 29, 29)
Label2.BackColor = RGB(35, 29, 29)
Label3.BackColor = RGB(35, 29, 29)
Label4.BackColor = RGB(35, 29, 29)
Label5.BackColor = RGB(35, 29, 29)
Label6.BackColor = RGB(35, 29, 29)
Label7.BackColor = RGB(35, 29, 29)
Label8.BackColor = RGB(35, 29, 29)
Label9.BackColor = RGB(35, 29, 29)
Label10.BackColor = RGB(35, 29, 29)
Label11.BackColor = RGB(35, 29, 29)
Label13.BackColor = RGB(35, 29, 29)
lbl_lastvdate.BackColor = RGB(35, 29, 29)
MSGrid.BackColorBkg = RGB(35, 29, 29)
MSGrid1.BackColorBkg = RGB(35, 29, 29)

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

stmt = "select * from tbl_itemmaster order by itemcode"
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid1.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname")
        rs.MoveNext
    Loop
End If
rs.Close

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

CmdModify.Enabled = False
CmdDelete.Enabled = False
End Sub

Private Sub MSGrid_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 112 Then 'F1 Key
'    If Not MSGrid.Rows = 1 Then
'        MSGrid.CellBackColor = vbWhite
'        txtfind.SetFocus
'    End If
'End If
'
'If KeyCode = 117 Then 'F6 Key for Delete the row
'    If Not MSGrid.Rows = 1 Then
'        txtgridtotqty.Text = Format(Val(txtgridtotqty.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 4)), "0.00")
'        txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 5)), "0.00")
'        txttotamt.Text = Format(Val(txtgridtotamt.Text), "0.00")
'
'        MSGrid.Row = MSGrid.Row
'        MSGrid.Col = 0
'        If MSGrid.Row = 1 Then
'            MSGrid.TextMatrix(1, 0) = ""
'            MSGrid.TextMatrix(1, 1) = ""
'            MSGrid.TextMatrix(1, 2) = ""
'            MSGrid.TextMatrix(1, 3) = ""
'            MSGrid.TextMatrix(1, 4) = ""
'            MSGrid.TextMatrix(1, 5) = ""
'        Else
'            MSGrid.RemoveItem MSGrid.Row
'        End If
'        MSGrid.CellBackColor = RGB(117, 145, 233)
'    End If
'End If
'
'If KeyCode = 118 Then 'F7 Key
'    txtinvno.SetFocus
'    txtinvno.SelStart = 0
'    txtinvno.SelLength = Len(txtinvno.Text)    'select the text
'End If
'
'If KeyCode = 119 Then 'F8 Key
'    txtsupname.SetFocus
'    txtsupname.SelStart = 0
'    txtsupname.SelLength = Len(txtsupname.Text)
'End If
'
'If KeyCode = 120 Then 'F9 Key
'    DTPicker1.SetFocus
'End If
'
''If KeyCode = 122 Then 'F11 Key
''    MSGrid.Row = MSGrid.Row
''    MSGrid.Col = 3                 ' Navigate to Item Rate Coloumn in Grid
''    MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
''    MSGrid.SetFocus
''End If
'
'If KeyCode = 27 Then   'esc Key for Clear
'    Call cmdclear_Click
'End If

End Sub

Private Sub MSGrid_KeyPress(KeyAscii As Integer)

If MSGrid.Col = 0 Or MSGrid.Col = 1 Or MSGrid.Col = 3 Or MSGrid.Col = 4 Then      'Bags, Itemcode, Quantity and item rate gridy coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
    Case 46         ' 46 keyascii is for dot symbol
        If MSGrid.Col = 3 Then   'For Itemrate Coloumn
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
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
        If MSGrid.Col = 0 Then  ' Bags
            If MSGrid.TextMatrix(MSGrid.Row, 0) <> "" Then
                MSGrid.Col = 1 'cursor changed to the next column
            Else
                MSGrid.CellBackColor = vbWhite
                txtcooly.SetFocus   'cursor navigation to the cooly textbox
            End If
        End If
        
        If MSGrid.Col = 1 Then      ' Item Code
            If MSGrid.TextMatrix(MSGrid.Row, 1) <> "" Then
                If rs.State = 1 Then rs.Close
                rs.Open "select * from tbl_itemmaster where itemcode=" & Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), db, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    MSGrid.TextMatrix(MSGrid.Row, 1) = Trim(rs.Fields("itemcode"))
                    MSGrid.TextMatrix(MSGrid.Row, 2) = Trim(rs.Fields("itemname"))
                    MSGrid.TextMatrix(MSGrid.Row, 3) = ""
                    MSGrid.Col = 3  ' Grid entry was changed to qty coloumn
                Else
                    MSGrid.TextMatrix(MSGrid.Row, 1) = ""
                    MSGrid.Row = MSGrid.Row
                    MSGrid.Col = 1
                End If
                rs.Close
            End If
        End If

        If MSGrid.Col = 3 Then      ' Quantity
            If MSGrid.TextMatrix(MSGrid.Row, 3) <> "" Then
                MSGrid.Col = 4                  'cursor position changed to the item rate column
            End If
        End If

        If MSGrid.Col = 4 Then      ' Item Rate
            If MSGrid.TextMatrix(MSGrid.Row, 4) <> "" Then
                MSGrid.TextMatrix(MSGrid.Row, 5) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 3)) * Val(MSGrid.TextMatrix(MSGrid.Row, 4)), "0.00")
                        
                txtgridtotamt.Text = 0
                txtgridtotqty.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Grid Total bill amount calculation
                    txtgridtotqty.Text = Val(txtgridtotqty.Text) + Val(MSGrid.TextMatrix(i, 3))   'Total quantity calculation
                Next i
                txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text), "0.00")
                txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")
                'txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")
                        
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

'Sub SendKeys(Text$, Optional wait As Boolean)
'    CreateObject("WScript.Shell").SendKeys Text$, wait
'End Sub

Private Sub MSGrid_EnterCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
If MSGrid.Row > 0 Then
    MSGrid.CellBackColor = RGB(117, 145, 233)
End If
End Sub

Private Sub MSGrid_LeaveCell()
MSGrid.Row = MSGrid.Row
MSGrid.Col = MSGrid.Col
If MSGrid.Row > 0 Then
    MSGrid.CellBackColor = vbWhite
End If
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
        MSGrid.TextMatrix(MSGrid.Row, 1) = rs1.Fields("itemcode")
        MSGrid.TextMatrix(MSGrid.Row, 2) = rs1.Fields("itemname")
        MSGrid.Col = MSGrid.Col + 3  ' Grid entry was changed to Item Rate coloumn
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
        MSGrid1.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname")
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
