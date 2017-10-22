VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCustomer 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   6120
      TabIndex        =   28
      Top             =   1440
      Width           =   5895
      Begin VB.Label Label15 
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
         Left            =   4680
         TabIndex        =   31
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   1680
         TabIndex        =   30
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C. ID"
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
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tamil-Ananthi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      Top             =   840
      Width           =   2295
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   11760
      TabIndex        =   17
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
      MICON           =   "FrmCustomer.frx":0000
      PICN            =   "FrmCustomer.frx":001C
      PICH            =   "FrmCustomer.frx":072E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FrameForm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   5535
      Begin VB.TextBox txtadvance 
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   27
         Top             =   4200
         Width           =   1215
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtphoneno 
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
         Left            =   2040
         TabIndex        =   4
         Top             =   3000
         Width           =   1935
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "641047"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtstate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tamil-Ananthi"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   2
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtaddress1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tamil-Ananthi"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   1
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtcustname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tamil-Ananthi"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   0
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtcustid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   2040
         TabIndex        =   6
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance *"
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
         Left            =   0
         TabIndex        =   26
         Top             =   4320
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State"
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
         Left            =   0
         TabIndex        =   16
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+91"
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
         Left            =   2040
         TabIndex        =   15
         Top             =   3720
         Width           =   435
      End
      Begin VB.Label Label9 
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
         Left            =   0
         TabIndex        =   14
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
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
         Left            =   0
         TabIndex        =   13
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label Label7 
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
         Left            =   0
         TabIndex        =   12
         Top             =   2520
         Width           =   855
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
         Left            =   0
         TabIndex        =   11
         Top             =   1320
         Width           =   1785
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
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   1365
      End
   End
   Begin Project1.Button BtnSave 
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      ToolTipText     =   "SAVE"
      Top             =   6120
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
      MICON           =   "FrmCustomer.frx":0E40
      PICN            =   "FrmCustomer.frx":0E5C
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
      TabIndex        =   19
      ToolTipText     =   "MODIFY"
      Top             =   6120
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
      MICON           =   "FrmCustomer.frx":156E
      PICN            =   "FrmCustomer.frx":158A
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
      TabIndex        =   20
      ToolTipText     =   "DELETE"
      Top             =   6120
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
      MICON           =   "FrmCustomer.frx":1C9C
      PICN            =   "FrmCustomer.frx":1CB8
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
      TabIndex        =   21
      ToolTipText     =   "CLEAR"
      Top             =   6120
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
      MICON           =   "FrmCustomer.frx":23CA
      PICN            =   "FrmCustomer.frx":23E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnFind 
      Height          =   375
      Left            =   10680
      TabIndex        =   23
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Find"
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
      MICON           =   "FrmCustomer.frx":2AF8
      PICN            =   "FrmCustomer.frx":2B14
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
      Height          =   4575
      Left            =   6000
      TabIndex        =   25
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "           .|                                                           .|                        ."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tamil-Ananthi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.Name/Mobile No"
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
      Left            =   6120
      TabIndex        =   24
      Top             =   960
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmCustomer.frx":3226
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
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
      Width           =   3540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   12735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   6000
      Width           =   12735
   End
End
Attribute VB_Name = "FrmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnDelete_Click()
db.Execute "delete from tbl_customer where cid=" & Val(txtcustid.Text)
MsgBox "Customer Deleted Successfully", vbInformation, "Press Management"
Call BtnClear_Click
End Sub

Private Sub BtnFind_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_customer where customername like'" & Trim(txtsearch.Text) & "%'", db, adOpenDynamic, adLockOptimistic
If rs2.State = 1 Then rs2.Close
rs2.Open "select * from tbl_customer where mobileno='" & Trim(txtsearch.Text) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    MSGrid.Rows = 1
    rs1.MoveFirst
    Do While Not rs1.EOF
        MSGrid.AddItem rs1.Fields("cid") & vbTab & rs1.Fields("customername") & vbTab & rs1.Fields("mobileno") & vbTab & rs1.Fields("ctype")
        rs1.MoveNext
    Loop
    rs1.Close
ElseIf Not rs2.EOF Then
    MSGrid.Rows = 1
    rs2.MoveFirst
    Do While Not rs2.EOF
        MSGrid.AddItem rs2.Fields("cid") & vbTab & rs2.Fields("customername") & vbTab & rs2.Fields("mobileno") & vbTab & rs1.Fields("ctype")
        rs2.MoveNext
    Loop
    rs2.Close
Else
    MsgBox "Customer is not found.", vbInformation, "Press Management"
End If
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmCustomer.Show
End Sub

Private Sub BtnModify_Click()
'-------------------Validation Starts Here-----------------------------
If txtcustname.Text = "" Then
    MsgBox "Enter the Customer Name Properly...", vbInformation, "Press Management"
    txtcustname.SetFocus
ElseIf txtaddress1.Text = "" Then
    MsgBox "Enter the Address Line 1 Properly...", vbInformation, "Press Management"
    txtaddress1.SetFocus
ElseIf Not IsNumeric(Val(txtmobileno.Text)) Then
    MsgBox "Enter the Mobile Number Properly...", vbInformation, "Press Management"
    txtmobileno.Text = ""
    txtmobileno.SetFocus
ElseIf Not IsNumeric(Val(txtadvance.Text)) Then
    MsgBox "Enter the Advance Amount Properly...", vbInformation, "Press Management"
    txtadvance.Text = ""
    txtadvance.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_customer where cid=" & Val(Trim(txtcustid.Text)), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.Fields("cid") = Val(Trim(txtcustid.Text))
        rs.Fields("customername") = Trim(txtcustname.Text)
        rs.Fields("address1") = Trim(txtaddress1.Text)
        rs.Fields("state") = Trim(txtstate.Text)
        rs.Fields("pincode") = Trim(txtpincode.Text)
        rs.Fields("phoneno") = Trim(txtphoneno.Text)
        rs.Fields("mobileno") = Trim(txtmobileno.Text)
        rs.Fields("advance") = Trim(txtadvance.Text)
        rs.Update
    End If
    rs.Close
    MsgBox "Customer Modified Successfully", vbInformation, "Press Management"
    
    Call BtnClear_Click
End If
End Sub

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If txtcustname.Text = "" Then
    MsgBox "Enter the Customer Name Properly...", vbInformation, "Press Management"
    txtcustname.SetFocus
ElseIf txtaddress1.Text = "" Then
    MsgBox "Enter the Address Line 1 Properly...", vbInformation, "Press Management"
    txtaddress1.SetFocus
ElseIf Not IsNumeric(Val(txtmobileno.Text)) Then
    MsgBox "Enter the Mobile Number Properly...", vbInformation, "Press Management"
    txtmobileno.Text = ""
    txtmobileno.SetFocus
ElseIf Not IsNumeric(Val(txtadvance.Text)) Then
    MsgBox "Enter the Advance Amount Properly...", vbInformation, "Press Management"
    txtadvance.Text = ""
    txtadvance.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_customer", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs.Fields("cid") = Val(Trim(txtcustid.Text))
    rs.Fields("customername") = Trim(txtcustname.Text)
    rs.Fields("address1") = Trim(txtaddress1.Text)
    rs.Fields("state") = Trim(txtstate.Text)
    rs.Fields("pincode") = Trim(txtpincode.Text)
    rs.Fields("phoneno") = Trim(txtphoneno.Text)
    rs.Fields("mobileno") = Trim(txtmobileno.Text)
    rs.Fields("advance") = Trim(txtadvance.Text)
    rs.Update
    rs.Close
    MsgBox "Customer Saved Successfully", vbInformation, "Press Management"
    
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Function Fill()
stmt = "select * from tbl_customer order by cid"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("cid") & vbTab & rs.Fields("customername") & vbTab & rs.Fields("mobileno")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select cid from tbl_customer order by cid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtcustid.Text = rs.Fields("cid") + 1
Else
    txtcustid.Text = 1
End If
rs.Close

BtnSave.Enabled = True
BtnModify.Enabled = False
BtnDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_customer where cid=" & Val(Trim(MSGrid.TextMatrix(MSGrid.Row, 0))), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    txtcustid.Text = rs1.Fields("cid")
    txtcustname.Text = rs1.Fields("customername")
    txtaddress1.Text = IIf(IsNull(rs1.Fields("address1")), "", rs1.Fields("address1"))
    txtstate.Text = IIf(IsNull(rs1.Fields("state")), "", rs1.Fields("state"))
    txtpincode.Text = IIf(IsNull(rs1.Fields("pincode")), "", rs1.Fields("pincode"))
    txtphoneno.Text = IIf(IsNull(rs1.Fields("phoneno")), "", rs1.Fields("phoneno"))
    txtmobileno.Text = IIf(IsNull(rs1.Fields("mobileno")), "", rs1.Fields("mobileno"))
    txtadvance.Text = IIf(IsNull(rs1.Fields("advance")), "", rs1.Fields("advance"))
End If

txtcustname.SetFocus
BtnSave.Enabled = False
BtnModify.Enabled = True
BtnDelete.Enabled = True
End Sub

Private Sub txtadvance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BtnSave.Enabled = True Then
        BtnSave.SetFocus
    Else
        BtnModify.SetFocus
    End If
End If
End Sub

Private Sub txtcustname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress1.SetFocus
    txtaddress1.SelStart = 0
    txtaddress1.SelLength = Len(txtaddress1.Text)    'select the text
End If
End Sub

Private Sub txtaddress1_KeyPress(KeyAscii As Integer)
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
    txtphoneno.SetFocus
    txtphoneno.SelStart = 0
    txtphoneno.SelLength = Len(txtphoneno.Text)    'select the text
End If
End Sub

Private Sub txtphoneno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtmobileno.SetFocus
    txtmobileno.SelStart = 0
    txtmobileno.SelLength = Len(txtmobileno.Text)    'select the text
End If
End Sub

Private Sub txtmobileno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtadvance.SetFocus
    txtadvance.SelStart = 0
    txtadvance.SelLength = Len(txtadvance.Text)    'select the text
End If
End Sub
