VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SupplierFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Supplier Details"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   12240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text4 
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
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
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
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Height          =   855
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5280
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9120
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7200
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton CmdModify 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3480
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1680
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   4215
      Left            =   5280
      TabIndex        =   9
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorBkg    =   16711680
      GridColor       =   12582912
      GridColorFixed  =   12582912
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "Sup ID  |Supplier Name                                               |Phone/Cell No     "
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "E-Mail ID"
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
      Top             =   4200
      Width           =   1365
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SUPPLIER DETAILS"
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
      Left            =   4440
      TabIndex        =   14
      Top             =   240
      Width           =   3810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Phone No"
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
      Top             =   3480
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Address"
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
      TabIndex        =   12
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Name"
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
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblhide 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "SupplierFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
Unload Me
SupplierFrm.Show
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub CmdDelete_Click()
If rs.State = 1 Then rs.Close
db.Execute "delete from tbl_suppliermaster where sid=" & lblhide.Caption
MsgBox "Successfully Deleted...", vbInformation, "Sumathi Stores"
Call cmdclear_Click
End Sub

Private Sub CmdModify_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_suppliermaster where sid=" & lblhide.Caption, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.Fields("sid") = lblhide.Caption
    rs.Fields("suppliername") = UCase(Text1.Text)
    rs.Fields("supplieraddress") = UCase(Text2.Text)
    rs.Fields("phone") = Text3.Text
    rs.Fields("email") = Text4.Text
    rs.Update
End If
rs.Close
MsgBox "Successfully Modified", vbInformation, "Sumathi Stores"
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If Text1.Text = "" Then
    MsgBox "Enter the Supplier Name Properly...", vbInformation, "Sumathi Stores"
    Text1.SetFocus
ElseIf Text2.Text = "" Then
    MsgBox "Enter the Supplier Address Properly...", vbInformation, "Sumathi Stores"
    Text2.SetFocus
ElseIf Text3.Text = "" Then
    MsgBox "Enter the Phone/Cell No Properly...", vbInformation, "Sumathi Stores"
    Text3.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_suppliermaster order by sid", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveLast
        s = rs.Fields("sid") + 1
    Else
        s = 1
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_suppliermaster where suppliername='" & UCase(Text1.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
        rs.Fields("sid") = CInt(s)
        rs.Fields("suppliername") = UCase(Text1.Text)
        rs.Fields("supplieraddress") = UCase(Text2.Text)
        rs.Fields("phone") = UCase(Text3.Text)
        rs.Fields("email") = UCase(Text4.Text)
        rs.Update
        MsgBox "Supplier Details Saved Successfully", vbInformation, "Sumathi Stores"
        rs.Close
    Else
        MsgBox "Enter The Another Name for Supplier", vbInformation, "Sumathi Stores"
    End If
    Call cmdclear_Click
End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(35, 29, 29)
Label1.BackColor = RGB(35, 29, 29)
Label2.BackColor = RGB(35, 29, 29)
Label3.BackColor = RGB(35, 29, 29)
Label4.BackColor = RGB(35, 29, 29)
Label6.BackColor = RGB(35, 29, 29)
MSGrid.BackColorBkg = RGB(35, 29, 29)

Call connect
Call Fill
CmdSave.Enabled = True
CmdModify.Enabled = False
CmdDelete.Enabled = False
End Sub

Private Sub MSGrid_Click()
lblhide.Caption = MSGrid.TextMatrix(MSGrid.Row, 0)
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_suppliermaster where sid=" & MSGrid.TextMatrix(MSGrid.Row, 0), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Text1.Text = rs.Fields("suppliername")
    If Not rs.Fields("supplieraddress") = "" Then
        Text2.Text = rs.Fields("supplieraddress")
    End If
    If Not rs.Fields("phone") = "" Then
        Text3.Text = rs.Fields("phone")
    End If
    If Not rs.Fields("email") = "" Then
        Text4.Text = rs.Fields("email")
    End If
End If
CmdSave.Enabled = False
CmdModify.Enabled = True
CmdDelete.Enabled = True
End Sub

Private Function Fill()
stmt = "select * from tbl_suppliermaster order by sid"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    While Not rs.EOF
        MSGrid.AddItem rs.Fields("sid") & vbTab & rs.Fields("suppliername") & vbTab & rs.Fields("phone")
        rs.MoveNext
    Wend
End If
rs.Close
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub
