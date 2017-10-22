VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ItemFrm 
   Caption         =   "Item Details"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   450
   ClientWidth     =   12690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleMode       =   0  'User
   ScaleWidth      =   12934.04
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSave 
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
      Left            =   1800
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdModify 
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
      Left            =   3600
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdDelete 
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
      Left            =   7320
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&CLOSE"
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
      Left            =   9240
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
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
      Left            =   5400
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   3855
      Left            =   5400
      TabIndex        =   7
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16761024
      AllowUserResizing=   2
      FormatString    =   "Item Code  |< Item Name                                                   |< Item Rate     "
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
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblhide 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   4920
      Width           =   12495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "ITEM DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   5160
      TabIndex        =   8
      Top             =   360
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   12495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Item Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   3720
      Width           =   1065
   End
End
Attribute VB_Name = "ItemFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub cmdclear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
CmdSave.Enabled = True
CmdModify.Enabled = False
cmdclear.Enabled = True
CmdDelete.Enabled = False
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
db.Execute "delete from tbl_itemmaster where itemcode=" & lblhide.Caption
MsgBox "Successfully Deleted...", vbInformation, "Fresh Park"

Call cmdclear_Click
Call Fill
CmdSave.Enabled = True
CmdModify.Enabled = False
cmdclear.Enabled = True
CmdDelete.Enabled = False
End Sub

Private Sub CmdModify_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_itemmaster where itemcode=" & lblhide.Caption, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.Fields("itemcode") = UCase(Text1.Text)
    rs.Fields("itemname") = UCase(Text2.Text)
    rs.Fields("rate") = UCase(Text3.Text)
    rs.Update
End If
rs.Close
        
MsgBox "Successfully Modified", vbInformation, "Fresh Park"
Call cmdclear_Click
Call Fill
CmdSave.Enabled = True
CmdModify.Enabled = False
cmdclear.Enabled = True
CmdDelete.Enabled = False
End Sub

Private Sub CmdSave_Click()
'-------------------Validation Starts Here-----------------------------
If Text1.Text = "" Then
    MsgBox "Enter the Item Code Properly...", vbInformation
    Text1.SetFocus
ElseIf Text2.Text = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation
    Text2.SetFocus
ElseIf Text3.Text = "" Then
    MsgBox "Enter the Item Rate Properly...", vbInformation
    Text3.SetFocus
ElseIf Not IsNumeric(Text3.Text) Then
    MsgBox "Enter the Item Rate Properly...", vbInformation
    Text3.Text = ""
    Text3.SetFocus
Else
'-------------------Validation Ends Here-------------------------------

    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemcode=" & UCase(Text1.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "Please Type Different Itemcode", vbInformation, "Fresh Park"
    Else
        rs.AddNew
        rs.Fields("itemcode") = UCase(Text1.Text)
        rs.Fields("itemname") = UCase(Text2.Text)
        rs.Fields("rate") = UCase(Text3.Text)
        rs.Update
        
        MsgBox "Item Saved Successfully", vbInformation, "Fresh Park"
    End If
    rs.Close
    
    Call cmdclear_Click
    Call Fill
End If
End Sub

Private Sub Form_Load()
If db.State = 1 Then db.Close
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\master.mdb" & ";jet oledb:database password=ragu_24993"
Call Fill
CmdSave.Enabled = True
CmdModify.Enabled = False
cmdclear.Enabled = True
CmdDelete.Enabled = False
End Sub

Private Sub MsGrid_Click()
    Text1.Text = MSGrid.TextMatrix(MSGrid.Row, 0)
    lblhide.Caption = MSGrid.TextMatrix(MSGrid.Row, 0)
    Text2.Text = MSGrid.TextMatrix(MSGrid.Row, 1)
    Text3.Text = MSGrid.TextMatrix(MSGrid.Row, 2)
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    cmdclear.Enabled = True
    CmdDelete.Enabled = True
End Sub

Private Function Fill()
stmt = "select * from tbl_itemmaster order by itemcode"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
MSGrid.Rows = 1
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        MSGrid.AddItem rs.Fields("itemcode") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("rate")
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text = "" Then
        Text1.SetFocus
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_itemmaster where itemcode=" & Val(Text1.Text)
        If Not rs.EOF Then
            lblhide.Caption = rs.Fields("itemcode")
            Text1.Text = rs.Fields("itemcode")
            Text2.Text = rs.Fields("itemname")
            Text3.Text = rs.Fields("rate")
            
            Text3.SetFocus
            Text3.SelStart = 0
            Text3.SelLength = Len(Text3.Text)    'select the text
            
            CmdSave.Enabled = False
            CmdModify.Enabled = True
            cmdclear.Enabled = True
            CmdDelete.Enabled = True
        Else
            Text2.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text2.Text = "" Then
        Text2.SetFocus
    Else
        Text3.SetFocus
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text3.Text = "" Then
        Text3.SetFocus
    Else
        If CmdSave.Enabled = True Then
            CmdSave.SetFocus
        Else
            CmdModify.SetFocus
        End If
    End If
End If
End Sub
