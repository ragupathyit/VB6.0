VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00400040&
   Caption         =   "Sri Saravana Bhavan"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   "waiter"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30052
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30764
            Key             =   "settings"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30E76
            Key             =   "help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31588
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31C9A
            Key             =   "table"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":323AC
            Key             =   "Order"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":32AC0
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":331D2
            Key             =   "sales"
            Object.Tag             =   "sales"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36149
            Key             =   "report"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3685B
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   6315
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Copyright @ 2016"
            TextSave        =   "Copyright @ 2016"
            Object.ToolTipText     =   "Sri Saravana Bhavan - Hotel Management System, Mettupalayam"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "06/10/2017"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:00 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   1164
      ButtonWidth     =   2117
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Item Master"
            Key             =   "item"
            ImageKey        =   "item"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Table Master"
            Key             =   "table"
            ImageKey        =   "table"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Waiter Master"
            Key             =   "waiter"
            ImageKey        =   "waiter"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "KOT / Billing"
            Key             =   "order"
            ImageKey        =   "Order"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Key             =   "reports"
            ImageKey        =   "report"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calculator"
            Key             =   "calc"
            ImageKey        =   "calc"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "help"
            ImageKey        =   "help"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Key             =   "settings"
            ImageKey        =   "settings"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "exit"
            ImageKey        =   "exit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Call connect
Toolbar1.Buttons("item").Visible = False
Toolbar1.Buttons("table").Visible = False
Toolbar1.Buttons("waiter").Visible = False
Toolbar1.Buttons("order").Visible = False
Toolbar1.Buttons("reports").Visible = False
Toolbar1.Buttons("settings").Visible = False

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_login where username='" & Trim(uname) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    While Not rs.EOF
        If Trim(rs.Fields("forms")) = "Item Master" Then
            Toolbar1.Buttons("item").Visible = True
        End If
        If Trim(rs.Fields("forms")) = "Table Master" Then
            Toolbar1.Buttons("table").Visible = True
        End If
        If Trim(rs.Fields("forms")) = "Waiter Master" Then
            Toolbar1.Buttons("waiter").Visible = True
        End If
        If Trim(rs.Fields("forms")) = "KOT / Billing" Then
            Toolbar1.Buttons("order").Visible = True
        End If
        If Trim(rs.Fields("forms")) = "Reports" Then
            Toolbar1.Buttons("reports").Visible = True
        End If
        If Trim(rs.Fields("forms")) = "Settings" Then
            Toolbar1.Buttons("settings").Visible = True
        End If
        rs.MoveNext
    Wend
End If
'If Not uname = "admin" Then
'    Toolbar1.Buttons("item").Visible = False
'    Toolbar1.Buttons("table").Visible = False
'    Toolbar1.Buttons("waiter").Visible = False
'    Toolbar1.Buttons("reports").Visible = False
'    Toolbar1.Buttons("settings").Visible = False
'End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call CloseForms
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Not Button.Caption = "Calculator" Then
    Call CloseForms
End If

If Button.Caption = "Item Master" Then
    FrmItemMaster.Show
End If

If Button.Caption = "Table Master" Then
    FrmTableMaster.Show
End If

If Button.Caption = "Waiter Master" Then
    FrmWaiterMaster.Show
End If

If Button.Caption = "KOT / Billing" Then
    FrmKOTOrder.Show
End If

If Button.Caption = "Reports" Then
    FrmReport.Show
End If

If Button.Caption = "Calculator" Then
    Call Shell("calc.exe", vbNormalFocus)
End If

If Button.Caption = "Help" Then
    FrmHelp.Show
End If

If Button.Caption = "Settings" Then
    FrmSettings.Show
End If

If Button.Caption = "Clear Bill No" Then
    If MsgBox("Are you sure to clear billno and start from 1", vbYesNo, "Sri Saravana Bhavan") = vbYes Then
        db.Execute "delete from tbl_tempbill"
        MsgBox "Bill is cleared successfully and start from Bill No 1", vbInformation, "Sri Saravana Bhavan"
    End If
End If

If Button.Caption = "Exit" Then
    End
End If

End Sub

Public Function CloseForms()
For Each frm In Forms
    If Not frm Is MDIForm1 Then
        Unload frm
    End If
Next
Set frm = Nothing
End Function
