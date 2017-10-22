VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCCashSales 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Customer Cash Sales"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   15030
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtmobno 
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txttotqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   24
      Text            =   "0"
      Top             =   7440
      Width           =   735
   End
   Begin VB.TextBox txtstax 
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
      Left            =   8280
      TabIndex        =   20
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   15
      Text            =   "0"
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox txtbillno 
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   14040
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
      MICON           =   "FrmCCashSales.frx":0000
      PICN            =   "FrmCCashSales.frx":001C
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
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "SAVE"
      Top             =   8160
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
      MICON           =   "FrmCCashSales.frx":072E
      PICN            =   "FrmCCashSales.frx":074A
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
      Left            =   7800
      TabIndex        =   6
      ToolTipText     =   "DELETE"
      Top             =   8160
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
      MICON           =   "FrmCCashSales.frx":0E5C
      PICN            =   "FrmCCashSales.frx":0E78
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
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "CLEAR"
      Top             =   8160
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
      MICON           =   "FrmCCashSales.frx":158A
      PICN            =   "FrmCCashSales.frx":15A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnPrevious 
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      ToolTipText     =   "Previous"
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
      MICON           =   "FrmCCashSales.frx":1CB8
      PICN            =   "FrmCCashSales.frx":1CD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnNext 
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      ToolTipText     =   "Next"
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
      MICON           =   "FrmCCashSales.frx":23E6
      PICN            =   "FrmCCashSales.frx":2402
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstItemName 
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
      Height          =   5700
      ItemData        =   "FrmCCashSales.frx":2B14
      Left            =   120
      List            =   "FrmCCashSales.frx":2B16
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1800
      Width           =   4545
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5655
      Left            =   4680
      TabIndex        =   2
      Top             =   1800
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Item Name                                                                   |Qty    |I. Amt      |Amount         "
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
   Begin Project1.Button BtnBill 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "BILL"
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Bill     "
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
      MICON           =   "FrmCCashSales.frx":2B18
      PICN            =   "FrmCCashSales.frx":2B34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtp_bdate 
      Height          =   375
      Left            =   11640
      TabIndex        =   23
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   97517571
      CurrentDate     =   42560
   End
   Begin VB.Label lblpayamt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   13230
      TabIndex        =   28
      Top             =   8040
      Width           =   1170
   End
   Begin MSForms.ComboBox cmbcname 
      Height          =   360
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   3705
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6535;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
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
      Left            =   9480
      TabIndex        =   26
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot Amt"
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
      Left            =   11880
      TabIndex        =   25
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   10680
      TabIndex        =   22
      Top             =   360
      Width           =   900
   End
   Begin VB.Label lblcancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANCEL BILL"
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
      Left            =   9480
      TabIndex        =   21
      Top             =   8160
      Width           =   3180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax 15%"
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
      Left            =   6360
      TabIndex        =   19
      Top             =   7560
      Width           =   1845
   End
   Begin MSForms.ComboBox cmbcid 
      Height          =   360
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   825
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1455;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label7 
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
      Left            =   3240
      TabIndex        =   18
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id *"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM NAME"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot QTY"
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
      Left            =   9840
      TabIndex        =   14
      Top             =   7560
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
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
      Left            =   8760
      TabIndex        =   11
      Top             =   0
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmCCashSales.frx":2D59
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER CASH SALES"
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
      Width           =   4200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   7920
      Width           =   14655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   9015
      Left            =   0
      Top             =   -120
      Width           =   14655
   End
End
Attribute VB_Name = "FrmCCashSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub BtnBill_Click()
'Dim xlapp As Excel.Application
'Dim xlbook As Excel.Workbook
'Dim xlsheet As Excel.Worksheet
'Set xlapp = CreateObject("excel.application")
'Set xlbook = xlapp.Workbooks.Add
'Set xlsheet = xlbook.Worksheets(1)
'
''xlsheet.Rows.WrapText = True
'xlsheet.Rows.RowHeight = 15  '----------------------Excel Row height
'xlsheet.Rows.Font.Name = "Arial"  '------------Each Row in this font
'xlsheet.Rows.Font.Size = 10
''------------------------Page setup---------------------------------------
'xlsheet.PageSetup.PaperSize = xlPaperA5
''xlsheet.PageSetup.PaperSize = xlPaperUser
'xlsheet.PageSetup.LeftMargin = Application.InchesToPoints(0.5)
'xlsheet.PageSetup.RightMargin = Application.InchesToPoints(0.2)
'xlsheet.PageSetup.TopMargin = Application.InchesToPoints(0.3)
'xlsheet.PageSetup.BottomMargin = Application.InchesToPoints(0.3)
'xlsheet.PageSetup.HeaderMargin = Application.InchesToPoints(0.2)
'xlsheet.PageSetup.FooterMargin = Application.InchesToPoints(0.2)
'xlsheet.PageSetup.Orientation = xlPortrait
''---------------1th Row-------------------------------------------------
'xlsheet.Range("A1").EntireRow.RowHeight = 20
'xlsheet.Range("A1:D1").Merge
'xlsheet.Range("A1:D1").Font.Size = 16
'xlsheet.Range("A1:D1").Font.Bold = True
''xlsheet.Range("A1:D1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'xlsheet.Cells(1, 1).Value = "A R I C"
''---------------2nd Row-------------------------------------------------
'xlsheet.Range("A2:D2").Merge
'xlsheet.Range("A2:D2").Font.Size = 12
'xlsheet.Range("A2:D2").Font.Bold = True
''xlsheet.Range("A2:D2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'xlsheet.Cells(2, 1).Value = "Advanced Research Instrumentation Centre"
''---------------3rd Row-------------------------------------------------
'xlsheet.Range("A3").EntireRow.RowHeight = 12
'xlsheet.Range("A3:D3").Merge
'xlsheet.Range("A3:D3").Font.Size = 8
''xlsheet.Range("A3:D3").Font.Bold = True
''xlsheet.Range("A3:D3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'xlsheet.Cells(3, 1).Value = "Sri Ramakrishna Mission Vidyalaya College of Arts and Science, Coimbatore-641020"
''---------------4th Row-------------------------------------------------
'xlsheet.Range("A4:D4").Merge
'xlsheet.Range("A4:D4").Font.Size = 12
'xlsheet.Range("A4:D4").Font.Bold = True
'xlsheet.Range("A4:D4").Font.Underline = True
'xlsheet.Range("A4:D4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'xlsheet.Cells(4, 1).Value = "INVOICE"
''---------------5th Row--------------------------------------------------
'xlsheet.Range("A5").Font.Bold = True
'xlsheet.Cells(5, 1).Value = "Course Name: " & Trim(cmbcoursename.Text)
'xlsheet.Range("D5").Font.Bold = True
'xlsheet.Cells(5, 4).Value = "Rpt No: " & Trim(txtbillno.Text)
'xlsheet.Cells(5, 4).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
''---------------6th Row-------------------------------------------------
'xlsheet.Range("A6").Font.Bold = True
'xlsheet.Cells(6, 1).Value = "Student Name: " & Trim(txtstudname.Text)
'xlsheet.Range("D6").Font.Bold = True
'xlsheet.Cells(6, 4).Value = "Date: " & Format(dtp_bdate.Value, "dd/mm/yyyy")
'xlsheet.Cells(6, 4).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
''---------------7th Row (Heading Row)------------------------------------
'xlsheet.Range("A7").Font.Bold = True
'xlsheet.Cells(7, 1).Value = "Address Line1: " & Trim(txtaddress1.Text)
''---------------8th Row (Heading Row)------------------------------------
'If Not txtaddress2.Text = "" Then
'    xlsheet.Range("A8").Font.Bold = True
'    xlsheet.Cells(8, 1).Value = "Address Line2: " & Trim(txtaddress2.Text)
'End If
''---------------9th Row (Heading Row)------------------------------------
''xlsheet.Range("A9").EntireRow.RowHeight = 20
'xlsheet.Range("A9:D9").Font.Bold = True
'xlsheet.Range("A9:D9").Font.Size = 11
'xlsheet.Range("A9:D9").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'
'xlsheet.Cells(9, 1).Value = "S.No"
'xlsheet.Cells(9, 2).Value = "Name of the Fees"
'xlsheet.Cells(9, 3).Value = "Qty"
'xlsheet.Cells(9, 4).Value = "Amount"
'
'xlsheet.Range("A9:D9").Borders.LineStyle = xlContinuous  '-----------This is for borders in all sides
''xlsheet.Range("A9:D9").Borders(xlEdgeLeft).LineStyle = xlContinuous
''xlsheet.Range("A9:D9").Borders(xlEdgeTop).LineStyle = xlContinuous
''xlsheet.Range("A9:D9").Borders(xlEdgeRight).LineStyle = xlContinuous
''xlsheet.Range("A9:D9").Borders(xlEdgeBottom).LineStyle = xlContinuous
''xlsheet.Range("A9:D9").Borders(xlInsideVertical).LineStyle = xlContinuous
''--------------------------Each Column Width------------------------------
''xlapp.Columns.AutoFit  ---------------------Automitically fits the column
'xlsheet.Range("A1").EntireColumn.ColumnWidth = 5
'xlsheet.Range("B1").EntireColumn.ColumnWidth = 40
'xlsheet.Range("C1").EntireColumn.ColumnWidth = 8
'xlsheet.Range("D1").EntireColumn.ColumnWidth = 10
''--------------------------Each Column Width------------------------------
'
'Set i = Nothing
'Set j = Nothing
'For i = 1 To 16
''For i = 1 To MSGrid.Rows - 2
'    For j = 1 To MSGrid.Cols
'        If Not j = 3 Then   '------------------------------------------------- To omit the 3rd coloumn of grid i.e: Fee Amt
'            If i <= MSGrid.Rows - 2 Then
'                xlsheet.Cells(i + 9, 1).Value = i
'                If j = 4 Then
'                    xlsheet.Cells(i + 9, j).Value = MSGrid.TextMatrix(i, j - 1)
'                Else
'                    xlsheet.Cells(i + 9, j + 1).Value = MSGrid.TextMatrix(i, j - 1)
'                End If
'            Else
'                xlsheet.Cells(i + 9, j + 1).Value = ""
'            End If
'        End If
'        '--------------------Border---------------------------------------------------------
'        xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        'xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeTop).LineStyle = xlContinuous
'        xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeRight).LineStyle = xlContinuous
'        'xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
'        xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlInsideVertical).LineStyle = xlContinuous
'        '--------------------Number Format 0.00-------------------------------------------
'        xlsheet.Range("D" & i + 9).NumberFormat = "0.00"
'    Next j
'Next i
'
''--------------------Border---------------------------------------------------------
'xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
''xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeTop).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeRight).LineStyle = xlContinuous
''xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 9 & ":D" & i + 9).Borders(xlInsideVertical).LineStyle = xlContinuous
'
''-----------------------------------------Extra Row for Total Amount-----------------------------
'xlsheet.Range("A" & i + 10 & ":C" & i + 10).Merge
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Font.Bold = True
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Font.Size = 11
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'xlsheet.Cells(i + 10, 1).Value = "Total Amount   "
'xlsheet.Cells(i + 10, 4).Value = txttotamt.Text
'xlsheet.Range("D" & i + 10).NumberFormat = "0.00"
''--------------------Border---------------------------------------------------------
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Borders(xlEdgeRight).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 10 & ":D" & i + 10).Borders(xlInsideVertical).LineStyle = xlContinuous
'
''----------------------Extra Row for Service Tax 15% Amount-----------------------------
'xlsheet.Range("A" & i + 11 & ":C" & i + 11).Merge
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Font.Bold = True
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Font.Size = 11
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'xlsheet.Cells(i + 11, 1).Value = "Service Tax 15%   "
'xlsheet.Cells(i + 11, 4).Value = txtstax.Text
'xlsheet.Range("D" & i + 11).NumberFormat = "0.00"
''--------------------Border---------------------------------------------------------
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Borders(xlEdgeTop).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Borders(xlEdgeRight).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 11 & ":D" & i + 11).Borders(xlInsideVertical).LineStyle = xlContinuous
'
''----------------------Extra Row for Payment Amount-----------------------------
'xlsheet.Range("A" & i + 12 & ":C" & i + 12).Merge
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Font.Bold = True
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Font.Size = 11
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'xlsheet.Cells(i + 12, 1).Value = "Payment Amount Rs.   "
'xlsheet.Cells(i + 12, 4).Value = txtpayamt.Text
'
'word = ConNumToEngLish(Val(Format(txtpayamt.Text, "0.00")))
'
'xlsheet.Range("D" & i + 12).NumberFormat = "0.00"
''--------------------Border---------------------------------------------------------
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Borders(xlEdgeLeft).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Borders(xlEdgeRight).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlsheet.Range("A" & i + 12 & ":D" & i + 12).Borders(xlInsideVertical).LineStyle = xlContinuous
'
''If Not txttotal.Text = "" Then
''    If Not txttotal.Text = "0.00" Then
''        xlsheet.Range("A" & i + 16 & ":D" & i + 16).Merge
''        xlsheet.Range("A" & i + 16 & ":D" & i + 16).Font.Bold = True
''        xlsheet.Range("A" & i + 16 & ":D" & i + 16).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
''        xlsheet.Cells(i + 16, 1).Value = "VAT " & Format(txtvat.Text, "0.00") & "%   (" & Format(txttotamt.Text, "0.00") & "+" & Format(Val(txttotal.Text) - Val(txttotamt.Text), "0.00") & ") : "
''        'xlsheet.Cells(i + 16, 4).Value = Val(txttotal.Text) - Val(txttotamt.Text)
''        'xlsheet.Range("D" & i + 16).NumberFormat = "0.00"
''        xlsheet.Cells(i + 16, 4).Value = txttotal.Text
''        xlsheet.Range("D" & i + 16).NumberFormat = "0.00"
''        word = ConNumToEngLish(Val(Format(txttotal.Text, "0.00")))
''    Else
''        word = ConNumToEngLish(Val(Format(txttotamt.Text, "0.00")))
''    End If
''Else
''    word = ConNumToEngLish(Val(Format(txttotamt.Text, "0.00")))
''End If
''
''xlsheet.Range("A" & i + 17 & ":D" & i + 17).Font.Bold = True
''xlsheet.Cells(i + 17, 1).Value = word & " Rupees Only"
'
''xlsheet.PageSetup.RightFooter = "Authorised Signature             "
'If Not Trim(txtddno.Text) = "" Then
'    xlsheet.Range("A" & i + 13 & ":D" & i + 13).Font.Bold = True
'    xlsheet.Cells(i + 13, 1).Value = "DD/Cheque No: " & Trim(txtddno.Text) & "          DD/Cheque Date: " & Format(dtpdddate.Value, "dd/mm/yyyy")
'End If
'
'xlsheet.Range("A" & i + 14 & ":D" & i + 14).Font.Bold = True
'xlsheet.Cells(i + 14, 1).Value = "Rupees " & word & " only"
'
''xlsheet.Cells(i + 15, 1).Value = "*ST-Service Tax, EC-Education Cess, HEC-Higher Edu.Cess"
''xlsheet.Range("D" & i + 15 & ":D" & i + 15).Font.Bold = True
''xlsheet.Range("D" & i + 15 & ":D" & i + 15).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
''xlsheet.Cells(i + 15, 4).Value = "E&OE"
''
''xlsheet.Cells(i + 16, 1).Value = "#(Rate per hour with a minimum of charge applicable)"
'
'xlsheet.Range("A" & i + 15).EntireRow.RowHeight = 10
'
'xlsheet.Range("A" & i + 16).EntireRow.RowHeight = 10
'xlsheet.Range("A" & i + 16 & ":D" & i + 16).Font.Size = 6
'xlsheet.Range("A" & i + 16 & ":D" & i + 16).Font.Bold = True
'xlsheet.Cells(i + 16, 1).Value = "Note:"
'
'xlsheet.Range("A" & i + 17).EntireRow.RowHeight = 10
'xlsheet.Range("A" & i + 17 & ":D" & i + 17).Font.Size = 6
'xlsheet.Cells(i + 17, 1).Value = "*) The Payment should be made either in Cash/NEFT/DD drawn in favour of"
'
'xlsheet.Range("A" & i + 18).EntireRow.RowHeight = 10
'xlsheet.Range("A" & i + 18 & ":D" & i + 18).Font.Size = 6
'xlsheet.Cells(i + 18, 1).Value = "   The Principal,SRMVCAS,Bank Commission,if any be borne by the customer"
'
''xlsheet.Range("A" & i + 20).EntireRow.RowHeight = 11
''xlsheet.Range("A" & i + 20 & ":D" & i + 20).Font.Size = 8
''xlsheet.Cells(i + 20, 1).Value = "*) Deduction of Income Tax at source is not applicable for payments to ARIC"
''
''xlsheet.Range("A" & i + 21).EntireRow.RowHeight = 11
''xlsheet.Range("A" & i + 21 & ":D" & i + 21).Font.Size = 8
''xlsheet.Cells(i + 21, 1).Value = "   under Sec. 12A(a) of Income Tax Act 1961."
'
'xlsheet.Range("A" & i + 19).EntireRow.RowHeight = 10
'xlsheet.Range("A" & i + 19 & ":D" & i + 19).Font.Size = 6
'xlsheet.Cells(i + 19, 1).Value = "*) Our Service Tax Registration Number is AAAAR1077PSD008"
'
''xlsheet.Range("A" & i + 23).EntireRow.RowHeight = 11
''xlsheet.Range("A" & i + 23 & ":D" & i + 23).Font.Size = 8
''xlsheet.Cells(i + 23, 1).Value = "*) Our Pancard Number is AAAAAAAAA"
'
'xlsheet.Range("A" & i + 20 & ":D" & i + 20).Font.Bold = True
'xlsheet.Range("A" & i + 20 & ":D" & i + 20).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'xlsheet.Cells(i + 20, 4).Value = "For Secretary / Principal"
'
'xlsheet.HPageBreaks.Add xlsheet.Range("A" & i + 21)
'
'xlapp.Application.Visible = True
'
'Sleep 100
'
'xlsheet.PrintPreview    '---------------Print Preview
End Sub

Private Sub BtnDelete_Click()
db.Execute "update tbl_cashsales set billcancel='Y' where billno='" & Trim(txtbillno.Text) & "'"
MsgBox "Bill Cancelled Successfully", vbInformation, "Kovai Cards"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmCCashSales.Show
End Sub

Private Sub BtnNext_Click()
If Mid(Trim(txtbillno.Text), 8, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 9, 1)) - 1
    If Val(bno) > 9 Then
        bno = "CCS0000" & bno
    Else
        bno = "CCS00000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 7, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 8, 2)) - 1
    If Val(bno) > 99 Then
        bno = "CCS000" & bno
    Else
        bno = "CCS0000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 6, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 7, 3)) - 1
    If Val(bno) > 999 Then
        bno = "CCS00" & bno
    Else
        bno = "CCS000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 5, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 6, 4)) - 1
    If Val(bno) > 9999 Then
        bno = "CCS0" & bno
    Else
        bno = "CCS00" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 6, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 5, 5)) - 1
    If Val(bno) > 99999 Then
        bno = "CCS" & bno
    Else
        bno = "CCS0" & bno
    End If
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_cashsales where billno='" & bno & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Call navigation

    BtnSave.Enabled = False
    BtnBill.Enabled = True
    BtnDelete.Enabled = True
Else
    Call BtnClear_Click
End If
End Sub

Private Sub BtnPrevious_Click()
If Mid(Trim(txtbillno.Text), 8, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 9, 1)) - 1
    If Val(bno) > 9 Then
        bno = "CCS0000" & bno
    Else
        bno = "CCS00000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 7, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 8, 2)) - 1
    If Val(bno) > 99 Then
        bno = "CCS000" & bno
    Else
        bno = "CCS0000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 6, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 7, 3)) - 1
    If Val(bno) > 999 Then
        bno = "CCS00" & bno
    Else
        bno = "CCS000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 5, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 6, 4)) - 1
    If Val(bno) > 9999 Then
        bno = "CCS0" & bno
    Else
        bno = "CCS00" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 6, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 5, 5)) - 1
    If Val(bno) > 99999 Then
        bno = "CCS" & bno
    Else
        bno = "CCS0" & bno
    End If
End If
    
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_cashsales where billno='" & bno & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Call navigation

    BtnSave.Enabled = False
    BtnBill.Enabled = True
    BtnDelete.Enabled = True
End If
End Sub

Function navigation()
txtbillno.Text = ""
cmbcid.Text = ""
cmbcname.Text = ""
txtmobno.Text = ""
txttotqty.Text = "0"
txttotamt.Text = "0"
txtstax.Text = ""
lblpayamt.Caption = "0.00"
MSGrid.Rows = 2
MSGrid.TextMatrix(1, 0) = ""
MSGrid.TextMatrix(1, 1) = ""
MSGrid.TextMatrix(1, 2) = ""
MSGrid.TextMatrix(1, 3) = ""
lblcancel.Visible = False

If rs.Fields("billcancel") = "Y" Then
    txtbillno.Text = rs.Fields("billno")
    lblcancel.Visible = True
    GoTo nxt:
End If

txtbillno.Text = rs.Fields("billno")
dtp_bdate.Value = rs.Fields("bdate")
cmbcid.Text = rs.Fields("cid")
cmbcname.Text = rs.Fields("cname")
txtmobno.Text = rs.Fields("mobno")
txttotqty.Text = IIf(IsNull(Trim(rs.Fields("totqty"))), 0, Trim(rs.Fields("totqty")))
txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
txtstax.Text = Format(rs.Fields("stax"), "0.00")
lblpayamt.Caption = Format(rs.Fields("pamt"), "0.00")

i = 1
While Not rs.EOF
    MSGrid.TextMatrix(i, 0) = rs.Fields("iname")
    MSGrid.TextMatrix(i, 1) = IIf(IsNull(rs.Fields("qty")), 0, rs.Fields("qty"))
    MSGrid.TextMatrix(i, 2) = Format(Val(rs.Fields("irate")), "0.00")
    MSGrid.TextMatrix(i, 3) = IIf(IsNull(rs.Fields("amount")), 0, Format(Val(rs.Fields("amount")), "0.00"))
    MSGrid.Rows = MSGrid.Rows + 1
    i = i + 1
    rs.MoveNext
Wend
rs.Close

nxt:

End Function

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If cmbcname.Text = "" Then
    MsgBox "Select the Customer Name Properly...", vbInformation, "Kovai Cards"
    cmbcname.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_cashsales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 1
        rs.AddNew
        rs.Fields("billno") = Trim(txtbillno.Text)
        rs.Fields("bdate") = dtp_bdate.Value
        rs.Fields("cid") = Trim(cmbcid.Text)
        rs.Fields("cname") = Trim(cmbcname.Text)
        rs.Fields("mobno") = Trim(txtmobno.Text)
        rs.Fields("sno") = i
        rs.Fields("iname") = Trim(MSGrid.TextMatrix(i, 0))
        rs.Fields("qty") = Val(MSGrid.TextMatrix(i, 1))
        rs.Fields("irate") = Format(Val(Trim(MSGrid.TextMatrix(i, 2))), "0.00")
        rs.Fields("amount") = Format(Val(Trim(MSGrid.TextMatrix(i, 3))), "0.00")
        rs.Fields("totqty") = Val(Trim(txttotqty.Text))
        rs.Fields("totamt") = Format(Val(Trim(txttotamt.Text)), "0.00")
        rs.Fields("stax") = Format(Val(Trim(txtstax.Text)), "0.00")
        rs.Fields("pamt") = Format(Val(Trim(lblpayamt.Caption)), "0.00")
        rs.Update
    Next i
    rs.Close

    MsgBox "Customer Cash Sales Saved Successfully...", vbInformation, "Kovai Cards"

    'Call BtnBill_Click
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call connect
Call Fill

Call generate_sno("CCS")    'CCS - Customer Cash Sales
txtbillno.Text = serialno

dtp_bdate.Value = Date

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i

lblcancel.Visible = False
BtnSave.Enabled = True
BtnDelete.Enabled = False
End Sub

Private Function Fill()
If rs.State = 1 Then rs.Close
rs.Open "select iname from tbl_itemmaster order by iid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        lstItemName.AddItem rs.Fields("iname")
        rs.MoveNext
    Loop
End If
rs.Close

rs.Open "select cid,cname from tbl_custmaster where ctype='Customer' order by cid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        cmbcid.AddItem rs.Fields("cid")
        cmbcname.AddItem rs.Fields("cname")
        rs.MoveNext
    Loop
End If
End Function

Private Sub cmbcid_Click()
If cmbcid.Text <> "" Then
    cmbcname.Text = cmbcname.List(cmbcid.ListIndex)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select mobno from tbl_custmaster where cid=" & Val(cmbcid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        txtmobno.Text = rs1.Fields("mobno")
    End If
    rs1.Close
    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcname_Click()
If cmbcname.Text <> "" Then
    cmbcid.Text = cmbcid.List(cmbcname.ListIndex)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select mobno from tbl_custmaster where cid=" & Val(cmbcid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        txtmobno.Text = rs1.Fields("mobno")
    End If
    rs1.Close
    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcname_LostFocus()
If cmbcname.Text <> "" Then
    cmbcid.Text = cmbcid.List(cmbcname.ListIndex)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select mobno from tbl_custmaster where cid=" & Val(cmbcid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        txtmobno.Text = rs1.Fields("mobno")
    End If
    rs1.Close
    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcname_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcid_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    lstItemName.SetFocus
End If
End Sub

Private Sub lstItemName_ItemCheck(Item As Integer)
If lstItemName.Selected(Item) = True Then
    If MSGrid.TextMatrix(1, 0) = "" Then
        MSGrid.Rows = 1
    End If

    If MSGrid.Rows = 17 Then
        lstItemName.Selected(Item) = False
        Exit Sub
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_itemmaster where iname='" & Trim(lstItemName.List(Item)) & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            MSGrid.AddItem rs.Fields("iname") & vbTab & "1" & vbTab & Format(rs.Fields("crate"), "0.00") & vbTab & Format(rs.Fields("crate"), "0.00")
        End If
    End If
Else
    For i = 1 To MSGrid.Rows - 1
        If MSGrid.TextMatrix(i, 0) = lstItemName.List(Item) Then
            If MSGrid.Rows = 2 Then
                MSGrid.TextMatrix(1, 0) = ""
                MSGrid.TextMatrix(1, 1) = ""
                MSGrid.TextMatrix(1, 2) = ""
                MSGrid.TextMatrix(1, 3) = ""
            Else
                MSGrid.RemoveItem i
            End If
            Exit For
        End If
    Next
End If
'--------------Calculating the total quantity and amount
txtstax.Text = 0
txttotamt.Text = 0
txttotqty.Text = 0
lblpayamt.Caption = 0
For i = 1 To MSGrid.Rows - 1
    txttotqty.Text = Val(txttotqty.Text) + Val(MSGrid.TextMatrix(i, 1))
    txttotamt.Text = Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 3))
Next
txtstax.Text = Format(Val(txttotamt.Text) * 15 / 100, "0.00")
txttotamt.Text = Format(Val(txttotamt.Text), "0.00")
lblpayamt.Caption = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
'--------------Calculating the total quantity and amount
End Sub

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

Private Sub MSGrid_KeyPress(KeyAscii As Integer)
If MSGrid.Col = 1 Then  'Qty coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
'    Case 32         ' 32 keyascii is for space bar key
'        If MSGrid.Col = 0 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
'    Case 46         ' 46 keyascii is for dot symbol
'        If MSGrid.Col = 4 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
'        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'    Case 97 To 122  ' 97-122 keyascii is for small a to z
'        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 1 Then  'Qty Coloumn
            If MSGrid.TextMatrix(MSGrid.Row, 1) <> "" Then
                
                MSGrid.TextMatrix(MSGrid.Row, 3) = ""
                MSGrid.TextMatrix(MSGrid.Row, 3) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 1)) * Val(MSGrid.TextMatrix(MSGrid.Row, 2)), "0.00")
                
                txttotqty.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotqty.Text = Val(txttotqty.Text) + Val(MSGrid.TextMatrix(i, 1))
                Next
                
                txttotamt.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotamt.Text = Format(Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 3)), "0.00")
                Next
                
                '--------------Calculating the total quantity and amount
                txtstax.Text = 0
                txttotamt.Text = 0
                txttotqty.Text = 0
                lblpayamt.Caption = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotqty.Text = Val(txttotqty.Text) + Val(MSGrid.TextMatrix(i, 1))
                    txttotamt.Text = Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 3))
                Next
                txtstax.Text = Format(Val(txttotamt.Text) * 15 / 100, "0.00")
                txttotamt.Text = Format(Val(txttotamt.Text), "0.00")
                lblpayamt.Caption = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
                '--------------Calculating the total quantity and amount
                
                On Error Resume Next
                'SendKeys "{DOWN}"   'For Windows 7 make your project as exe. Then right click -> propertirs
                                    'then select compatibility tab then select windows xp sp2. Now u run the exe file, it will
                                    'work properly
                Set WshShell = CreateObject("WScript.Shell")
                WshShell.SendKeys "{DOWN}"
                
            End If
        End If
    End Select
End If
End Sub

Private Sub txtstax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnSave.SetFocus
End If
End Sub

Private Sub txtstax_LostFocus()
If txtstax.Text = "" Then
    MsgBox "Sales tax should not empty", vbInformation, "Kovai Cards"
Else
    txtstax.Text = Format(Val(txtstax.Text), "0.00")
    txttotamt.Text = Format(Val(txttotamt.Text), "0.00")
    lblpayamt.Caption = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
End If
End Sub
