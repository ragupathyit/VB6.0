VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmKOTOrder 
   BackColor       =   &H00800080&
   Caption         =   "KOT / Billing - Kitchen Order Tickets"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Project1.Button BtnNext 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      ToolTipText     =   "Next"
      Top             =   120
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
      MICON           =   "FrmKOTOrder.frx":0000
      PICN            =   "FrmKOTOrder.frx":001C
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
      Left            =   9120
      TabIndex        =   5
      ToolTipText     =   "Previous"
      Top             =   120
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
      MICON           =   "FrmKOTOrder.frx":072E
      PICN            =   "FrmKOTOrder.frx":074A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txttableno 
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
      Left            =   2400
      TabIndex        =   38
      Top             =   1320
      Width           =   1935
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
      IMEMode         =   3  'DISABLE
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtvattax 
      Alignment       =   1  'Right Justify
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0"
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtservicetax 
      Alignment       =   1  'Right Justify
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
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
      ItemData        =   "FrmKOTOrder.frx":0E5C
      Left            =   705
      List            =   "FrmKOTOrder.frx":0E5E
      TabIndex        =   22
      Top             =   2565
      Visible         =   0   'False
      Width           =   6795
   End
   Begin VB.TextBox txtwaitername 
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
      Left            =   6360
      TabIndex        =   21
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
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
      Left            =   9675
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   6960
      Width           =   1750
   End
   Begin VB.TextBox txttotalqty 
      Alignment       =   1  'Right Justify
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
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtotime 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtoid 
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
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   120
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
      MICON           =   "FrmKOTOrder.frx":0E60
      PICN            =   "FrmKOTOrder.frx":0E7C
      PICH            =   "FrmKOTOrder.frx":158E
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
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "SAVE"
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Save     "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmKOTOrder.frx":1CA0
      PICN            =   "FrmKOTOrder.frx":1CBC
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
      Left            =   8400
      TabIndex        =   2
      ToolTipText     =   "DELETE"
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Delete  "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12583104
      BCOLO           =   12583104
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmKOTOrder.frx":23CE
      PICN            =   "FrmKOTOrder.frx":23EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Button BtnBill 
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "BILL"
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "   &Bill      "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmKOTOrder.frx":2AFC
      PICN            =   "FrmKOTOrder.frx":2B18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtp_odate 
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   97255427
      CurrentDate     =   42772
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid1 
      Height          =   8775
      Left            =   11760
      TabIndex        =   17
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   15478
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "T. No  |Amt (Rs)      |Order Time          "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5055
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "S.No |Item Name                                                                        |Qty    |Price(Rs)    |Amount (Rs)   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Button BtnClear 
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      ToolTipText     =   "CLEAR"
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Refresh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8388608
      BCOLO           =   8388608
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmKOTOrder.frx":2D3D
      PICN            =   "FrmKOTOrder.frx":2D59
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid MSG_Cancel 
      Height          =   7935
      Left            =   16560
      TabIndex        =   35
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   13996
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16761024
      BackColorBkg    =   16777215
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "O. No    |Amt (Rs)        "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   17640
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canceled Orders"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   16920
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label9 
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
      Left            =   8640
      TabIndex        =   33
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   9000
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2760
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   5880
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 %"
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
      Left            =   6480
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14.5 %"
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
      TabIndex        =   28
      Top             =   7680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblpayamt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   9480
      TabIndex        =   25
      Top             =   7320
      Width           =   1065
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vat Tax"
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
      Left            =   6480
      TabIndex        =   24
      Top             =   7440
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serivce Tax"
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
      TabIndex        =   23
      Top             =   7440
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waiter Name"
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
      Left            =   4560
      TabIndex        =   20
      Top             =   1440
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table No *"
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
      Left            =   840
      TabIndex        =   19
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   6480
      TabIndex        =   16
      Top             =   7080
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Time"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
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
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label lbldeleted 
      BackStyle       =   0  'Transparent
      Caption         =   "KOT ORDER CANCELED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmKOTOrder.frx":346B
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KOT - Kitchen Order Tickets"
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
      TabIndex        =   3
      Top             =   120
      Width           =   4875
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   11775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   8040
      Width           =   11775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   8775
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "FrmKOTOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<================================ Printer Code ===========================================>
Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
   pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
   pcWritten As Long) As Long
   
Private Sub BtnBill_Click()
If txttableno.Text <> "" Then
    db.Execute "update tbl_order set iscomplete=true where orderid=" & Val(txtoid.Text)
    
    If rs.State = 1 Then rs.Close
    Sql = "SELECT billno, orderid, orderdate, tableno, waitername, sno, itemname, quantity, sum(quantity) as qty, price, amount, sum(amount) as amt, totamt, servicetax, vattax, payamt From tbl_order where iscomplete=true GROUP BY billno, orderdate, orderid, tableno, waitername, sno, itemname, quantity, price, amount, totamt, servicetax, vattax, payamt HAVING billno=" & Val(txtbillno.Text) & " and orderid=" & Val(txtoid.Text)
    Debug.Print Sql
    rs.Open Sql, db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txttableno.Text = rs.Fields("tableno")
        
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select * from tbl_runorder where tableno='" & Trim(rs.Fields("tableno")) & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            rs1.Fields("payamt") = Null
            rs1.Fields("ordertime") = Null
            rs1.update
        End If
        rs1.Close
        
        '----------Notepad print------------------
        Open App.Path & "\bill.txt" For Output As #1
            Print #1, Chr(27); Chr(77);         ' Printer 12 Pitch=Chr(27)    Form feed=Chr(12); 10 pitch=Chr(18);
            Print #1, "            Sri Saravana Bhavan           "
            Print #1, "              Then Thirupathy             "
            Print #1, "             Nall Road Junction           "
            Print #1, "          Annur Road, Mettupalayam        "
            Print #1, "              TIN: 33066422947            "
            Print #1, ""
            
            billno = rs.Fields("billno")
            
            Print #1, "Bill No : " & rs.Fields("billno") & Space(6 - Len(rs.Fields("billno"))) & Space(12) & "Date: " & Format(Date, "DD/MM/YY")
            Print #1, "Table No: " & rs.Fields("tableno") & Space(4 - Len(rs.Fields("tableno"))) & Space(14) & "Time: " & Format(Time, "HH:MM AMPM")
            Print #1, "Waiter Name: " & IIf(IsNull(rs.Fields("waitername")), "", rs.Fields("waitername"))
            Print #1, "------------------------------------------"      '42 characters
            Print #1, "Item Name " & Space(12) & Space(1) & "Qty" & Space(1) & " Price" & Space(1) & "  Amount"
            Print #1, "------------------------------------------"
            
            tamt = Round(Format(rs.Fields("totamt"), "0.00")) & ".00"
            itamt = 10 - Len(Format(tamt, "0.00"))
            
            pamt = Round(Format(rs.Fields("payamt"), "0.00")) & ".00"
            ipamt = 10 - Len(Format(pamt, "0.00"))
            
            word = ConNumToEngLish(Val(pamt))
            
'            tservicetax = rs.Fields("servicetax")
'            iservicetax = 10 - Len(Format(tservicetax, "0.00"))
    
            tvattax = rs.Fields("vattax")
            ivattax = 10 - Len(Format(tvattax, "0.00"))
            
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select distinct itemcode from tbl_order where iscomplete=true and billno=" & Val(txtbillno.Text) & " and orderid=" & Val(txtoid.Text), db, adOpenDynamic, adLockOptimistic
            i = 1
            While Not rs1.EOF
                If rs2.State = 1 Then rs2.Close
                rs2.Open "select itemname,sum(quantity) as q,price,sum(amount) as a from tbl_order where itemcode=" & Trim(Val(rs1.Fields("itemcode"))) & " and iscomplete=true and billno=" & Val(txtbillno.Text) & " and orderid=" & Val(txtoid.Text) & " group by itemname,price", db, adOpenDynamic, adLockOptimistic
                If Not rs2.EOF Then
                    ii = 2 - Len(i)
                    IName = 22 - Len(Mid(rs2.Fields("itemname"), InStr(1, rs2.Fields("itemname"), "-") + 1, 22))
                    iqty = 3 - Len(rs2.Fields("q"))
                    iprice = 6 - Len(Format(rs2.Fields("price"), "0.00"))
                    iamt = 8 - Len(Format(rs2.Fields("a"), "0.00"))
        
                    Print #1, UCase(Mid(rs2.Fields("itemname"), InStr(1, rs2.Fields("itemname"), "-") + 1, 22)) & Space(IName) & Space(1) & Space(iqty) & rs2.Fields("q") & Space(1) & Space(iprice) & Format(rs2.Fields("price"), "0.00") & Space(1) & Space(iamt) & Format(rs2.Fields("a"), "0.00")
                    
                    i = i + 1
                End If
                rs1.MoveNext
            Wend
            rs1.Close
            rs2.Close
            
            Print #1, "------------------------------------------"
            Print #1, "Items: " & Val(i) - 1 & Space(ii) & Space(13) & "Total:    " & Space(itamt) & Format(tamt, "0.00")
            
'            If Val(tservicetax) <> 0 Then
'                Print #1, ""
'                Print #1, "            Service Tax:        " & Space(iservicetax) & Format(tservicetax, "0.00")
'            End If

'            If Val(tvattax) <> 0 Then
'                Print #1, "                     Vat 2%:    " & Space(ivattax) & Format(tvattax, "0.00")
'            End If
                       
            'MsgBox Val(Right(tvattax, InStr(tvattax, ".")))
            
'            If Format(Val(Right(tvattax, InStr(Format(tvattax, "0.00"), ".") - 1)), "0.00") > Val(0.19) Then
'                roff = Val(Left(tvattax, InStr(tvattax, ".") - 1)) + 1
'                roff = Val(roff) - Val(tvattax)
'                iroff = 10 - Len(Format(roff, "0.00"))
'                Print #1, "                  Round Off:    " & Space(iroff) & Format(roff, "0.00")
'            ElseIf Format(Val(Right(tvattax, InStr(tvattax, ".") - 1)), "0.00") <= Val(0.19) Then
'                roff = Val(Left(tvattax, InStr(tvattax, ".") - 1)) - 1
'                roff = Val(roff) - Val(tvattax)
'                iroff = 10 - Len(Format(roff, "0.00"))
'                Print #1, "                  Round Off:   " & Space(iroff) & Format(roff, "0.00")
'            End If
            
'            Print #1, Space(32) & "----------"
'            Print #1, Space(22) & "Total:    " & Space(ipamt) & Format(pamt, "0.00")
'            Print #1, Space(32) & "----------"
            'Print #1, word & " Rupees Only"
            'Print #1, Space(43) & "Authorized Signatory"
            
            Print #1, "         Thank You! Visit Again!          "
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Chr$(&H1B); "m"; Chr$(&HA);   'Cutter Code
        Close #1
        retval = Shell("notepad.exe bill.txt", vbHide)
    End If
    
    '<==================== Printing Code ========================>
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    MyDocInfo.pDocName = "AAAAAA"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)

    Dim var1 As String
    Open App.Path & "\bill.txt" For Input As #1
    var1 = Input(LOF(1), #1)
    Close #1

    sWrittenData = var1 '& vbFormFeed

    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    '<==================== Printing Code ========================>
    
    txttableno.Text = ""
    Call BtnClear_Click
End If
End Sub

Private Sub BtnDelete_Click()
If uname = "admin" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_order where iscomplete=false and orderid=" & Val(txtoid.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        db.Execute "update tbl_order set iscomplete=true, iscancel=true where iscomplete=false and orderid=" & Val(txtoid.Text)
        
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select * from tbl_runorder where tableno='" & Trim(txttableno.Text) & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            rs1.Fields("payamt") = Null
            rs1.Fields("ordertime") = Null
            rs1.update
        End If
        rs1.Close
        
        MsgBox "KOT Order is Canceled Successfully", vbInformation, "Sri Saravana Bhavan"
    Else
        MsgBox "KOT Order is Not Canceled. Because It is Billed.", vbInformation, "Sri Saravana Bhavan"
    End If
End If
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmKOTOrder.Show
End Sub

Private Sub BtnNext_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_order where orderid=" & Val(txtoid.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    lbldeleted.Visible = False
    txtoid.Text = ""
    txtotime.Text = ""
    txtbillno.Text = ""
    txttableno.Text = ""
    txtwaitername.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""

    If rs1.Fields("iscancel") = True Then
        txtoid.Text = rs1.Fields("orderid")
        lbldeleted.Visible = True
        
        txttableno.SetFocus
        txttableno.SelStart = 0
        txttableno.SelLength = Len(txttableno.Text)    'select the text
        
        BtnSave.Enabled = False
        BtnDelete.Enabled = True
        Exit Sub
    Else
        txtoid.Text = rs1.Fields("orderid")
    End If
    txtoid.Text = rs1.Fields("orderid")
    dtp_odate.Value = rs1.Fields("orderdate")
    txtotime.Text = rs1.Fields("ordertime")
    txtbillno.Text = rs1.Fields("billno")
    txttableno.Text = rs1.Fields("tableno")
    txtwaitername.Text = IIf(IsNull(rs1.Fields("waitername")), "", rs1.Fields("waitername"))
    txttotalqty.Text = rs1.Fields("totqty")
    txttotamt.Text = Format(Val(rs1.Fields("totamt")), "0.00")
    'txtservicetax.Text = Format(Val(rs1.Fields("servicetax")), "0.00")
    'txtvattax.Text = Format(Val(rs1.Fields("vattax")), "0.00")
    lblpayamt.Caption = Format(Val(rs1.Fields("payamt")), "0.00")

    i = 1
    While Not rs1.EOF
        MSGrid.TextMatrix(i, 0) = Trim(i)
        MSGrid.TextMatrix(i, 1) = Trim(rs1.Fields("itemname"))
        MSGrid.TextMatrix(i, 2) = Trim(rs1.Fields("quantity"))
        MSGrid.TextMatrix(i, 3) = Format(Val(rs1.Fields("price")), "0.00")
        MSGrid.TextMatrix(i, 4) = Format(Val(rs1.Fields("amount")), "0.00")
        MSGrid.Rows = MSGrid.Rows + 1
        i = i + 1
        rs1.MoveNext
    Wend
    
    txttableno.SetFocus
    txttableno.SelStart = 0
    txttableno.SelLength = Len(txttableno.Text)    'select the text
    MSGrid.Enabled = False
    
    BtnSave.Enabled = False
    BtnDelete.Enabled = True
Else
    'MSGrid.Enabled = True
    Call BtnClear_Click
End If
    
End Sub

Private Sub BtnPrevious_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_order where orderid=" & Val(txtoid.Text) - 1 & " order by orderid", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    lbldeleted.Visible = False
    txtoid.Text = ""
    txtotime.Text = ""
    txtbillno.Text = ""
    txttableno.Text = ""
    txtwaitername.Text = ""
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""

    If rs1.Fields("iscancel") = True Then
        txtoid.Text = rs1.Fields("orderid")
        lbldeleted.Visible = True
        
        txttableno.SetFocus
        txttableno.SelStart = 0
        txttableno.SelLength = Len(txttableno.Text)    'select the text
        'MSGrid.Enabled = False
        
        BtnSave.Enabled = False
        BtnDelete.Enabled = True
    Else
        txtoid.Text = rs1.Fields("orderid")
    End If
    txtoid.Text = rs1.Fields("orderid")
    dtp_odate.Value = rs1.Fields("orderdate")
    txtotime.Text = rs1.Fields("ordertime")
    txtbillno.Text = rs1.Fields("billno")
    txttableno.Text = rs1.Fields("tableno")
    txtwaitername.Text = IIf(IsNull(rs1.Fields("waitername")), "", rs1.Fields("waitername"))
    txttotalqty.Text = rs1.Fields("totqty")
    txttotamt.Text = Format(Val(rs1.Fields("totamt")), "0.00")
    'txtservicetax.Text = Format(Val(rs1.Fields("servicetax")), "0.00")
    'txtvattax.Text = Format(Val(rs1.Fields("vattax")), "0.00")
    lblpayamt.Caption = Format(Val(rs1.Fields("payamt")), "0.00")

    i = 1
    While Not rs1.EOF
        MSGrid.TextMatrix(i, 0) = Trim(i)
        MSGrid.TextMatrix(i, 1) = Trim(rs1.Fields("itemname"))
        MSGrid.TextMatrix(i, 2) = Trim(rs1.Fields("quantity"))
        MSGrid.TextMatrix(i, 3) = Format(Val(rs1.Fields("price")), "0.00")
        MSGrid.TextMatrix(i, 4) = Format(Val(rs1.Fields("amount")), "0.00")
        MSGrid.Rows = MSGrid.Rows + 1
        i = i + 1
        rs1.MoveNext
    Wend
    
    txttableno.SetFocus
    txttableno.SelStart = 0
    txttableno.SelLength = Len(txttableno.Text)    'select the text
    MSGrid.Enabled = False
    
    BtnSave.Enabled = False
    BtnDelete.Enabled = True
End If
End Sub

Function update()
'-------------------Validation Starts Here-----------------------------
If txttableno.Text = "" Then
    MsgBox "Select the Table Name Properly...", vbInformation, "Sri Saravana Bhavan"
    txttableno.SetFocus
    txttableno.SelStart = 0
    txttableno.SelLength = Len(txttableno.Text)    'select the text
'ElseIf txtwaitername.Text = "" Then
'    MsgBox "Select the Waiter Name Properly...", vbInformation, "Sri Saravana Bhavan"
'    txtwaitername.SetFocus
ElseIf MSGrid.TextMatrix(1, 1) = "" Then
    MsgBox "Enter the Item Name Properly...", vbInformation, "Sri Saravana Bhavan"
    txttableno.SetFocus
    txttableno.SelStart = 0
    txttableno.SelLength = Len(txttableno.Text)    'select the text
Else
'-------------------Validation Ends Here-------------------------------
    
    On Error Resume Next
    db.Execute "delete from tbl_order where orderid=" & Val(txtoid.Text)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_order", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 1
        If Not MSGrid.TextMatrix(i, 1) = "" Then
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select tableid from tbl_tablemaster where tableno='" & Trim(txttableno.Text) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                tid = rs1.Fields("tableid")
            End If
            
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select waiterid from tbl_waitermaster where waitername='" & Trim(txtwaitername.Text) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                wid = rs1.Fields("waiterid")
            End If
            
            If rs1.State = 1 Then rs1.Close
            'rs1.Open "select itemcode from tbl_itemmaster where itemname='" & UCase(Trim(MSGrid.TextMatrix(i, 1))) & "'", db, adOpenDynamic, adLockOptimistic
            'MsgBox UCase(Mid(Trim(MSGrid.TextMatrix(i, 1)), 1, InStr(1, Trim(MSGrid.TextMatrix(i, 1)), "-") - 1))
            rs1.Open "select itemcode from tbl_itemmaster where itemid='" & UCase(Mid(Trim(MSGrid.TextMatrix(i, 1)), 1, InStr(1, Trim(MSGrid.TextMatrix(i, 1)), "-") - 1)) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                ic = rs1.Fields("itemcode")
            End If
            
            rs.AddNew
            rs.Fields("orderid") = Val(Trim(txtoid.Text))
            rs.Fields("orderdate") = Format(Trim(dtp_odate.Value), "mm/dd/yyyy")    'Trim(dtp_odate.Value)
            rs.Fields("ordertime") = Trim(txtotime.Text)
            rs.Fields("billno") = Trim(txtbillno.Text)
            rs.Fields("tableid") = tid
            rs.Fields("tableno") = UCase(Trim(txttableno.Text))
            rs.Fields("waiterid") = wid
            rs.Fields("waitername") = UCase(Trim(txtwaitername))
            rs.Fields("sno") = Val(Trim(MSGrid.TextMatrix(i, 0)))
            rs.Fields("itemcode") = ic
            rs.Fields("itemname") = UCase(Trim(MSGrid.TextMatrix(i, 1)))
            rs.Fields("quantity") = Val(Trim(MSGrid.TextMatrix(i, 2)))
            rs.Fields("price") = Format(Val(Trim(MSGrid.TextMatrix(i, 3))), "0.00")
            rs.Fields("amount") = Format(Val(Trim(MSGrid.TextMatrix(i, 4))), "0.00")
            rs.Fields("totqty") = Val(Trim(txttotalqty.Text))
            rs.Fields("totamt") = Format(Val(Trim(txttotamt.Text)), "0.00")
            'rs.Fields("servicetax") = Format(Val(Trim(txtservicetax.Text)), "0.00")
            'rs.Fields("vattax") = Format(Val(Trim(txtvattax.Text)), "0.00")
            rs.Fields("payamt") = Format(Val(Trim(lblpayamt.Caption)), "0.00")
            rs.update
        End If
    Next i
    rs.Close
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_runorder where tableno='" & UCase(Trim(txttableno.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.Fields("payamt") = Format(Val(Trim(lblpayamt.Caption)), "0.00")
        rs.Fields("ordertime") = Trim(txtotime.Text)
        rs.update
    End If

    'db.Execute "delete from tbl_tempbill where billno=" & Val(txtbillno.Text)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_tempbill order by billno", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.AddNew
        rs.Fields("bdate") = Format(Trim(dtp_odate.Value), "mm/dd/yyyy")
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.update
    Else
        rs.AddNew
        rs.Fields("bdate") = Format(Trim(dtp_odate.Value), "mm/dd/yyyy")
        rs.Fields("billno") = 1
        rs.update
    End If
    
    'MsgBox "Items Saved Successfully", vbInformation, "Sri Saravana Bhavan"
End If
End Function

Private Sub BtnSave_Click()
Call update
Call BtnClear_Click

Call fillrorders
txttableno.SetFocus
txttableno.SelStart = 0
txttableno.SelLength = Len(txttableno.Text)    'select the text
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call connect

If rs.State = 1 Then rs.Close
rs.Open "select distinct orderid from tbl_order order by orderid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtoid.Text = Val(rs.Fields("orderid")) + 1
Else
    txtoid.Text = 1
End If

dtp_odate.Value = Date
txtotime.Text = TimeValue(Now)
txttableno.Text = ""

'---------------------------------This Line is Important-------------------------------------------------
If rs.State = 1 Then rs.Close
rs.Open "select billno from tbl_tempbill where bdate='" & Format(Date, "mm/dd/yyyy") & "' order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = Val(rs.Fields("billno")) + 1
Else
    db.Execute "Delete from tbl_tempbill"
    txtbillno.Text = 1
End If
rs.Close
'---------------------------------This Line is Important-------------------------------------------------

Call fillrorders
'Call fillcancelorders

'If rs.State = 1 Then rs.Close
'rs.Open "select itemname from tbl_itemmaster order by itemcode", db, adOpenDynamic, adLockOptimistic
'While Not rs.EOF
'    List1.AddItem rs.Fields("itemname")
'    rs.MoveNext
'Wend
'rs.Close

For i = 0 To MSGrid.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid.Row = 0
    MSGrid.Col = i
    MSGrid.CellAlignment = flexAlignCenterCenter
    MSGrid.CellFontBold = True
    'MSGrid.CellBackColor = vbWhite
Next i
MSGrid.TextMatrix(1, 0) = "1"

For i = 0 To MSGrid1.Cols - 1    ' Grid First Row all columns in center wiht bold
    MSGrid1.Row = 0
    MSGrid1.Col = i
    MSGrid1.CellAlignment = flexAlignCenterCenter
    MSGrid1.CellFontBold = True
    'MSGrid1.CellBackColor = vbWhite
Next i

Me.Show
txttableno.SetFocus
txttableno.SelStart = 0
txttableno.SelLength = Len(txttableno.Text)    'select the text

BtnSave.Enabled = True
BtnDelete.Enabled = False
End Sub

'Public Sub fillcancelorders()
'If rs.State = 1 Then rs.Close
'rs.Open "select distinct orderid, payamt from tbl_order where iscancel=true and orderdate=#" & Format(Now, "m/d/yyyy") & "#", db, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'    MSG_Cancel.Rows = 1
'    While Not rs.EOF
'        MSG_Cancel.AddItem rs.Fields("orderid") & vbTab & Format(rs.Fields("payamt"), "0.00")
'        rs.MoveNext
'    Wend
'End If
'End Sub

Public Sub fillrorders()
MSGrid1.Rows = 1
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_runorder order by tableid", db, adOpenDynamic, adLockOptimistic
While Not rs1.EOF
    If IsNull(rs1.Fields("payamt")) Then
        MSGrid1.AddItem rs1.Fields("tableno")
    Else
        MSGrid1.AddItem rs1.Fields("tableno") & vbTab & Format(Val(rs1.Fields("payamt")), "0.00") & vbTab & rs1.Fields("ordertime")
    End If
    rs1.MoveNext
Wend
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MSGrid.TextMatrix(MSGrid.Row, 1) = List1.List(List1.ListIndex)
    
    If rs.State = 1 Then rs.Close
    'rs.Open "select price from tbl_itemmaster where itemname='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 1)) & "'", db, adOpenDynamic, adLockOptimistic
    'MsgBox Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), 1, Val(InStr(1, Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), "-")) - 1)
    rs.Open "select price from tbl_itemmaster where itemid='" & Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), 1, Val(InStr(1, Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), "-")) - 1) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MSGrid.TextMatrix(MSGrid.Row, 3) = Format(Val(rs.Fields("price")), "0.00")
    End If
            
    List1.Visible = False
    MSGrid.Col = 2
    MSGrid.SetFocus
End If
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

If MSGrid.Col = 1 Then
    List1.Visible = False
End If
End Sub

Private Sub MSGrid_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
If MSGrid.Col = 1 Or MSGrid.Col = 2 Then  'Itemname and Qty coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
        
        If MSGrid.Col = 1 Then
            Me.List1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top + Me.MSGrid.CellHeight
            Me.List1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
            List1.Visible = True
            
            stmt = "select itemid,itemname from tbl_itemmaster where itemid like'" & Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) & "%' order by itemid"
            If rs1.State = 1 Then rs1.Close
            rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
            List1.Clear
            If Not rs1.EOF Then
                rs1.MoveFirst
                Do While Not rs1.EOF
                    'List1.AddItem rs1.Fields("itemname")
                    List1.AddItem rs1.Fields("itemid") & "-" & rs1.Fields("itemname")
                    rs1.MoveNext
                Loop
            End If
            rs1.Close
                        
            If Not List1.ListCount = 0 Then
                List1.ListIndex = 0
            End If
        End If
    Case 32         ' 32 keyascii is for space bar key
        If MSGrid.Col = 1 Then
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
            
            Me.List1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top + Me.MSGrid.CellHeight
            Me.List1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
            List1.Visible = True
            
            stmt = "select itemid,itemname from tbl_itemmaster where itemid like'" & Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) & "%' order by itemid"
            If rs1.State = 1 Then rs1.Close
            rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
            List1.Clear
            If Not rs1.EOF Then
                rs1.MoveFirst
                Do While Not rs1.EOF
                    'List1.AddItem rs1.Fields("itemname")
                    List1.AddItem rs1.Fields("itemid") & "-" & rs1.Fields("itemname")
                    rs1.MoveNext
                Loop
            End If
            rs1.Close
            
            If Not List1.ListCount = 0 Then
                List1.ListIndex = 0
            End If
        End If
'    Case 46         ' 46 keyascii is for dot symbol
'        If MSGrid.Col = 2 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        If MSGrid.Col = 2 Then
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
        If MSGrid.Col = 1 Then
            If KeyAscii = 88 Then ' This keyascii is for X. To print the bill
                Call update
                Call BtnBill_Click
            Else
                MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
                
                Me.List1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top + Me.MSGrid.CellHeight
                Me.List1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
                List1.Visible = True
                
                stmt = "select itemid,itemname from tbl_itemmaster where itemid like'" & Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) & "%' order by itemid"
                If rs1.State = 1 Then rs1.Close
                rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
                List1.Clear
                If Not rs1.EOF Then
                    rs1.MoveFirst
                    Do While Not rs1.EOF
                        'List1.AddItem rs1.Fields("itemname")
                        List1.AddItem rs1.Fields("itemid") & "-" & rs1.Fields("itemname")
                        rs1.MoveNext
                    Loop
                End If
                rs1.Close
                
                If Not List1.ListCount = 0 Then
                    List1.ListIndex = 0
                End If
            End If
        End If
    Case 97 To 122  ' 97-122 keyascii is for small a to z
        If MSGrid.Col = 1 Then
            If KeyAscii = 120 Then ' This keyascii is for x. To print the bill
                Call update
                Call BtnBill_Click
            Else
                MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
                
                Me.List1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top + Me.MSGrid.CellHeight
                Me.List1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
                List1.Visible = True
                
                stmt = "select itemid,itemname from tbl_itemmaster where itemid like'" & Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) & "%' order by itemid"
                If rs1.State = 1 Then rs1.Close
                rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
                List1.Clear
                If Not rs1.EOF Then
                    rs1.MoveFirst
                    Do While Not rs1.EOF
                        'List1.AddItem rs1.Fields("itemname")
                        List1.AddItem rs1.Fields("itemid") & "-" & rs1.Fields("itemname")
                        rs1.MoveNext
                    Loop
                End If
                rs1.Close
                
                If Not List1.ListCount = 0 Then
                    List1.ListIndex = 0
                End If
            End If
        End If
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 1 Then  'Itemname Coloumn
            If MSGrid.TextMatrix(MSGrid.Row, 1) <> "" Then
                If Not List1.ListCount = 0 Then
                    MSGrid.TextMatrix(MSGrid.Row, 1) = List1.List(List1.ListIndex)
                    
                    If rs.State = 1 Then rs.Close
                    'rs.Open "select price from tbl_itemmaster where itemname='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 1)) & "'", db, adOpenDynamic, adLockOptimistic
                    'MsgBox Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), 1, Val(InStr(1, Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), "-")) - 1)
                    rs.Open "select price from tbl_itemmaster where itemid='" & Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), 1, Val(InStr(1, Trim(MSGrid.TextMatrix(MSGrid.Row, 1)), "-")) - 1) & "'", db, adOpenDynamic, adLockOptimistic
                    If Not rs.EOF Then
                        MSGrid.TextMatrix(MSGrid.Row, 3) = Format(Val(rs.Fields("price")), "0.00")
                    End If
                    
                    MSGrid.Col = MSGrid.Col + 1  ' Grid entry was changed to Qty coloumn
                Else
                    MSGrid.TextMatrix(MSGrid.Row, 1) = ""
                    MSGrid.Col = 1
                    List1.Visible = False
                End If
            Else
                List1.Visible = False
                MSGrid.CellBackColor = vbWhite
                'BtnSave.SetFocus   'cursor navigation to the BtnSave Button
                If BtnSave.Enabled = False Then
                    MsgBox "You cannot modify/edit this bill", vbInformation, "Sri Saravana Bhavan"
                Else
                    Call BtnSave_Click
                End If
            End If
        End If

        If MSGrid.Col = 2 Then  'Qty Coloumn
            'List1.Visible = False
'            If rs.State = 1 Then rs.Close
'            rs.Open "select price from tbl_itemmaster where itemname='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 1)) & "'", db, adOpenDynamic, adLockOptimistic
'            If Not rs.EOF Then
'                MSGrid.TextMatrix(MSGrid.Row, 3) = Format(Val(rs.Fields("price")), "0.00")
'            End If
            If MSGrid.TextMatrix(MSGrid.Row, 2) = "" And flag = 1 Then
                MSGrid.TextMatrix(MSGrid.Row, 2) = 1
                flag = 0
            Else
                flag = 1
            End If
            
            If MSGrid.TextMatrix(MSGrid.Row, 2) <> "" Then
                flag = 0
                MSGrid.TextMatrix(MSGrid.Row, 4) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 2)) * Val(MSGrid.TextMatrix(MSGrid.Row, 3)), "0.00")
                            
                txttotalqty.Text = 0
                txttotamt.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txttotalqty.Text = Val(txttotalqty.Text) + Val(MSGrid.TextMatrix(i, 2))   'Grid Total Quantity calculation
                    txttotamt.Text = Format(Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 4)), "0.00")   'Grid Total Amount calculation
                Next i
                
                'txtservicetax.Text = Format(Val(txttotamt.Text) * 14.5 / 100, "0.00")
                
                'vtax = Format(Val(txttotamt.Text) * 2 / 100, "0.00")
                'txtvattax.Text = ""
                'txtvattax.Text = Val(vtax)
                
                '====================================This line is for Round off the value===========================
                'MsgBox Right(vtax, InStr(vtax, "."))
                'If Val(Right(vtax, InStr(vtax, "."))) > 19 Then
                '    taxv = Val(Left(vtax, InStr(vtax, ".") - 1)) + 1
                'ElseIf Val(Right(vtax, InStr(vtax, "."))) <= 19 Then
                '    taxv = Val(Left(vtax, InStr(vtax, ".") - 1))
                'End If
                '====================================This line is for Round off the value===========================
                
                'txtvattax.Text = Format(Val(txtvattax.Text), "0.00")
                'lblpayamt.Caption = Format(Round(Val(txttotamt.Text) + Val(txtservicetax.Text) + Val(txtvattax.Text), 0), "0.00")
                'lblpayamt.Caption = Format(Round(Val(txttotamt.Text) + Val(taxv), 0), "0.00")
                lblpayamt.Caption = Format(Round(Val(txttotamt.Text), 0), "0.00")
                
                If MSGrid.TextMatrix(MSGrid.Rows - 1, 1) = "" Then
                    MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                End If
                
                MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                MSGrid.Row = MSGrid.Rows - 1     'cursor position changed to the newlly created row
                MSGrid.Col = 1                  'cursor position changed to the first coloumn of that newly created row
                MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = MSGrid.Rows - 1
                
                On Error Resume Next
                'SendKeys "{DOWN}"   'For Windows 7 make your project as exe. Then right click -> propertirs
                                    'then select compatibility tab then select windows xp sp2. Now u run the exe file, it will
                Set WshShell = CreateObject("WScript.Shell")
                WshShell.SendKeys "{DOWN}"                    'work properly
            End If
            
        End If
'    Case 120    ' x
'        Call BtnBill_Click
'    Case 88     ' X
'        Call BtnBill_Click
    End Select
End If
End Sub

Private Sub MSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
'If vbKeyDown = True Then
'    MsgBox "hai"
'End If
If KeyCode = 17 Then    'F1 key
    List1.Visible = True
    List1.SetFocus
End If
If KeyCode = 117 Then 'F6 Key for Delete the row
    If Not MSGrid.Rows = 1 Then
        If MSGrid.Row = 1 Then
            MSGrid.TextMatrix(1, 0) = ""
            MSGrid.TextMatrix(1, 1) = ""
            MSGrid.TextMatrix(1, 2) = ""
            MSGrid.TextMatrix(1, 3) = ""
            MSGrid.TextMatrix(1, 4) = ""
            'MSGrid.RemoveItem MSGrid.Row
        Else
            If MSGrid.TextMatrix(MSGrid.Row, 0) <> "" Then
                MSGrid.RemoveItem MSGrid.Row
            End If
        End If
        
        For i = 1 To MSGrid.Rows - 1
            MSGrid.TextMatrix(i, 0) = i
        Next i
        
        '----------------------------------Calculation Starts Here-------------------------------------
        txttotalqty.Text = 0
        txttotamt.Text = 0
        For i = 1 To MSGrid.Rows - 1
            txttotalqty.Text = Val(txttotalqty.Text) + Val(MSGrid.TextMatrix(i, 2))   'Grid Total Quantity calculation
            txttotamt.Text = Format(Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 4)), "0.00")   'Grid Total Amount calculation
        Next i
        
        'txtservicetax.Text = Format(Val(txttotamt.Text) * 14.5 / 100, "0.00")
        
        'vtax = Format(Val(txttotamt.Text) * 2 / 100, "0.00")
        'txtvattax.Text = ""
        'txtvattax.Text = Val(vtax)
        
        '====================================This line is for Round off the value===========================
        'MsgBox Left(vtax, InStr(vtax, ".") - 1)
        'If Val(Right(vtax, InStr(vtax, "."))) > 19 Then
        '    taxv = Val(Left(vtax, InStr(vtax, ".") - 1)) + 1
        'ElseIf Val(Right(vtax, InStr(vtax, "."))) <= 19 Then
        '    taxv = Val(Left(vtax, InStr(vtax, ".") - 1))
        'End If
        '====================================This line is for Round off the value===========================
        
        'txtvattax.Text = Format(Val(txtvattax.Text), "0.00")
        'lblpayamt.Caption = Format(Round(Val(txttotamt.Text) + Val(txtservicetax.Text) + Val(txtvattax.Text), 0), "0.00")
        'lblpayamt.Caption = Format(Round(Val(txttotamt.Text) + Val(taxv), 0), "0.00")
        lblpayamt.Caption = Format(Round(Val(txttotamt.Text), 0), "0.00")
        '----------------------------------Calculation Ends Here--------------------------------------------
        
        MSGrid.Row = MSGrid.Rows - 1
        MSGrid.Col = 1
        MSGrid.CellBackColor = RGB(117, 145, 233)
    End If
End If
End Sub

Private Sub MSGrid1_Click()

'txtwaitername.Text = ""

If MSGrid1.TextMatrix(MSGrid1.Row, 2) <> "" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_order where tableno='" & MSGrid1.TextMatrix(MSGrid1.Row, 0) & "' and iscomplete=false order by sno", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txtoid.Text = rs.Fields("orderid")
        dtp_odate.Value = rs.Fields("orderdate")
        txtotime.Text = rs.Fields("ordertime")
        txtbillno.Text = rs.Fields("billno")
        txttableno.Text = rs.Fields("tableno")
        'txtwaitername.Text = rs.Fields("waitername")
        txttotalqty.Text = rs.Fields("totqty")
        txttotamt.Text = Format(Val(rs.Fields("totamt")), "0.00")
        'txtservicetax.Text = Format(Val(rs.Fields("servicetax")), "0.00")
        'txtvattax.Text = Format(Val(rs.Fields("vattax")), "0.00")
        lblpayamt.Caption = Format(Val(rs.Fields("payamt")), "0.00")
        
        MSGrid.Rows = 1
        While Not rs.EOF
            MSGrid.AddItem rs.Fields("sno") & vbTab & rs.Fields("itemname") & vbTab & rs.Fields("quantity") & vbTab & Format(Val(rs.Fields("price")), "0.00") & vbTab & Format(Val(rs.Fields("amount")), "0.00")
            rs.MoveNext
        Wend
        
        MSGrid.Rows = MSGrid.Rows + 1
        MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = MSGrid.Rows - 1
    End If
Else
    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = "1"
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""
End If
BtnDelete.Enabled = True

MSGrid.Row = MSGrid.Rows - 1
MSGrid.Col = 1  ' Grid entry focused to itemname coloumn
MSGrid.CellBackColor = RGB(117, 145, 233)
MSGrid.SetFocus
End Sub

Function print_token()
If txttableno.Text <> "" Then
    If rs.State = 1 Then rs.Close
    Sql = "SELECT orderid, orderdate, ordertime, tableno, itemname, quantity From tbl_order where orderid=" & Val(txtoid.Text) & " and iscomplete=false"
    Debug.Print Sql
    rs.Open Sql, db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        txttableno.Text = rs.Fields("tableno")
        '----------Notepad print------------------
        Open App.Path & "\token.txt" For Output As #1
            'Print #1, Chr(18); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
            Print #1, "                    KOT                   "
            Print #1, "            Sri Saravana Bhavan           "
            Print #1, ""
            Print #1, "                             Table No: " & rs.Fields("tableno")
            Print #1, "------------------------------------------" '42 characters
            'Print #1, "Table No  : " & rs.Fields("tableno") & Space(4 - Len(rs.Fields("tableno"))) & Space(10) & "KOT Order No: " & rs.Fields("orderid")
            Print #1, "KOT Order No: " & rs.Fields("orderid")
            Print #1, "Order Date: " & Format(rs.Fields("orderdate"), "dd/mm/yy") & Space(8 - Len(Format(rs.Fields("orderdate"), "dd/mm/yy"))) & Space(8) & "Time: " & Format(rs.Fields("ordertime"), "HH:MM AMPM")
            Print #1, ""
            Print #1, "------------------------------------------"
            Print #1, "Item Name " & Space(26) & Space(1) & " Qty"
            Print #1, "------------------------------------------"
            While Not rs.EOF
                If rs1.State = 1 Then rs1.Close
                rs1.Open "select itemtype from tbl_itemmaster where itemname='" & Trim(rs.Fields("itemname")) & "' and itemtype='Chinese'", db, adOpenDynamic, adLockOptimistic
                If Not rs1.EOF Then
                    IName = 36 - Len(Mid(rs.Fields("itemname"), 1, 35))
                    iqty = 4 - Len(rs.Fields("quantity"))
                    
                    Print #1, UCase(Mid(rs.Fields("itemname"), 1, 35)) & Space(IName) & Space(1) & Space(iqty) & rs.Fields("quantity")
                End If
                rs.MoveNext
            Wend
            Print #1, "------------------------------------------"
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Chr$(&H1B); "m"; Chr$(&HA);   'Cutter Code
        Close #1
        retval = Shell("notepad.exe token.txt", vbHide)
    End If
    
'    Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'    Print #1, "start DOSPrinter.exe /ESC E /F'Lucida Console' token.txt"
'    Close #1
'    retval = Shell(App.Path & "\print.bat", vbHide)
    '<==================== Printing Code ========================>
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Function
    End If
    MyDocInfo.pDocName = "AAAAAA"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)

    Dim var1 As String
    Open App.Path & "\token.txt" For Input As #1
    var1 = Input(LOF(1), #1)
    Close #1
    
    sWrittenData = var1 '& vbFormFeed

    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    '<==================== Printing Code ========================>
    
    txttableno.Text = ""
    txttableno.SetFocus
    txttableno.SelStart = 0
    txttableno.SelLength = Len(txttableno.Text)    'select the text
End If
End Function

'Private Sub txttableno_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox KeyCode
''If KeyCode = 83 Then    's or S to save the order
''    Call BtnSave_Click
''End If
''tno = Trim(txttableno.Text)
'If KeyCode = 80 Then    'p or P to save the order
'    Call BtnBill_Click
'End If
'If KeyCode = 84 Then    'T or t to take KOT to kitched order for chinese
'    Call print_token
'End If
'End Sub

Private Sub txttableno_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
If KeyAscii = 13 Then
    If rs2.State = 1 Then rs2.Close
    txttableno.Text = Trim(UCase(txttableno.Text))
    rs2.Open "select tableno from tbl_tablemaster where tableno='" & Trim(UCase(txttableno.Text)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs2.EOF Then
        Call fillrec
        BtnDelete.Enabled = True
    Else
        MsgBox "Enter The Correct Table No", vbInformation, "Sri Saravana Bhavan"
        txttableno.Text = ""
        txttableno.SetFocus
        
        MSGrid.Rows = 2
        MSGrid.TextMatrix(1, 0) = "1"
        MSGrid.TextMatrix(1, 1) = ""
        MSGrid.TextMatrix(1, 2) = ""
        MSGrid.TextMatrix(1, 3) = ""
        MSGrid.TextMatrix(1, 4) = ""
    End If
    rs2.Close
End If

'If KeyAscii = 120 Or KeyAscii = 88 Then 'x=120 or X=88 for Printing Bill
'    Call BtnBill_Click
'    txttableno.Text = ""
'End If

'If KeyAscii = 107 Or KeyAscii = 75 Then 'k or K for Printing Token
'    Call print_token
'End If
End Sub

Function fillrec()
If rs.State = 1 Then rs.Close
rs.Open "select distinct orderid,orderdate,ordertime,billno,tableno,waitername,totqty,totamt,payamt,itemcode from tbl_order where tableno='" & Trim(txttableno.Text) & "' and iscomplete=false", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    txtoid.Text = rs.Fields("orderid")
    dtp_odate.Value = rs.Fields("orderdate")
    txtotime.Text = rs.Fields("ordertime")
    txtbillno.Text = rs.Fields("billno")
    txttableno.Text = rs.Fields("tableno")
    txtwaitername.Text = IIf(IsNull(rs.Fields("waitername")), "", rs.Fields("waitername"))
    txttotalqty.Text = rs.Fields("totqty")
    txttotamt.Text = Format(Val(rs.Fields("totamt")), "0.00")
    '.Text = Format(Val(rs.Fields("servicetax")), "0.00")
    'txtvattax.Text = Format(Val(rs.Fields("vattax")), "0.00")
    lblpayamt.Caption = Format(Val(rs.Fields("payamt")), "0.00")
    
    i = 1
    MSGrid.Rows = 1
    While Not rs.EOF
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select itemname,sum(quantity) as q,price,sum(amount) as a from tbl_order where tableno='" & Trim(txttableno.Text) & "' and itemcode=" & Trim(Val(rs.Fields("itemcode"))) & " and iscomplete=false group by itemname,price", db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            MSGrid.AddItem i & vbTab & rs1.Fields("itemname") & vbTab & rs1.Fields("q") & vbTab & Format(Val(rs1.Fields("price")), "0.00") & vbTab & Format(Val(rs1.Fields("a")), "0.00")
        End If
        i = i + 1
        rs.MoveNext
    Wend
    
    MSGrid.Rows = MSGrid.Rows + 1
    MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = MSGrid.Rows - 1
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select orderid from tbl_order order by orderid", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveLast
        txtoid.Text = rs.Fields("orderid") + 1
    Else
        txtoid.Text = 1
    End If
    
    Call fillrorders

    If rs.State = 1 Then rs.Close
    rs.Open "select itemname from tbl_itemmaster order by itemcode", db, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        List1.AddItem rs.Fields("itemname")
        rs.MoveNext
    Wend
    rs.Close

    dtp_odate.Value = Date
    txtotime.Text = TimeValue(Now)

    MSGrid.Rows = 2
    MSGrid.TextMatrix(1, 0) = ""
    MSGrid.TextMatrix(1, 1) = ""
    MSGrid.TextMatrix(1, 2) = ""
    MSGrid.TextMatrix(1, 3) = ""
    MSGrid.TextMatrix(1, 4) = ""

    For i = 1 To MSGrid.Rows - 1
        MSGrid.TextMatrix(i, 0) = i
    Next i
        
End If

MSGrid.Row = MSGrid.Rows - 1
MSGrid.Col = 1  ' Grid entry focused to itemname coloumn
MSGrid.CellBackColor = RGB(117, 145, 233)
On Error Resume Next
MSGrid.SetFocus
End Function

Private Sub txttableno_LostFocus()
If rs3.State = 1 Then rs3.Close
rs3.Open "select waitername from tbl_waitermaster where tableno='" & Trim(UCase(txttableno.Text)) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs3.EOF Then
    txtwaitername.Text = ""
    txtwaitername.Text = Trim(rs3.Fields("waitername"))
End If
rs3.Close
End Sub
