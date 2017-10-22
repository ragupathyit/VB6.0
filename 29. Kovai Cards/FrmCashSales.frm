VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCashSales 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Cash Sales"
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
   Begin VB.TextBox txtitemsearch 
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
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtgst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   12000
      TabIndex        =   28
      Text            =   "5"
      Top             =   7560
      Width           =   735
   End
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
      Left            =   12120
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txttotqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtstax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   7200
      Width           =   1575
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   14880
      TabIndex        =   10
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
      MICON           =   "FrmCashSales.frx":0000
      PICN            =   "FrmCashSales.frx":001C
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
      TabIndex        =   5
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
      MICON           =   "FrmCashSales.frx":072E
      PICN            =   "FrmCashSales.frx":074A
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
      Left            =   7680
      TabIndex        =   8
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
      MICON           =   "FrmCashSales.frx":0E5C
      PICN            =   "FrmCashSales.frx":0E78
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
      Left            =   5760
      TabIndex        =   7
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
      MICON           =   "FrmCashSales.frx":158A
      PICN            =   "FrmCashSales.frx":15A6
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
      Left            =   8880
      TabIndex        =   13
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
      MICON           =   "FrmCashSales.frx":1CB8
      PICN            =   "FrmCashSales.frx":1CD4
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
      Left            =   10560
      TabIndex        =   14
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
      MICON           =   "FrmCashSales.frx":23E6
      PICN            =   "FrmCashSales.frx":2402
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
      Height          =   5430
      ItemData        =   "FrmCashSales.frx":2B14
      Left            =   120
      List            =   "FrmCashSales.frx":2B16
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1800
      Width           =   4568
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5775
      Left            =   4680
      TabIndex        =   4
      Top             =   1440
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "Item Name                                                                             |Qty    |I. Amt      |Amount          "
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
      Left            =   3840
      TabIndex        =   6
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
      MICON           =   "FrmCashSales.frx":2B18
      PICN            =   "FrmCashSales.frx":2B34
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
      Left            =   12600
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
      Format          =   95944707
      CurrentDate     =   42560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   12840
      TabIndex        =   29
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label lblpayamt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Digital-7"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   13560
      TabIndex        =   27
      Top             =   7920
      Width           =   1560
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GST"
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
      Left            =   11520
      TabIndex        =   19
      Top             =   7680
      Width           =   405
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
      Left            =   10560
      TabIndex        =   15
      Top             =   7320
      Width           =   840
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
      Left            =   12480
      TabIndex        =   25
      Top             =   7320
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   10440
      Top             =   7080
      Width           =   5055
   End
   Begin MSForms.ComboBox cmbcname 
      Height          =   360
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   4425
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7805;635"
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
      Left            =   10920
      TabIndex        =   26
      Top             =   1080
      Width           =   1050
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
      Left            =   11520
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
      Left            =   3240
      TabIndex        =   21
      Top             =   7320
      Width           =   3180
   End
   Begin MSForms.ComboBox cmbcid 
      Height          =   360
      Left            =   1920
      TabIndex        =   0
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
      Left            =   3480
      TabIndex        =   18
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id"
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
      Left            =   360
      TabIndex        =   17
      Top             =   1080
      Width           =   1350
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
      Left            =   7920
      TabIndex        =   12
      Top             =   360
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmCashSales.frx":2D59
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASH SALES"
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
      TabIndex        =   9
      Top             =   240
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   15495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   7920
      Width           =   15495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   9015
      Left            =   0
      Top             =   -120
      Width           =   15495
   End
End
Attribute VB_Name = "FrmCashSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub BtnBill_Click()
Dim xlapp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
Set xlapp = CreateObject("excel.application")
'Set xlbook = xlapp.Workbooks.Add
Set xlbook = xlapp.Workbooks.Open(App.Path & "\KC_BILL.xls")
Set xlsheet = xlbook.Worksheets(1)

'------------------------Page setup---------------------------------------
xlsheet.PageSetup.PaperSize = xlPaperA5
xlsheet.PageSetup.LeftMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.RightMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.TopMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.BottomMargin = Application.InchesToPoints(0.5)
xlsheet.PageSetup.HeaderMargin = Application.InchesToPoints(0.1)
xlsheet.PageSetup.FooterMargin = Application.InchesToPoints(0.5)
xlsheet.PageSetup.Orientation = xlPortrait
'------------------------Page setup---------------------------------------
'==========================================================================================================================

ci = 7
'---------------7th Row-------------------------------------------------
xlsheet.Cells(ci, 1).Value = Trim(cmbcname.Text)

ci = ci + 1 '---------------8th Row-------------------------------------------------
If rs1.State = 1 Then rs1.Close
rs1.Open "select address1 from tbl_custmaster where cid=" & Val(Trim(cmbcid.Text)), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    If rs1.Fields("address1") <> "" Then
        xlsheet.Cells(ci, 1).Value = IIf(IsNull(rs1.Fields("address1")), "Tamilnadu", rs1.Fields("address1"))
    End If
End If
rs1.Close
xlsheet.Cells(ci, 9).Value = Format(dtp_bdate.Value, "DD/MM/YYYY")
xlsheet.Range("I" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

ci = ci + 1 '---------------9th Row-------------------------------------------------
If rs1.State = 1 Then rs1.Close
rs1.Open "select state,pincode from tbl_custmaster where cid=" & Val(Trim(cmbcid.Text)), db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    If rs1.Fields("state") <> "" Then
        xlsheet.Cells(ci, 1).Value = IIf(IsNull(rs1.Fields("state")), "", rs1.Fields("state"))
    End If
    If rs1.Fields("pincode") <> "" Then
        xlsheet.Cells(ci, 1).Value = xlsheet.Cells(ci, 1).Value & ", " & IIf(IsNull(rs1.Fields("pincode")), "", rs1.Fields("pincode"))
    End If
End If
rs1.Close
xlsheet.Cells(ci, 9).Value = Trim(txtbillno.Text)
xlsheet.Range("I" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

ci = ci + 3 '---------------12th Row-------------------------------------------------
Set i = Nothing
Set j = Nothing
For i = 1 To MSGrid.Rows - 1
    '----------------Excel Bill S.No-----------------------------------------
    xlsheet.Cells(ci, 1).Value = i
    xlsheet.Cells(ci, 1).WrapText = False
    xlsheet.Range("A" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    '----------------Excel Bill S.No-----------------------------------------
    For j = 0 To MSGrid.Cols - 1
        If j = 0 Then   '---------Grid Item Name Coloumn--------------------------
            xlsheet.Cells(ci, j + 3).Value = Trim(MSGrid.TextMatrix(i, j))
            xlsheet.Cells(ci, j + 3).WrapText = False
            xlsheet.Range("C" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        End If
        If j = 1 Then   '---------Grid Qty Coloumn--------------------------
            xlsheet.Cells(ci, j + 5).Value = Trim(MSGrid.TextMatrix(i, j))
            xlsheet.Cells(ci, j + 5).WrapText = False
            xlsheet.Range("F" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        End If
        If j = 2 Then   '---------Grid I.Amt Coloumn--------------------------
            xlsheet.Cells(ci, j + 6).Value = Trim(MSGrid.TextMatrix(i, j))
            xlsheet.Cells(ci, j + 6).WrapText = False
            xlsheet.Range("H" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            xlsheet.Range("H" & ci).NumberFormat = "0.00"
        End If
        If j = 3 Then   '---------Grid Amount Coloumn--------------------------
            xlsheet.Cells(ci, j + 6).Value = Trim(MSGrid.TextMatrix(i, j))
            xlsheet.Cells(ci, j + 6).WrapText = False
            xlsheet.Range("I" & ci).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            xlsheet.Range("I" & ci).NumberFormat = "0.00"
        End If
    Next j
    ci = ci + 1
Next i

xlsheet.Cells(30, 9).Value = txttotamt.Text
xlsheet.Range("I30").NumberFormat = "0.00"

xlsheet.Cells(31, 8).Value = txtgst.Text
xlsheet.Cells(31, 9).Value = txtstax.Text
xlsheet.Range("I31").NumberFormat = "0.00"

xlsheet.Cells(32, 9).Value = lblpayamt.Caption
xlsheet.Range("I32").NumberFormat = "0.00"

xlsheet.Cells(33, 9).Value = lblpayamt.Caption
xlsheet.Range("I33").NumberFormat = "0.00"

xlsheet.Cells(34, 9).Value = "0"
xlsheet.Range("I34").NumberFormat = "0.00"

'xlsheet.HPageBreaks.Add xlsheet.Range("A" & i + 21)

xlapp.Application.Visible = True

Sleep 100

xlsheet.PrintPreview    '---------------Print Preview
End Sub

Private Sub BtnDelete_Click()
a = MsgBox("Are You Sure to Delete This Bill", vbYesNo)
If a = vbYes Then
    db.Execute "update tbl_cashsales set billcancel='Y' where billno='" & Trim(txtbillno.Text) & "'"
    MsgBox "Bill Cancelled Successfully", vbInformation, "Kovai Cards"
End If
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmCashSales.Show
End Sub

Private Sub BtnNext_Click()
If Mid(Trim(txtbillno.Text), 8, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 9, 1)) - 1
    If Val(bno) > 9 Then
        bno = "CS00000" & bno
    Else
        bno = "CS000000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 7, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 8, 2)) - 1
    If Val(bno) > 99 Then
        bno = "CS0000" & bno
    Else
        bno = "CS00000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 6, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 7, 3)) - 1
    If Val(bno) > 999 Then
        bno = "CS000" & bno
    Else
        bno = "CS0000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 5, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 6, 4)) - 1
    If Val(bno) > 9999 Then
        bno = "CS00" & bno
    Else
        bno = "CS000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 4, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 5, 5)) - 1
    If Val(bno) > 99999 Then
        bno = "CS0" & bno
    Else
        bno = "CS00" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 3, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 4, 6)) - 1
    If Val(bno) > 99999 Then
        bno = "CS" & bno
    Else
        bno = "CS0" & bno
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
        bno = "CS00000" & bno
    Else
        bno = "CS000000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 7, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 8, 2)) - 1
    If Val(bno) > 99 Then
        bno = "CS0000" & bno
    Else
        bno = "CS00000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 6, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 7, 3)) - 1
    If Val(bno) > 999 Then
        bno = "CS000" & bno
    Else
        bno = "CS0000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 5, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 6, 4)) - 1
    If Val(bno) > 9999 Then
        bno = "CS00" & bno
    Else
        bno = "CS000" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 4, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 5, 5)) - 1
    If Val(bno) > 99999 Then
        bno = "CS0" & bno
    Else
        bno = "CS00" & bno
    End If
ElseIf Mid(Trim(txtbillno.Text), 3, 1) = "0" Then
    bno = Val(Mid(Trim(txtbillno.Text), 4, 6)) - 1
    If Val(bno) > 99999 Then
        bno = "CS" & bno
    Else
        bno = "CS0" & bno
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
lblpayamt.Caption = Format(Round(rs.Fields("pamt")), "0.00")

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
    rs.Open "select * from tbl_cashsales where billno='" & Trim(txtbillno.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MsgBox "This Bill is Already Saved", vbInformation, "Kovai Cards"
        Exit Sub
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_cashsales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 1
        rs.AddNew
        rs.Fields("billno") = Trim(txtbillno.Text)
        rs.Fields("bdate") = dtp_bdate.Value
        rs.Fields("cid") = IIf(cmbcid.Text = "", 0, Trim(cmbcid.Text))
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
        rs.Fields("pamt") = Format(Round(Val(Trim(lblpayamt.Caption))), "0.00")
        rs.Update
    Next i
    rs.Close

    MsgBox "Cash Sales Saved Successfully...", vbInformation, "Kovai Cards"

    Call BtnBill_Click
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call connect
Call Fill

Call generate_sno("CS")    'CS - Cash Sales
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
lstItemName.Clear
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

rs.Open "select cid,cname from tbl_custmaster order by cid", db, adOpenDynamic, adLockOptimistic
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
    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcid_LostFocus()
If cmbcid.Text <> "" Then
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ctype,mobno from tbl_custmaster where cid=" & Val(Trim(cmbcid.Text)) & " and cname='" & Trim(cmbcname.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        ctype = rs1.Fields("ctype")
        txtmobno.Text = rs1.Fields("mobno")
    Else
        ctype = "Customer"
        txtmobno.Text = ""
        cmbcid.Text = 0
    End If
    rs1.Close
    
    For i = 1 To MSGrid.Rows - 1
        If MSGrid.TextMatrix(i, 0) <> "" Then
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_itemmaster where iname='" & Trim(MSGrid.TextMatrix(i, 0)) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If ctype = "Customer" Then
                    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("crate"), "0.00")
                    MSGrid.TextMatrix(i, 3) = Format(Val(MSGrid.TextMatrix(i, 1)) * Val(MSGrid.TextMatrix(i, 2)), "0.00")
                ElseIf ctype = "Dealer" Then
                    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("drate"), "0.00")
                    MSGrid.TextMatrix(i, 3) = Format(Val(MSGrid.TextMatrix(i, 1)) * Val(MSGrid.TextMatrix(i, 2)), "0.00")
                Else
                    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("crate"), "0.00")
                    MSGrid.TextMatrix(i, 3) = Format(Val(MSGrid.TextMatrix(i, 1)) * Val(MSGrid.TextMatrix(i, 2)), "0.00")
                End If
            End If
        End If
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
    txtstax.Text = Format(Val(txttotamt.Text) * Val(txtgst.Text) / 100, "0.00")
    txttotamt.Text = Format(Val(txttotamt.Text), "0.00")
    
    pamt = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
    '====================================This line is for Round off the value===========================
    'MsgBox Right(pamt, InStr(pamt, ".") - 1)
    If Val(Right(pamt, InStr(pamt, ".") - 1)) > 49 Then
        amtp = Val(Left(pamt, InStr(pamt, ".") - 1)) + 1
    ElseIf Val(Right(pamt, InStr(pamt, ".") - 1)) <= 49 Then
        amtp = Val(Left(pamt, InStr(pamt, ".") - 1))
    End If
    '====================================This line is for Round off the value===========================
                
    lblpayamt.Caption = Format(Round(Val(amtp)), "0.00")
    '--------------Calculating the total quantity and amount

    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcname_Click()
If cmbcname.Text <> "" Then
    cmbcid.Text = cmbcid.List(cmbcname.ListIndex)
    lstItemName.SetFocus
End If
End Sub

Private Sub cmbcname_LostFocus()
If cmbcname.Text <> "" Then
    'On Error GoTo nxt:
    'cmbcid.Text = cmbcid.List(cmbcname.ListIndex)
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ctype,mobno from tbl_custmaster where cid=" & Val(Trim(cmbcid.Text)) & " and cname='" & Trim(cmbcname.Text) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        ctype = rs1.Fields("ctype")
        txtmobno.Text = rs1.Fields("mobno")
    Else
        ctype = "Customer"
        txtmobno.Text = ""
        cmbcid.Text = 0
    End If
    rs1.Close
    
    For i = 1 To MSGrid.Rows - 1
        If MSGrid.TextMatrix(i, 0) <> "" Then
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tbl_itemmaster where iname='" & Trim(MSGrid.TextMatrix(i, 0)) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If ctype = "Customer" Then
                    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("crate"), "0.00")
                    MSGrid.TextMatrix(i, 3) = Format(Val(MSGrid.TextMatrix(i, 1)) * Val(MSGrid.TextMatrix(i, 2)), "0.00")
                ElseIf ctype = "Dealer" Then
                    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("drate"), "0.00")
                    MSGrid.TextMatrix(i, 3) = Format(Val(MSGrid.TextMatrix(i, 1)) * Val(MSGrid.TextMatrix(i, 2)), "0.00")
                Else
                    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("crate"), "0.00")
                    MSGrid.TextMatrix(i, 3) = Format(Val(MSGrid.TextMatrix(i, 1)) * Val(MSGrid.TextMatrix(i, 2)), "0.00")
                End If
            End If
        End If
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
    txtstax.Text = Format(Val(txttotamt.Text) * Val(txtgst.Text) / 100, "0.00")
    txttotamt.Text = Format(Val(txttotamt.Text), "0.00")
    
    pamt = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
    '====================================This line is for Round off the value===========================
    'MsgBox Right(pamt, InStr(pamt, ".") - 1)
    If Val(Right(pamt, InStr(pamt, ".") - 1)) > 49 Then
        amtp = Val(Left(pamt, InStr(pamt, ".") - 1)) + 1
    ElseIf Val(Right(pamt, InStr(pamt, ".") - 1)) <= 49 Then
        amtp = Val(Left(pamt, InStr(pamt, ".") - 1))
    End If
    '====================================This line is for Round off the value===========================
                
    lblpayamt.Caption = Format(Round(Val(amtp)), "0.00")
    '--------------Calculating the total quantity and amount

    lstItemName.SetFocus
End If

'nxt:
    'lstItemName.SetFocus
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
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select ctype from tbl_custmaster where cid=" & Val(Trim(cmbcid.Text)), db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            ctype = rs1.Fields("ctype")
        End If
        rs1.Close
            
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_itemmaster where iname='" & Trim(lstItemName.List(Item)) & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            If ctype = "Customer" Then
                MSGrid.AddItem rs.Fields("iname") & vbTab & "1" & vbTab & Format(rs.Fields("crate"), "0.00") & vbTab & Format(rs.Fields("crate"), "0.00")
            ElseIf ctype = "Dealer" Then
                MSGrid.AddItem rs.Fields("iname") & vbTab & "1" & vbTab & Format(rs.Fields("drate"), "0.00") & vbTab & Format(rs.Fields("drate"), "0.00")
            Else
                MSGrid.AddItem rs.Fields("iname") & vbTab & "1" & vbTab & Format(rs.Fields("crate"), "0.00") & vbTab & Format(rs.Fields("crate"), "0.00")
            End If
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
txtstax.Text = Format(Val(txttotamt.Text) * Val(txtgst.Text) / 100, "0.00")
txttotamt.Text = Format(Val(txttotamt.Text), "0.00")

pamt = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
'====================================This line is for Round off the value===========================
'MsgBox Left(pamt, InStr(pamt, ".") - 1)
If Val(Right(pamt, InStr(pamt, ".") - 1)) > 49 Then
    amtp = Val(Left(pamt, InStr(pamt, ".") - 1)) + 1
ElseIf Val(Right(pamt, InStr(pamt, ".") - 1)) <= 49 Then
    amtp = Val(Left(pamt, InStr(pamt, ".") - 1))
End If
'====================================This line is for Round off the value===========================
            
lblpayamt.Caption = Format(Round(Val(amtp)), "0.00")
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
                txtstax.Text = Format(Val(txttotamt.Text) * Val(txtgst.Text) / 100, "0.00")
                txttotamt.Text = Format(Val(txttotamt.Text), "0.00")
                
                pamt = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
                '====================================This line is for Round off the value===========================
                'MsgBox Left(pamt, InStr(pamt, ".") - 1)
                If Val(Right(pamt, InStr(pamt, ".") - 1)) > 49 Then
                    amtp = Val(Left(pamt, InStr(pamt, ".") - 1)) + 1
                ElseIf Val(Right(pamt, InStr(pamt, ".") - 1)) <= 49 Then
                    amtp = Val(Left(pamt, InStr(pamt, ".") - 1))
                End If
                '====================================This line is for Round off the value===========================
                            
                lblpayamt.Caption = Format(Round(Val(amtp)), "0.00")
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

Private Sub txtgst_GotFocus()
txtgst.SetFocus
txtgst.SelStart = 0
txtgst.SelLength = Len(txtgst.Text)    'select the text
End Sub

Private Sub txtgst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnSave.SetFocus
End If
End Sub

Private Sub txtgst_LostFocus()
'--------------Calculating the total quantity and amount
txtstax.Text = 0
txttotamt.Text = 0
txttotqty.Text = 0
lblpayamt.Caption = 0
For i = 1 To MSGrid.Rows - 1
    txttotqty.Text = Val(txttotqty.Text) + Val(MSGrid.TextMatrix(i, 1))
    txttotamt.Text = Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 3))
Next
txtstax.Text = Format(Val(txttotamt.Text) * Val(txtgst.Text) / 100, "0.00")
txttotamt.Text = Format(Val(txttotamt.Text), "0.00")

pamt = Format(Val(txttotamt.Text) + Val(txtstax.Text), "0.00")
'====================================This line is for Round off the value===========================
'MsgBox Left(pamt, InStr(pamt, ".") - 1)
If Val(Right(pamt, InStr(pamt, ".") - 1)) > 49 Then
    amtp = Val(Left(pamt, InStr(pamt, ".") - 1)) + 1
ElseIf Val(Right(pamt, InStr(pamt, ".") - 1)) <= 49 Then
    amtp = Val(Left(pamt, InStr(pamt, ".") - 1))
End If
'====================================This line is for Round off the value===========================
            
lblpayamt.Caption = Format(Round(Val(amtp)), "0.00")

'--------------Calculating the total quantity and amount
End Sub

Private Sub txtitemsearch_Change()
lstItemName.Clear
If rs.State = 1 Then rs.Close
rs.Open "select iname from tbl_itemmaster where iid like '" & Trim(txtitemsearch.Text) & "%' order by iid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        lstItemName.AddItem rs.Fields("iname")
        rs.MoveNext
    Loop
End If
rs.Close
End Sub
