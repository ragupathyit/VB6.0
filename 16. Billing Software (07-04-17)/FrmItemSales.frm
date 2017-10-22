VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmItemSales 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbcustid 
      BeginProperty Font 
         Name            =   "Tamil-Ananthi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "FrmItemSales.frx":0000
      Left            =   1800
      List            =   "FrmItemSales.frx":0002
      TabIndex        =   31
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cmbcustname 
      BeginProperty Font 
         Name            =   "Tamil-Ananthi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "FrmItemSales.frx":0004
      Left            =   2760
      List            =   "FrmItemSales.frx":0006
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   8760
      TabIndex        =   28
      Top             =   6240
      Width           =   1335
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Balance"
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
         Width           =   1290
      End
   End
   Begin VB.TextBox txtobalance 
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
      Left            =   10200
      TabIndex        =   27
      Text            =   "0"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txtpayamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5280
      TabIndex        =   2
      Text            =   "0"
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   2160
      Width           =   6855
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Rate"
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
         Left            =   5520
         TabIndex        =   25
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Name"
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
         Left            =   2520
         TabIndex        =   24
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P. ID"
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
         TabIndex        =   23
         Top             =   0
         Width           =   510
      End
   End
   Begin VB.TextBox txtadvance 
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
      Left            =   1800
      TabIndex        =   21
      Text            =   "0"
      Top             =   1560
      Width           =   1455
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
      Left            =   10200
      TabIndex        =   13
      Text            =   "0"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtbillno 
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
      Left            =   8280
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin Project1.Button BtnClose 
      Height          =   375
      Left            =   11400
      TabIndex        =   5
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
      MICON           =   "FrmItemSales.frx":0008
      PICN            =   "FrmItemSales.frx":0024
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
      Left            =   1680
      TabIndex        =   14
      ToolTipText     =   "SAVE"
      Top             =   7920
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
      MICON           =   "FrmItemSales.frx":0736
      PICN            =   "FrmItemSales.frx":0752
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
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "BILL"
      Top             =   7920
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
      MICON           =   "FrmItemSales.frx":0E64
      PICN            =   "FrmItemSales.frx":0E80
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
      Left            =   9360
      TabIndex        =   19
      ToolTipText     =   "DELETE"
      Top             =   7920
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
      MICON           =   "FrmItemSales.frx":10A5
      PICN            =   "FrmItemSales.frx":10C1
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
      Left            =   7440
      TabIndex        =   17
      ToolTipText     =   "CLEAR"
      Top             =   7920
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
      MICON           =   "FrmItemSales.frx":17D3
      PICN            =   "FrmItemSales.frx":17EF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtp_sdate 
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   88473603
      CurrentDate     =   41799
   End
   Begin Project1.Button BtnPrevious 
      Height          =   375
      Left            =   7920
      TabIndex        =   10
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
      MICON           =   "FrmItemSales.frx":1F01
      PICN            =   "FrmItemSales.frx":1F1D
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
      Left            =   9600
      TabIndex        =   11
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
      MICON           =   "FrmItemSales.frx":262F
      PICN            =   "FrmItemSales.frx":264B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstpapername 
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
      Height          =   4755
      ItemData        =   "FrmItemSales.frx":2D5D
      Left            =   120
      List            =   "FrmItemSales.frx":2D5F
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   2160
      Width           =   4545
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   4455
      Left            =   4680
      TabIndex        =   18
      Top             =   2160
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   16761024
      GridColor       =   16761024
      Appearance      =   0
      FormatString    =   "         .|                                                                              .|                     ."
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
   Begin Project1.Button BtnBills 
      Height          =   495
      Left            =   5520
      TabIndex        =   30
      ToolTipText     =   "BILL"
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Bills    "
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
      MICON           =   "FrmItemSales.frx":2D61
      PICN            =   "FrmItemSales.frx":2D7D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Amount"
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
      Left            =   3120
      TabIndex        =   26
      Top             =   7200
      Width           =   1965
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Amt"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblcancel 
      BackColor       =   &H00FFC0C0&
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
      Left            =   8520
      TabIndex        =   16
      Top             =   7200
      Width           =   3345
   End
   Begin VB.Label Label20 
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
      Left            =   9360
      TabIndex        =   12
      Top             =   6720
      Width           =   540
   End
   Begin VB.Label Label4 
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
      Left            =   8280
      TabIndex        =   9
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cust. Name *"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
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
      Left            =   8640
      TabIndex        =   7
      Top             =   0
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "FrmItemSales.frx":2FA2
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BILL DETAILS"
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
      TabIndex        =   4
      Top             =   240
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   12135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   0
      Top             =   7800
      Width           =   12135
   End
End
Attribute VB_Name = "FrmItemSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnBill_Click()
Dim xlapp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
Set xlapp = CreateObject("excel.application")
Set xlbook = xlapp.Workbooks.Add
Set xlsheet = xlbook.Worksheets(1)

'xlsheet.Rows.WrapText = True
xlsheet.Rows.RowHeight = 20  '----------------------Excel Row height
xlsheet.Rows.Font.Name = "Arial"  '------------Each Row in this font
xlsheet.Rows.Font.Size = 12
'------------------------Page setup---------------------------------------
xlsheet.PageSetup.PaperSize = xlPaperA5
'xlsheet.PageSetup.PaperSize = xlPaperUser
xlsheet.PageSetup.LeftMargin = Application.InchesToPoints(0.3)
xlsheet.PageSetup.RightMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.TopMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.BottomMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.HeaderMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.FooterMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.Orientation = xlPortrait
'------------------------Page setup---------------------------------------
'---------------1th Row-------------------------------------------------
xlsheet.Range("A1:B1").Merge
xlsheet.Range("A1:B1").EntireRow.RowHeight = 20
xlsheet.Range("A1:B1").Font.Name = "Tamil-Ananthi"
xlsheet.Range("A1:B1").Font.Size = 14
xlsheet.Range("A1:B1").Font.Bold = True
xlsheet.Range("A1:B1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
xlsheet.Cells(1, 1).Value = "nahfpehj;"
'---------------1th Row-------------------------------------------------
'---------------2nd Row-------------------------------------------------
xlsheet.Range("A2:B2").Merge
xlsheet.Range("A2:B2").EntireRow.RowHeight = 20
xlsheet.Range("A2:B2").Font.Name = "Tamil-Ananthi"
xlsheet.Range("A2:B2").Font.Size = 14
xlsheet.Range("A2:B2").Font.Bold = True
xlsheet.Range("A2:B2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
xlsheet.Cells(2, 1).Value = "bra;jpjhs; tpw;gidahsu;"
'---------------2nd Row-------------------------------------------------
'---------------3rd Row-------------------------------------------------
xlsheet.Range("A3:B3").Merge
xlsheet.Range("A3:B3").EntireRow.RowHeight = 20
xlsheet.Range("A3:B3").Font.Name = "Tamil-Ananthi"
xlsheet.Range("A3:B3").Font.Size = 14
xlsheet.Range("A3:B3").Font.Bold = True
xlsheet.Range("A3:B3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
xlsheet.Cells(3, 1).Value = "59/32/ fw;g{uefu;/ n$hjpg[uk ;(m)/ nfhit - 47"
'---------------3rd Row-------------------------------------------------
'---------------4th Row-------------------------------------------------
xlsheet.Range("A4:B4").Merge
xlsheet.Range("A4:B4").EntireRow.RowHeight = 20
xlsheet.Range("A4:B4").Font.Size = 14
xlsheet.Range("A4:B4").Font.Bold = True
xlsheet.Range("A4:B4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
xlsheet.Cells(4, 1).Value = "9944027311, 7373027311, 9042666834"
'---------------4th Row-------------------------------------------------
'---------------5th Row--------------------------------------------------
xlsheet.Range("A5").EntireRow.RowHeight = 10
'---------------5th Row--------------------------------------------------
'---------------6th Row-------------------------------------------------
xlsheet.Range("A6:B6").Merge
xlsheet.Range("A6:B6").EntireRow.RowHeight = 20
xlsheet.Range("A6:B6").Font.Size = 14
xlsheet.Range("A6:B6").Font.Bold = True
xlsheet.Range("A6:B6").Font.Underline = True
xlsheet.Range("A6:B6").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
xlsheet.Cells(6, 1).Value = "CASH BILL"
'---------------6th Row-------------------------------------------------
'---------------7th Row-------------------------------------------------
xlsheet.Range("A7").Font.Bold = True
xlsheet.Cells(7, 1).Value = "To"
xlsheet.Cells(7, 2).Font.Bold = True
xlsheet.Cells(7, 2).Value = "Bill No."
'---------------7th Row-------------------------------------------------
'---------------8th Row-------------------------------------------------
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_customer where cid=" & Val(cmbcustid.Text) & " and customername='" & Trim(cmbcustname.Text) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    xlsheet.Range("A8:A9").Font.Name = "Tamil-Ananthi"
    xlsheet.Cells(8, 1).Value = "     " & rs.Fields("customername")
    xlsheet.Cells(8, 2).Value = Trim(txtbillno.Text)
    xlsheet.Cells(9, 1).Value = "     " & rs.Fields("address1") & "/ " & rs.Fields("state")
    xlsheet.Cells(9, 2).Font.Bold = True
    xlsheet.Cells(9, 2).Value = "Date"
    xlsheet.Cells(10, 1).Value = "     Mobile No :" & rs.Fields("mobileno")
    xlsheet.Cells(10, 2).Value = dtp_sdate.Value
End If
rs.Close
'---------------8th Row-------------------------------------------------
'---------------11th Row--------------------------------------------------
xlsheet.Range("A11").Font.Bold = True
xlsheet.Cells(11, 1).Value = "                      Advance Amount is " & Format(txtadvance.Text, "0.00")
'---------------11th Row--------------------------------------------------
'--------------------------Heading Row------------------------------------
xlsheet.Range("A12:B12").Font.Bold = True
xlsheet.Range("A12:B12").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

xlsheet.Cells(12, 1).Value = "PAPER NAME"
xlsheet.Cells(12, 2).Value = "AMOUNT"

'xlsheet.Range("A15:D15").Borders.LineStyle = xlContinuous  -----------This is for borders in all sides
xlsheet.Range("A12:B12").Borders(xlEdgeLeft).LineStyle = xlContinuous
xlsheet.Range("A12:B12").Borders(xlEdgeTop).LineStyle = xlContinuous
xlsheet.Range("A12:B12").Borders(xlEdgeRight).LineStyle = xlContinuous
xlsheet.Range("A12:B12").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlsheet.Range("A12:B12").Borders(xlInsideVertical).LineStyle = xlContinuous
'--------------------------Heading Row------------------------------------
'--------------------------Each Column Width------------------------------
'xlapp.Columns.AutoFit  ---------------------Automitically fits the column
xlsheet.Range("A1").EntireColumn.ColumnWidth = 55
xlsheet.Range("B1").EntireColumn.ColumnWidth = 12
'--------------------------Each Column Width------------------------------

Set i = Nothing
Set j = Nothing
For i = 1 To MSGrid.Rows - 2
    For j = 1 To MSGrid.Cols - 1
        xlsheet.Cells(i + 12, j).Value = MSGrid.TextMatrix(i, j)
        '--------------------Border---------------------------------------------------------
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------Border---------------------------------------------------------
        '--------------------Number Format 0.00-------------------------------------------
        xlsheet.Range("B" & i + 12).NumberFormat = "0.00"
        '--------------------Number Format 0.00-------------------------------------------
        '--------------------Only A column comes in tamil font----------------------------
        xlsheet.Range("A" & i + 12 & ":A" & i + 12).Font.Name = "Tamil-Ananthi"
        '--------------------Only A column comes in tamil font----------------------------
    Next j
Next i

'--------------------Border---------------------------------------------------------
xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeRight).LineStyle = xlContinuous
xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlInsideVertical).LineStyle = xlContinuous
'--------------------Border---------------------------------------------------------

'-----------------------------------------Extra Row for Old Balance-----------------------------
xlsheet.Range("A" & i + 13 & ":A" & i + 13).Font.Name = "Tamil-Ananthi"
xlsheet.Cells(i + 13, 1).Value = "epYitj; bjhif   "
xlsheet.Range("A" & i + 13 & ":B" & i + 13).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
xlsheet.Cells(i + 13, 2).Value = Format(Val(txtobalance.Text), "0.00")
xlsheet.Range("B" & i + 13).NumberFormat = "0.00"
'--------------------Border---------------------------------------------------------
xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeTop).LineStyle = xlContinuous
xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeRight).LineStyle = xlContinuous
xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlInsideVertical).LineStyle = xlContinuous
'--------------------Border---------------------------------------------------------
'-----------------------------------------Extra Row for Old Balance-----------------------------
'-----------------------------------------Extra Row for Total Amount-----------------------------
xlsheet.Range("A" & i + 14 & ":B" & i + 14).Font.Bold = True
xlsheet.Range("A" & i + 14 & ":B" & i + 14).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
xlsheet.Range("A" & i + 14 & ":A" & i + 14).Font.Name = "Tamil-Ananthi"
xlsheet.Cells(i + 14, 1).Value = "bkhj;jk; (+gha;)   "  '-------------------Total-------------------
xlsheet.Cells(i + 14, 2).Value = txttotamt.Text
xlsheet.Range("B" & i + 14).NumberFormat = "0.00"
'--------------------Border---------------------------------------------------------
xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeTop).LineStyle = xlContinuous
xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeRight).LineStyle = xlContinuous
xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlInsideVertical).LineStyle = xlContinuous
'--------------------Border---------------------------------------------------------
'-----------------------------------------Extra Row for Total Amount-----------------------------

xlsheet.Range("A" & i + 15 & ":B" & i + 15).Font.Bold = True
xlsheet.Range("A" & i + 15 & ":B" & i + 15).Font.Name = "Tamil-Ananthi"
xlsheet.Range("A" & i + 15 & ":B" & i + 15).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
If Month(dtp_sdate.Value) = 1 Then
    xlsheet.Cells(i + 15, 2).Value = "(ork;gu; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 2 Then
    xlsheet.Cells(i + 15, 2).Value = "($dtup khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 3 Then
    xlsheet.Cells(i + 15, 2).Value = "(gpg;utup khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 4 Then
    xlsheet.Cells(i + 15, 2).Value = "(khh;r; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 5 Then
    xlsheet.Cells(i + 15, 2).Value = "(Vg;uy; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 6 Then
    xlsheet.Cells(i + 15, 2).Value = "(nk khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 7 Then
    xlsheet.Cells(i + 15, 2).Value = "($%d; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 8 Then
    xlsheet.Cells(i + 15, 2).Value = "($%iy khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 9 Then
    xlsheet.Cells(i + 15, 2).Value = "(Mf!;l; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 10 Then
    xlsheet.Cells(i + 15, 2).Value = "(brg;lk;gu; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 11 Then
    xlsheet.Cells(i + 15, 2).Value = "(mf;nlhgu; khjj;jpw;Fz;lhdJ)"
ElseIf Month(dtp_sdate.Value) = 12 Then
    xlsheet.Cells(i + 15, 2).Value = "(etk;gu; khjj;jpw;Fz;lhdJ)"
End If

'If Not txttotal.Text = "" Then
'    If Not txttotal.Text = "0.00" Then
'        xlsheet.Range("A" & i + 16 & ":C" & i + 16).Merge
'        xlsheet.Range("A" & i + 16 & ":D" & i + 16).Font.Bold = True
'        xlsheet.Range("A" & i + 16 & ":C" & i + 16).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'        xlsheet.Cells(i + 16, 1).Value = "VAT " & Format(txtvat.Text, "0.00") & "%   (" & Format(txttotamt.Text, "0.00") & "+" & Format(Val(txttotal.Text) - Val(txttotamt.Text), "0.00") & ") : "
'        'xlsheet.Cells(i + 16, 3).Value = Val(txttotal.Text) - Val(txttotamt.Text)
'        'xlsheet.Range("C" & i + 16).NumberFormat = "0.00"
'        xlsheet.Cells(i + 16, 4).Value = txttotal.Text
'        xlsheet.Range("D" & i + 16).NumberFormat = "0.00"
'        word = ConNumToEngLish(Val(Format(txttotal.Text, "0.00")))
'    Else
'        word = ConNumToEngLish(Val(Format(txttotamt.Text, "0.00")))
'    End If
'Else
'    word = ConNumToEngLish(Val(Format(txttotamt.Text, "0.00")))
'End If
'
'xlsheet.Range("A" & i + 17 & ":D" & i + 17).Font.Bold = True
'xlsheet.Cells(i + 17, 1).Value = word & " Rupees Only"

'xlsheet.PageSetup.RightFooter = "Authorised Signature             "
'-----------------Sign Image-------------------------------------------
Dim o As Object
Set o = xlsheet.Pictures.Insert(App.Path & "\sign.jpg")
o.Top = xlsheet.Cells(i + 16, 2).Top
o.Left = xlsheet.Cells(i + 16, 2).Left
'-----------------Sign Image-------------------------------------------

xlsheet.Range("A" & i + 18 & ":B" & i + 18).Font.Bold = True
xlsheet.Range("A" & i + 18 & ":B" & i + 18).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
xlsheet.Cells(i + 18, 2).Value = "Signature      ."

xlapp.Application.Visible = True
xlsheet.PrintPreview    '---------------Print Preview
End Sub

Private Sub BtnBills_Click()
Dim xlapp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
Set xlapp = CreateObject("excel.application")
Set xlbook = xlapp.Workbooks.Add
Set xlsheet = xlbook.Worksheets(1)

Set i = Nothing
i = 1

bfrom = InputBox("Enter the Bill No From", "Bill From")
bto = InputBox("Enter the Bill No To", "Bill To")

xlsheet.Rows.RowHeight = 20  '----------------------Excel Row height
xlsheet.Rows.Font.Name = "Arial"  '------------Each Row in this font
xlsheet.Rows.Font.Size = 12
'------------------------Page setup---------------------------------------
xlsheet.PageSetup.PaperSize = xlPaperA5
'xlsheet.PageSetup.PaperSize = xlPaperUser
xlsheet.PageSetup.LeftMargin = Application.InchesToPoints(0.3)
xlsheet.PageSetup.RightMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.TopMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.BottomMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.HeaderMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.FooterMargin = Application.InchesToPoints(0.2)
xlsheet.PageSetup.Orientation = xlPortrait
'------------------------Page setup---------------------------------------
'xlapp.Columns.AutoFit  ---------------------Automitically fits the column
xlsheet.Range("A1").EntireColumn.ColumnWidth = 55
xlsheet.Range("B1").EntireColumn.ColumnWidth = 12
'--------------------------Each Column Width------------------------------

For F = bfrom To bto
    '---------------1th Row-------------------------------------------------
    xlsheet.Range("A" & i & ":B" & i).Merge
    xlsheet.Range("A" & i & ":B" & i).EntireRow.RowHeight = 20
    xlsheet.Range("A" & i & ":B" & i).Font.Name = "Tamil-Ananthi"
    xlsheet.Range("A" & i & ":B" & i).Font.Size = 14
    xlsheet.Range("A" & i & ":B" & i).Font.Bold = True
    xlsheet.Range("A" & i & ":B" & i).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    xlsheet.Cells(i, 1).Value = "nahfpehj;"
    '---------------1th Row-------------------------------------------------
    '---------------2nd Row-------------------------------------------------
    xlsheet.Range("A" & i + 1 & ":B" & i + 1).Merge
    xlsheet.Range("A" & i + 1 & ":B" & i + 1).EntireRow.RowHeight = 20
    xlsheet.Range("A" & i + 1 & ":B" & i + 1).Font.Name = "Tamil-Ananthi"
    xlsheet.Range("A" & i + 1 & ":B" & i + 1).Font.Size = 14
    xlsheet.Range("A" & i + 1 & ":B" & i + 1).Font.Bold = True
    xlsheet.Range("A" & i + 1 & ":B" & i + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    xlsheet.Cells(i + 1, 1).Value = "bra;jpjhs; tpw;gidahsu;"
    '---------------2nd Row-------------------------------------------------
    '---------------3rd Row-------------------------------------------------
    xlsheet.Range("A" & i + 2 & ":B" & i + 2).Merge
    xlsheet.Range("A" & i + 2 & ":B" & i + 2).EntireRow.RowHeight = 20
    xlsheet.Range("A" & i + 2 & ":B" & i + 2).Font.Name = "Tamil-Ananthi"
    xlsheet.Range("A" & i + 2 & ":B" & i + 2).Font.Size = 14
    xlsheet.Range("A" & i + 2 & ":B" & i + 2).Font.Bold = True
    xlsheet.Range("A" & i + 2 & ":B" & i + 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    xlsheet.Cells(i + 2, 1).Value = "59/32/ fw;g{uefu;/ n$hjpg[uk ;(m)/ nfhit - 47"
    '---------------3rd Row-------------------------------------------------
    '---------------4th Row-------------------------------------------------
    xlsheet.Range("A" & i + 3 & ":B" & i + 3).Merge
    xlsheet.Range("A" & i + 3 & ":B" & i + 3).EntireRow.RowHeight = 20
    xlsheet.Range("A" & i + 3 & ":B" & i + 3).Font.Size = 14
    xlsheet.Range("A" & i + 3 & ":B" & i + 3).Font.Bold = True
    xlsheet.Range("A" & i + 3 & ":B" & i + 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    xlsheet.Cells(i + 3, 1).Value = "9944027311, 7373027311, 9042666834"
    '---------------4th Row-------------------------------------------------
    '---------------5th Row--------------------------------------------------
    xlsheet.Range("A" & i + 4).EntireRow.RowHeight = 10
    '---------------5th Row--------------------------------------------------
    '---------------6th Row-------------------------------------------------
    xlsheet.Range("A" & i + 5 & ":B" & i + 5).Merge
    xlsheet.Range("A" & i + 5 & ":B" & i + 5).EntireRow.RowHeight = 20
    xlsheet.Range("A" & i + 5 & ":B" & i + 5).Font.Size = 14
    xlsheet.Range("A" & i + 5 & ":B" & i + 5).Font.Bold = True
    xlsheet.Range("A" & i + 5 & ":B" & i + 5).Font.Underline = True
    xlsheet.Range("A" & i + 5 & ":B" & i + 5).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    xlsheet.Cells(i + 5, 1).Value = "CASH BILL"
    '---------------6th Row-------------------------------------------------
    '---------------7th Row-------------------------------------------------
    xlsheet.Range("A" & i + 6).Font.Bold = True
    xlsheet.Cells(i + 6, 1).Value = "To"
    xlsheet.Cells(i + 6, 2).Font.Bold = True
    xlsheet.Cells(i + 6, 2).Value = "Bill No."
    '---------------7th Row-------------------------------------------------
    '---------------8th Row-------------------------------------------------
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from tbl_sales where billno=" & Val(F) & " order by pid", db, adOpenDynamic, adLockOptimistic    '############### Bill No From ################
    If Not rs1.EOF Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_customer where cid=" & Val(rs1.Fields("cid")) & " and customername='" & Trim(rs1.Fields("custname")) & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            xlsheet.Range("A" & i + 7 & ":A" & i + 8).Font.Name = "Tamil-Ananthi"
            xlsheet.Cells(i + 7, 1).Value = "     " & rs.Fields("customername")
            xlsheet.Cells(i + 7, 2).Value = Trim(rs1.Fields("billno"))
            xlsheet.Cells(i + 8, 1).Value = "     " & rs.Fields("address1") & "/ " & rs.Fields("state")
            xlsheet.Cells(i + 8, 2).Font.Bold = True
            xlsheet.Cells(i + 8, 2).Value = "Date"
            xlsheet.Cells(i + 9, 1).Value = "     Mobile No :" & rs.Fields("mobileno")
            xlsheet.Cells(i + 9, 2).Value = rs1.Fields("sdate")
            xlsheet.Range("B" & i + 9).NumberFormat = "DD-MM-YYYY"
        End If
        rs.Close
        '---------------8th Row-------------------------------------------------
        '---------------11th Row--------------------------------------------------
        xlsheet.Range("A" & i + 10).Font.Bold = True
        xlsheet.Cells(i + 10, 1).Value = "                      Advance Amount is " & Format(rs1.Fields("advamt"), "0.00")
        '---------------11th Row--------------------------------------------------
        '--------------------------Heading Row------------------------------------
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).Font.Bold = True
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        
        xlsheet.Cells(i + 11, 1).Value = "PAPER NAME"
        xlsheet.Cells(i + 11, 2).Value = "AMOUNT"
        
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 11 & ":B" & i + 11).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------------Heading Row------------------------------------
        '--------------------------Each Column Width------------------------------
        
        obalance = ""
        totamt = ""
        sdate = ""
        obalance = rs1.Fields("obalance")
        totamt = rs1.Fields("totamt")
        sdate = rs1.Fields("sdate")
        
        While Not rs1.EOF
            xlsheet.Cells(i + 12, 1).Value = rs1.Fields("papername")
            xlsheet.Cells(i + 12, 2).Value = rs1.Fields("prate")
            '--------------------Border---------------------------------------------------------
            xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeLeft).LineStyle = xlContinuous
            xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
            xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeRight).LineStyle = xlContinuous
            xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlInsideVertical).LineStyle = xlContinuous
            '--------------------Border---------------------------------------------------------
            '--------------------Number Format 0.00-------------------------------------------
            xlsheet.Range("B" & i + 12).NumberFormat = "0.00"
            '--------------------Number Format 0.00-------------------------------------------
            '--------------------Only A column comes in tamil font----------------------------
            xlsheet.Range("A" & i + 12 & ":A" & i + 12).Font.Name = "Tamil-Ananthi"
            '--------------------Only A column comes in tamil font----------------------------
            i = i + 1
            rs1.MoveNext
        Wend

        '--------------------Border---------------------------------------------------------
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 12 & ":B" & i + 12).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------Border---------------------------------------------------------

        '-----------------------------------------Extra Row for Old Balance-----------------------------
        xlsheet.Range("A" & i + 13 & ":A" & i + 13).Font.Name = "Tamil-Ananthi"
        xlsheet.Cells(i + 13, 1).Value = "epYitj; bjhif   "
        xlsheet.Range("A" & i + 13 & ":B" & i + 13).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlsheet.Cells(i + 13, 2).Value = Format(obalance, "0.00")
        xlsheet.Range("B" & i + 13).NumberFormat = "0.00"
        '--------------------Border---------------------------------------------------------
        xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 13 & ":B" & i + 13).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------Border---------------------------------------------------------
        '-----------------------------------------Extra Row for Old Balance-----------------------------
        '-----------------------------------------Extra Row for Total Amount-----------------------------
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).Font.Bold = True
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlsheet.Range("A" & i + 14 & ":A" & i + 14).Font.Name = "Tamil-Ananthi"
        xlsheet.Cells(i + 14, 1).Value = "bkhj;jk; (+gha;)   "  '-------------------Total-------------------
        xlsheet.Cells(i + 14, 2).Value = Val(totamt)
        xlsheet.Range("B" & i + 14).NumberFormat = "0.00"
        '--------------------Border---------------------------------------------------------
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsheet.Range("A" & i + 14 & ":B" & i + 14).Borders(xlInsideVertical).LineStyle = xlContinuous
        '--------------------Border---------------------------------------------------------
        '-----------------------------------------Extra Row for Total Amount-----------------------------

        xlsheet.Range("A" & i + 15 & ":B" & i + 15).Font.Bold = True
        xlsheet.Range("A" & i + 15 & ":B" & i + 15).Font.Name = "Tamil-Ananthi"
        xlsheet.Range("A" & i + 15 & ":B" & i + 15).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        If Month(sdate) = 1 Then
            xlsheet.Cells(i + 15, 2).Value = "(ork;gu; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 2 Then
            xlsheet.Cells(i + 15, 2).Value = "($dtup khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 3 Then
            xlsheet.Cells(i + 15, 2).Value = "(gpg;utup khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 4 Then
            xlsheet.Cells(i + 15, 2).Value = "(khh;r; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 5 Then
            xlsheet.Cells(i + 15, 2).Value = "(Vg;uy; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 6 Then
            xlsheet.Cells(i + 15, 2).Value = "(nk khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 7 Then
            xlsheet.Cells(i + 15, 2).Value = "($%d; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 8 Then
            xlsheet.Cells(i + 15, 2).Value = "($%iy khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 9 Then
            xlsheet.Cells(i + 15, 2).Value = "(Mf!;l; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 10 Then
            xlsheet.Cells(i + 15, 2).Value = "(brg;lk;gu; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 11 Then
            xlsheet.Cells(i + 15, 2).Value = "(mf;nlhgu; khjj;jpw;Fz;lhdJ)"
        ElseIf Month(dtp_sdate.Value) = 12 Then
            xlsheet.Cells(i + 15, 2).Value = "(etk;gu; khjj;jpw;Fz;lhdJ)"
        End If

        '-----------------Sign Image-------------------------------------------
        Dim o As Object
        Set o = xlsheet.Pictures.Insert(App.Path & "\sign.jpg")
        o.Top = xlsheet.Cells(i + 16, 2).Top
        o.Left = xlsheet.Cells(i + 16, 2).Left
        '-----------------Sign Image-------------------------------------------

        xlsheet.Range("A" & i + 18 & ":B" & i + 18).Font.Bold = True
        xlsheet.Range("A" & i + 18 & ":B" & i + 18).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlsheet.Cells(i + 18, 2).Value = "Signature      ."

        ActiveSheet.HPageBreaks.Add Before:=Rows(i + 19)
        i = i + 19
    End If
Next F
xlapp.Application.Visible = True
xlsheet.PrintPreview    '---------------Print Preview

End Sub

Private Sub BtnDelete_Click()
db.Execute "update tbl_sales set billcancel='Y' where billno=" & Val(txtbillno.Text)
MsgBox "Bill Cancelled Successfully", vbInformation, "Press Management"
Call BtnClear_Click
End Sub

Private Sub BtnClear_Click()
Unload Me
FrmItemSales.Show
End Sub

Private Sub BtnNext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) + 1, db, adOpenDynamic, adLockOptimistic
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
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Call navigation
    
    BtnSave.Enabled = False
    BtnBill.Enabled = True
    BtnDelete.Enabled = True
End If
End Sub

Function navigation()
txtbillno.Text = ""
cmbcustname.Text = ""
txtadvance.Text = ""
txttotamt.Text = ""
txtpayamt.Text = ""
MSGrid.Rows = 2
MSGrid.TextMatrix(1, 0) = ""
MSGrid.TextMatrix(1, 1) = ""
MSGrid.TextMatrix(1, 2) = ""

If rs.Fields("billcancel") = "Y" Then
    lblcancel.Caption = "CANCEL BILL"
    GoTo nxt:
Else
    lblcancel.Caption = ""
End If

txtbillno.Text = rs.Fields("billno")
cmbcustid.Text = IIf(IsNull(rs.Fields("cid")), "", rs.Fields("cid"))
cmbcustname.Text = rs.Fields("custname")
dtp_sdate.Value = rs.Fields("sdate")
txtadvance.Text = rs.Fields("advamt")
txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
txtobalance.Text = Format(rs.Fields("obalance"), "0.00")
txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")

i = 1
While Not rs.EOF
    MSGrid.TextMatrix(i, 0) = rs.Fields("pid")
    MSGrid.TextMatrix(i, 1) = rs.Fields("papername")
    MSGrid.TextMatrix(i, 2) = Format(rs.Fields("prate"), "0.00")
    MSGrid.Rows = MSGrid.Rows + 1
    i = i + 1
    rs.MoveNext
Wend
rs.Close

nxt:

End Function

Private Sub BtnSave_Click()
'-------------------Validation Starts Here-----------------------------
If cmbcustname.Text = "" Then
    MsgBox "Select the Customer Name Properly...", vbInformation, "Press Management"
    cmbcustname.SetFocus
Else
'-------------------Validation Ends Here-------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 1
        rs.AddNew
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.Fields("cid") = Val(cmbcustid.Text)
        rs.Fields("custname") = Trim(cmbcustname.Text)
        rs.Fields("sdate") = dtp_sdate.Value
        rs.Fields("advamt") = Val(txtadvance.Text)
        rs.Fields("pid") = MSGrid.TextMatrix(i, 0)
        rs.Fields("papername") = MSGrid.TextMatrix(i, 1)
        rs.Fields("prate") = Format(Val(MSGrid.TextMatrix(i, 2)), "0.00")
        rs.Fields("totamt") = Format(Round(Val(txttotamt.Text)), "0.00")
        rs.Fields("obalance") = Format(Round(Val(txtobalance.Text)), "0.00")
        rs.Fields("payamt") = Format(Round(Val(txtpayamt.Text)), "0.00")
        rs.Update
    Next i
    rs.Close
    
    If Val(txtpayamt.Text) < Val(txttotamt.Text) Then
        '====================Sales Balance====================
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_salesbalance", db, adOpenDynamic, adLockOptimistic
        rs.AddNew
            rs.Fields("billno") = Val(txtbillno.Text)
            rs.Fields("salesdate") = dtp_sdate.Value
            rs.Fields("cid") = Val(cmbcustid.Text)
            rs.Fields("custname") = UCase(cmbcustname.Text)
            rs.Fields("balamt") = Format(Round(Val(txttotamt.Text) - Val(txtpayamt.Text)), "0.00")
            rs.Fields("obalance") = Format(Round(Val(txtobalance.Text)), "0.00")
            rs.Fields("totamt") = Format(Round(Val(txttotamt.Text)), "0.00")
            rs.Fields("payamt") = Format(Round(Val(txtpayamt.Text)), "0.00")
            rs.Fields("baldesc") = Format(Round(Val(txttotamt.Text)), "0.00") & "-" & Format(Round(Val(txtpayamt.Text)), "0.00")
        rs.Update
        rs.Close
        '====================Sales Balance====================
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_salesbalance where custname='" & Trim(cmbcustname.Text) & "' order by billno", db, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            rs.MoveLast
            If rs.Fields("balamt") = 0 Then
                db.Execute "delete from tbl_salesbalance where custname='" & Trim(cmbcustname.Text) & "'"
            End If
        End If
    End If
    
    MsgBox "Bill Details Saved Successfully...", vbInformation, "Press Management"
    
    'Call BtnBill_Click
    Call BtnClear_Click
End If
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub cmbcustid_Click()
cmbcustname.Text = cmbcustname.List(cmbcustid.ListIndex)
cmbcustname.SetFocus
End Sub

Private Sub cmbcustname_Click()
cmbcustid.Text = cmbcustid.List(cmbcustname.ListIndex)
lstpapername.SetFocus
End Sub

Private Sub Form_Load()
Call connect
Call Fill

If rs.State = 1 Then rs.Close
rs.Open "select billno from tbl_sales order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = rs.Fields("billno") + 1
Else
    txtbillno.Text = 1
End If
rs.Close

BtnSave.Enabled = True
BtnDelete.Enabled = False
dtp_sdate.Value = Date
End Sub

Private Function Fill()
stmt = "select * from tbl_papermaster order by pid"
If rs.State = 1 Then rs.Close
rs.Open stmt, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        lstpapername.AddItem rs.Fields("papername")
        rs.MoveNext
    Loop
End If
rs.Close

rs.Open "select cid,customername from tbl_customer order by cid", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        cmbcustid.AddItem rs.Fields("cid")
        cmbcustname.AddItem rs.Fields("customername")
        rs.MoveNext
    Loop
End If
End Function

Private Sub lstpapername_ItemCheck(Item As Integer)
If lstpapername.Selected(Item) = True Then
    If MSGrid.TextMatrix(1, 0) = "" Then
        MSGrid.Rows = 1
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_papermaster where papername='" & Trim(lstpapername.List(Item)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MSGrid.AddItem rs.Fields("pid") & vbTab & rs.Fields("papername") & vbTab & Format(rs.Fields("prate"), "0.00")
    End If
Else
    For i = 1 To MSGrid.Rows - 1
        If MSGrid.TextMatrix(i, 1) = lstpapername.List(Item) Then
            If MSGrid.Rows = 2 Then
                MSGrid.TextMatrix(1, 0) = ""
            Else
                MSGrid.RemoveItem i
            End If
            Exit For
        End If
    Next
End If
'--------------Calculating the total amount
txttotamt.Text = 0
For i = 1 To MSGrid.Rows - 1
    txttotamt.Text = Val(txttotamt.Text) + Val(MSGrid.TextMatrix(i, 2))
Next
txttotamt.Text = Format(Val(txttotamt.Text) + Val(txtobalance.Text), "0.00")

txtpayamt.Text = Format(Val(txttotamt.Text), "0.00")

txtpayamt.SetFocus
txtpayamt.SelStart = 0
txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
End Sub

Private Sub cmbcustname_LostFocus()
If cmbcustname.Text <> "" Then
    '--------------------------Customer Advance Amount-----------------------------------
    stmt = "select advance from tbl_customer where cid=" & Trim(Val(cmbcustid.Text)) & " and customername='" & Trim(cmbcustname.Text) & "'"
    If rs1.State = 1 Then rs1.Close
    rs1.Open stmt, db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        txtadvance.Text = Format(Val(rs1.Fields("advance")), "0.00")
    End If
    rs1.Close
    
    '--------------------------Customer Old Balance Amount--------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_salesbalance where cid=" & Trim(Val(cmbcustid.Text)) & " and custname='" & cmbcustname.Text & "' order by billno", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveLast
        txtobalance.Text = Format(Val(rs.Fields("balamt")), "0.00")
    Else
        txtobalance.Text = "0.00"
    End If
    rs.Close
End If
End Sub

Private Sub txtpayamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnSave.SetFocus
End If
End Sub

Private Sub txtpayamt_LostFocus()
txtpayamt.Text = Format(Val(txtpayamt.Text), "0.00")
End Sub
