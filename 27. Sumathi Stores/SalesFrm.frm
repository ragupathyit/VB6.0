VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SalesFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Item Purchase"
   ClientHeight    =   9210
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   18150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   18150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   8055
      Left            =   5880
      TabIndex        =   36
      Top             =   720
      Visible         =   0   'False
      Width           =   8775
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "SalesFrm.frx":0000
         Left            =   6960
         List            =   "SalesFrm.frx":000A
         TabIndex        =   39
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Print"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox rtfbill 
         Height          =   8055
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   14208
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"SalesFrm.frx":002F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
      Left            =   1680
      TabIndex        =   27
      Text            =   "0"
      Top             =   6600
      Visible         =   0   'False
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
      Left            =   1680
      TabIndex        =   17
      Text            =   "0"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txttotamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0"
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   795
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   6230
      Begin MSForms.ComboBox cmb_itemlist 
         Height          =   420
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   6230
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "10989;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Arial Narrow"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
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
      Height          =   405
      Left            =   7320
      TabIndex        =   30
      Text            =   "0"
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton CmdBill 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Bill"
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
      Left            =   4680
      TabIndex        =   26
      Top             =   7800
      Width           =   1215
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
      TabIndex        =   24
      Text            =   "0"
      Top             =   7440
      Width           =   1575
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
      Left            =   6600
      TabIndex        =   22
      Top             =   7800
      Width           =   1215
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
      Left            =   1560
      TabIndex        =   21
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
      Left            =   9360
      TabIndex        =   20
      Top             =   240
      Width           =   975
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
      TabIndex        =   19
      Top             =   0
      Width           =   5895
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
      TabIndex        =   15
      Text            =   "0"
      Top             =   6480
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
      Height          =   405
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   6480
      Width           =   855
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
      Left            =   720
      TabIndex        =   6
      Top             =   7800
      Width           =   1335
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
      Left            =   2760
      TabIndex        =   7
      Top             =   7800
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
      Left            =   3720
      TabIndex        =   5
      Top             =   8400
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
      Left            =   5760
      TabIndex        =   4
      Top             =   8400
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
      Left            =   1680
      TabIndex        =   9
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox txtbillno 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid1 
      Height          =   8775
      Left            =   12240
      TabIndex        =   13
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   15478
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   "I Code   |Item Name                                                             "
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
      Format          =   98041859
      CurrentDate     =   42430
   End
   Begin MSFlexGridLib.MSFlexGrid MSGrid 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
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
      FormatString    =   $"SalesFrm.frx":00AF
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
      Left            =   120
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   930
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
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Total Amt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7800
      TabIndex        =   35
      Top             =   8520
      Width           =   1455
   End
   Begin MSForms.ComboBox txtcustname 
      Height          =   420
      Left            =   5040
      TabIndex        =   1
      Top             =   840
      Width           =   3375
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5953;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "S"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Discount"
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
      Left            =   5640
      TabIndex        =   31
      Top             =   7080
      Width           =   1035
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
      Left            =   5640
      TabIndex        =   29
      Top             =   6600
      Width           =   1590
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
      TabIndex        =   25
      Top             =   7560
      Width           =   1380
   End
   Begin VB.Label lbldeleted 
      BackStyle       =   0  'Transparent
      Caption         =   "SALES DELETED"
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
      Left            =   720
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   3960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Total"
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
      TabIndex        =   16
      Top             =   6600
      Width           =   570
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
      Left            =   8760
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Customer Name"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "SALES ITEM DETAILS"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   4230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Bill No"
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
      TabIndex        =   8
      Top             =   840
      Width           =   1020
   End
End
Attribute VB_Name = "SalesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------Printer Code----------------------------------------------
'Private Const AnInch As Long = 1440   '1440 twips per inch
Private Const AnInch As Long = 0
Private Const QuarterInch As Long = 0
'---------------------------------Printer Code----------------------------------------------

'<================================ Printer Code ===========================================>
'Private Type DOCINFO
'    pDocName As String
'    pOutputFile As String
'    pDatatype As String
'End Type
'
'Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
'   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
'    ByVal pDefault As Long) As Long
'Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
'   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
'   pDocInfo As DOCINFO) As Long
'Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long) As Long
'Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
'   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
'   pcWritten As Long) As Long

Private Sub cmb_itemlist_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.TextMatrix(MSGrid.Row, 1) = cmb_itemlist.List(cmb_itemlist.ListIndex)
    Frame1.Visible = False
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_itemmaster where itemname='" & Trim(MSGrid.TextMatrix(MSGrid.Row, 1)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        MSGrid.TextMatrix(MSGrid.Row, 2) = Format(Val(rs.Fields("rate")), "0.00")
        MSGrid.TextMatrix(MSGrid.Row, 4) = rs.Fields("qty_type")
    End If
    
    MSGrid.Col = 3
    MSGrid.SetFocus
End If
'MsgBox KeyCode
End Sub

Private Sub CmdBill_Click()
'If Not txtcustname.Text = "" Then
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then
'        '----------Notepad print------------------
'        Open App.Path & "\bill.txt" For Output As #1
'        Print #1, Chr(27); Chr(77);         ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, Space(22) & "SUMATHI STORES"
'        Print #1, Space(19) & "Railway Station Road"
'        Print #1, Space(19) & "Mettupalayam - 641304"
'        Print #1, Space(19) & "Mobile - 99659 32314"
'        'Print #1, "------------------------------------------------------------"      '60 characters
'        'Print #1, Space(26) & "CASH BILL"
'        Print #1, "To: " & Mid(rs.Fields("custname"), 1, 30) & Space(30 - Len(Mid(rs.Fields("custname"), 1, 30)))
'        Print #1, "Bill No: " & rs.Fields("billno") & Space(6 - Len(rs.Fields("billno"))) & "                    Date: " & Format(rs.Fields("salesdate"), "DD/MM/YY") & " (" & Format(Time, "HH:MM AMPM") & ")"
'        Print #1, "------------------------------------------------------------"      '60 characters
'        Print #1, "S.No" & Space(1) & "Item Name " & Space(16) & Space(1) & "  I.Rate" & Space(1) & "Quantity" & Space(1) & "    Amount"
'        Print #1, "------------------------------------------------------------"
'        tamt = Format(rs.Fields("totamt"), "0.00")
'        word = ConNumToEngLish(Val(tamt))
'        itamt = 10 - Len(Format(tamt, "0.00"))
'
'        tqty = rs.Fields("gridtotqty")
'        itqty = 4 - Len(rs.Fields("gridtotqty"))
'
'        tsround = Round(Val(rs.Fields("totamt"))) - Val(rs.Fields("totamt"))
'        If tsround = "-0.5" Then
'            tsround = "0.5"
'        End If
'        itsround = 10 - Len(Format(tsround, "0.00"))
'
'        tsdis = Format(rs.Fields("discount"), "0.00")
'        idis = 10 - Len(Format(tsdis, "0.00"))
'
'        obalance = Format(Val(rs.Fields("obalance")), "0.00")
'        iobalance = 10 - Len(Format(obalance, "0.00"))
'
'        payamt = Format(Val(rs.Fields("payamt")), "0.00")
'        ipayamt = 10 - Len(Format(payamt, "0.00"))
'
'        balamt = Format(Val(rs.Fields("balamt")), "0.00")
'        ibalamt = 10 - Len(Format(balamt, "0.00"))
'
'        i = 1
'        While Not rs.EOF
'            ii = 3 - Len(i)
'            isno = 4 - Len(rs.Fields("sno"))
'            iname = 26 - Len(Mid(rs.Fields("itemname"), 1, 26))
'            irate = 8 - Len(Format(rs.Fields("itemrate"), "0.00"))
'            iqty = 8 - Len(rs.Fields("qty") & " " & rs.Fields("qtytype"))
'            iamt = 10 - Len(Format(rs.Fields("amount"), "0.00"))
'
'            Print #1, Space(isno) & Trim(rs.Fields("sno")) & Space(1) & UCase(Mid(rs.Fields("itemname"), 1, 25)) & Space(iname) & Space(1) & Space(irate) & Format(Val(rs.Fields("itemrate")), "0.00") & Space(1) & rs.Fields("qty") & " " & rs.Fields("qtytype") & Space(iqty) & Space(1) & Space(iamt) & Format(rs.Fields("amount"), "0.00")
'            i = i + 1
'            rs.MoveNext
'        Wend
'        Print #1, ""
'
'        If Val(obalance) <> 0 Then
'            Print #1, Space(37) & "Old Balance: " & Space(iobalance) & Format(obalance, "0.00")
'        End If
'
'        If Val(tsdis) <> 0 Then
'            Print #1, "                - Discount: " & Space(21) & Space(1) & Space(idis) & Format(tsdis, "0.00")
'        End If
'
'        Print #1, "------------------------------------------------------------"
'        Print #1, "Items: " & i - 1 & Space(ii) & Space(21) & "Total: " & Space(itqty) & tqty & Space(8) & Space(itamt) & Format(tamt, "0.00")
'
'        If Val(tsround) <> 0 Then
'            Print #1, Space(43) & "Round: " & Space(itsround) & Format(tsround, "0.00")
'            Print #1, Space(39) & "---------------------"
'            t = Format(Val(tamt) + Val(tsround), "0.00")
'            Print #1, Space(43) & "Total: " & Space(10 - Len(t)) & Format(t, "0.00")
'        End If
'
'        Print #1, Space(40) & " Payment: " & Space(ipayamt) & Format(payamt, "0.00")
'
'        If Val(balamt) <> 0 Then
'            Print #1, Space(39) & "---------------------"
'            Print #1, Space(37) & "Net Balance: " & Space(ibalamt) & Format(balamt, "0.00")
'        End If
'
'        Print #1, Space(39) & "---------------------"
'        Print #1, word & " Rupees Only"
'        Print #1, "                    Thank You! Visit Again!"
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Close #1
'    End If
'    retval = Shell("notepad.exe bill.txt", vbHide)
'    rs.Close
'
'    'Open App.Path & "\print.bat" For Output As #1 '//Creating Batch file
'    'Print #1, "TYPE bill.txt>PRN"
'    'Print #1, "EXIT"
'    'Close #1
'    'retval = Shell(App.Path & "\print.bat", vbHide)
'
'    '<==================== Printing Code ========================>
'    Dim lhPrinter As Long
'    Dim lReturn As Long
'    Dim lpcWritten As Long
'    Dim lDoc As Long
'    Dim sWrittenData As String
'    Dim MyDocInfo As DOCINFO
'    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
'    If lReturn = 0 Then
'        MsgBox "The Printer Name you typed wasn't recognized."
'        Exit Sub
'    End If
'    MyDocInfo.pDocName = "AAAAAA"
'    MyDocInfo.pOutputFile = vbNullString
'    MyDocInfo.pDatatype = vbNullString
'    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
'    Call StartPagePrinter(lhPrinter)
'
'    Dim var1 As String
'    Open App.Path & "\bill.txt" For Input As #1
'    var1 = Input(LOF(1), #1)
'    Close #1
'
'    sWrittenData = var1 '& vbFormFeed
'
'    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
'    Len(sWrittenData), lpcWritten)
'    lReturn = EndPagePrinter(lhPrinter)
'    lReturn = EndDocPrinter(lhPrinter)
'    lReturn = ClosePrinter(lhPrinter)
'    '<==================== Printing Code ========================>
'End If
'Call cmdclear_Click

rtfbill.Text = ""

If Not txtcustname.Text = "" Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        Frame2.Visible = True
        
        'rtfbill.Text = Chr(27) & Chr(77)        ' Printer Pitch 12    Form feed=Chr(12); 10 pitch=Chr(18);
        rtfbill.Text = rtfbill.Text & Space(12) & "SUMATHI STORES" & vbCrLf
        rtfbill.Text = rtfbill.Text & Space(12) & "Railway Station Road" & vbCrLf
        rtfbill.Text = rtfbill.Text & Space(12) & "Mettupalayam - 641304" & vbCrLf
        rtfbill.Text = rtfbill.Text & Space(12) & "Mobile - 99659 32314" & vbCrLf
        
        rtfbill.Text = rtfbill.Text & "To: " & Mid(rs.Fields("custname"), 1, 30) & Space(30 - Len(Mid(rs.Fields("custname"), 1, 30))) & vbCrLf
        rtfbill.Text = rtfbill.Text & "Bill No: " & rs.Fields("billno") & Space(6 - Len(rs.Fields("billno"))) & "    Date: " & Format(rs.Fields("salesdate"), "DD/MM/YY") & " (" & Format(Time, "HH:MM AMPM") & ")" & vbCrLf
        rtfbill.Text = rtfbill.Text & "--------------------------------------------" & vbCrLf      '44 characters
        rtfbill.Text = rtfbill.Text & "Item Name " & Space(5) & Space(1) & "  I.Rate" & Space(1) & "Quantity" & Space(1) & "    Amount" & vbCrLf
        rtfbill.Text = rtfbill.Text & "--------------------------------------------" & vbCrLf

        tamt = Format(rs.Fields("totamt"), "0.00")
        'word = ConNumToEngLish(Val(tamt))
        itamt = 10 - Len(Format(tamt, "0.00"))

        tqty = rs.Fields("gridtotqty")
        itqty = 4 - Len(rs.Fields("gridtotqty"))

        tsround = Round(Val(rs.Fields("totamt"))) - Val(rs.Fields("totamt"))
        If tsround = "-0.5" Then
            tsround = "0.5"
        End If
        itsround = 10 - Len(Format(tsround, "0.00"))

        tsdis = Format(rs.Fields("discount"), "0.00")
        idis = 10 - Len(Format(tsdis, "0.00"))

        obalance = Format(Val(rs.Fields("obalance")), "0.00")
        iobalance = 10 - Len(Format(obalance, "0.00"))

        payamt = Format(Val(rs.Fields("payamt")), "0.00")
        ipayamt = 10 - Len(Format(payamt, "0.00"))

        balamt = Format(Val(rs.Fields("balamt")), "0.00")
        ibalamt = 10 - Len(Format(balamt, "0.00"))

        Dim tname(100) As String
        i = 1
        j = 1
        While Not rs.EOF
            ii = 3 - Len(i)
            'isno = 4 - Len(rs.Fields("sno"))
            iname = 15 - Len(Mid(rs.Fields("itemname"), 1, 15))

            If rs1.State = 1 Then rs1.Close
            rs1.Open "select tamilname from tbl_itemmaster where itemname='" & Trim(rs.Fields("itemname")) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                tname(j) = IIf(IsNull(rs1.Fields("tamilname")), "", rs1.Fields("tamilname"))
                itname = 16 - Len(Mid(tname(j), 1, 16))
            End If
            rs1.Close

            irate = 8 - Len(Format(rs.Fields("itemrate"), "0.00"))
            iqty = 8 - Len(rs.Fields("qty") & " " & rs.Fields("qtytype"))
            iamt = 10 - Len(Format(rs.Fields("amount"), "0.00"))

            If tname(j) = "" Then
                'rtfbill.Text = rtfbill.Text & Space(isno) & Trim(rs.Fields("sno")) & Space(1) & UCase(Mid(rs.Fields("itemname"), 1, 26)) & Space(iname) & vbTab & Space(irate) & Format(Val(rs.Fields("itemrate")), "0.00") & Space(1) & rs.Fields("qty") & " " & rs.Fields("qtytype") & Space(iqty) & Space(1) & Space(iamt) & Format(rs.Fields("amount"), "0.00") & vbCrLf
                rtfbill.Text = rtfbill.Text & UCase(Mid(rs.Fields("itemname"), 1, 15)) & Space(iname) & Space(1) & Space(irate) & Format(Val(rs.Fields("itemrate")), "0.00") & Space(1) & rs.Fields("qty") & " " & rs.Fields("qtytype") & Space(iqty) & Space(1) & Space(iamt) & Format(rs.Fields("amount"), "0.00") & vbCrLf
            Else
                rtfbill.Text = rtfbill.Text & Mid(tname(j), 1, 16) & Space(itname) & vbTab & Space(irate) & Format(Val(rs.Fields("itemrate")), "0.00") & Space(1) & rs.Fields("qty") & " " & rs.Fields("qtytype") & Space(iqty) & Space(1) & Space(iamt) & Format(rs.Fields("amount"), "0.00") & vbCrLf
                j = j + 1
            End If
            
            i = i + 1
            rs.MoveNext
        Wend
        rtfbill.Text = rtfbill.Text & vbCrLf

        If Val(obalance) <> 0 Then
            rtfbill.Text = rtfbill.Text & Space(21) & "Old Balance: " & Space(iobalance) & Format(obalance, "0.00") & vbCrLf
        End If

        If Val(tsdis) <> 0 Then
            rtfbill.Text = rtfbill.Text & " - Discount: " & Space(21) & Space(1) & Space(idis) & Format(tsdis, "0.00") & vbCrLf
        End If

        rtfbill.Text = rtfbill.Text & "--------------------------------------------" & vbCrLf
        rtfbill.Text = rtfbill.Text & "Items: " & i - 1 & Space(ii) & Space(8) & "Total: " & tqty & Space(itqty) & Space(5) & Space(itamt) & Format(tamt, "0.00") & vbCrLf

        If Val(tsround) <> 0 Then
            rtfbill.Text = rtfbill.Text & Space(27) & "Round: " & Space(itsround) & Format(tsround, "0.00") & vbCrLf
            rtfbill.Text = rtfbill.Text & Space(23) & "---------------------" & vbCrLf
            t = Format(Val(tamt) + Val(tsround), "0.00")
            rtfbill.Text = rtfbill.Text & Space(24) & "Total: " & Space(10 - Len(t)) & Format(t, "0.00") & vbCrLf
        End If

        rtfbill.Text = rtfbill.Text & Space(24) & " Payment: " & Space(ipayamt) & Format(payamt, "0.00") & vbCrLf

        If Val(balamt) <> 0 Then
            rtfbill.Text = rtfbill.Text & Space(23) & "---------------------" & vbCrLf
            rtfbill.Text = rtfbill.Text & Space(21) & "Net Balance: " & Space(ibalamt) & Format(balamt, "0.00") & vbCrLf
        End If

        rtfbill.Text = rtfbill.Text & Space(23) & "---------------------" & vbCrLf
        'rtfbill.Text = rtfbill.Text & word & " Rupees Only" & vbCrLf
        rtfbill.Text = rtfbill.Text & "           Thank You! Visit Again!" & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf
        rtfbill.Text = rtfbill.Text & vbCrLf

        rtfbill.SelStart = InStr(rtfbill.Text, "SUMATHI STORES") - 1
        rtfbill.SelLength = Len("SUMATHI STORES")
        rtfbill.SelFontName = "Courier New"
        rtfbill.SelFontSize = 18
        rtfbill.SelBold = True
            
        rtfbill.SelStart = InStr(rtfbill.Text, "Railway Station Road") - 1
        rtfbill.SelLength = Len("Railway Station Road")
        rtfbill.SelFontName = "Courier New"
        rtfbill.SelFontSize = 12
        rtfbill.SelBold = True

        rtfbill.SelStart = InStr(rtfbill.Text, "Mettupalayam - 641304") - 1
        rtfbill.SelLength = Len("Mettupalayam - 641304")
        rtfbill.SelFontName = "Courier New"
        rtfbill.SelFontSize = 12
        rtfbill.SelBold = True

        rtfbill.SelStart = InStr(rtfbill.Text, "Mobile - 99659 32314") - 1
        rtfbill.SelLength = Len("Mobile - 99659 32314")
        rtfbill.SelFontName = "Courier"
        rtfbill.SelFontSize = 12
        rtfbill.SelBold = True

        For j = 1 To UBound(tname)
            If Not tname(j) = "" Then
                rtfbill.SelStart = InStr(rtfbill.Text, Mid(tname(j), 1, 23)) - 1
                rtfbill.SelLength = Len(Mid(tname(j), 1, 23))
                rtfbill.SelFontName = "Tamil-Aiswarya"
            Else
                Exit For
            End If
        Next j
    End If
End If
End Sub

Private Sub cmdclear_Click()
Unload Me
SalesFrm.Show
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
db.Execute "update tbl_sales set isdel=false where billno=" & Val(txtbillno.Text)

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    diff = Format(Val(rs.Fields("balamt")) - Val(rs.Fields("obalance")), "0.00")
    rs.Fields("isdel") = False
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance where custname='" & Trim(txtcustname.Text) & "' and isdel=true order by id", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    obalance = Format(Val(rs.Fields("balamt")), "0.00")
    totamt = Format(0 + Val(obalance), "0.00")
    balamt = Format(Val(totamt) - Val(diff), "0.00")    'Sales Balance difference Delete Panna Minus Aaganum=============================================
End If
rs.Close

'====================Sales Balance====================
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance", db, adOpenDynamic, adLockOptimistic
rs.AddNew
    rs.Fields("billno") = Val(txtbillno.Text)
    rs.Fields("salesdate") = Format(DTPicker1.Value, "MM/DD/YYYY")
    rs.Fields("cid") = Val(c)
    rs.Fields("custname") = Trim(UCase(txtcustname.Text))
    rs.Fields("balamt") = Format(Val(balamt), "0.00")
    rs.Fields("obalance") = Format(Val(obalance), "0.00")
    rs.Fields("totamt") = Format(Val(totamt), "0.00")
    rs.Fields("payamt") = Format(Val(diff), "0.00")
    rs.Fields("baldesc") = "Cancel Bill No:" & Val(txtbillno.Text)
rs.Update
rs.Close
'====================Sales Balance====================
MsgBox "Sales Item is Canceled Successfully", vbInformation, "Sumathi Stores"
Call cmdclear_Click
End Sub

Private Sub cmdcontinue_Click()
MSGrid.Row = MSGrid.Rows - 1
MSGrid.Col = 1
MSGrid.SetFocus
MSGrid.CellBackColor = RGB(117, 145, 233)
End Sub

Private Sub cmdcontinue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub CmdModify_Click()
db.Execute "update tbl_sales set isdel=false where billno=" & Val(txtbillno.Text)

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance where billno=" & Val(txtbillno.Text), db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    diff = Format(Val(rs.Fields("balamt")) - Val(rs.Fields("obalance")), "0.00")
    rs.Fields("isdel") = False
    rs.Update
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance where custname='" & Trim(txtcustname.Text) & "' and isdel=true order by id", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    obalance = Format(Val(rs.Fields("balamt")), "0.00")
    totamt = Format(0 + Val(obalance), "0.00")
    balamt = Format(Val(totamt) - Val(diff), "0.00")
End If
rs.Close

'====================Sales Balance====================
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance", db, adOpenDynamic, adLockOptimistic
rs.AddNew
    rs.Fields("billno") = Val(txtbillno.Text)
    rs.Fields("salesdate") = Format(DTPicker1.Value, "MM/DD/YYYY")
    rs.Fields("cid") = Val(c)
    rs.Fields("custname") = Trim(UCase(txtcustname.Text))
    rs.Fields("balamt") = Format(Val(balamt), "0.00")
    rs.Fields("obalance") = Format(Val(obalance), "0.00")
    rs.Fields("totamt") = Format(Val(totamt), "0.00")
    rs.Fields("payamt") = Format(Val(diff), "0.00")
    rs.Fields("baldesc") = "Modify Bill No:" & Val(txtbillno.Text)
rs.Update
rs.Close
'====================Sales Balance====================

'-----------------------------------------------------------------Add the records--------------------------------------------------------------------
If rs1.State = 1 Then rs1.Close
rs1.Open "select cid from tbl_custmaster where customername='" & Trim(txtcustname.Text) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    c = rs1.Fields("cid")
End If
rs1.Close

If txtcustname.Text = "" Then
    MsgBox "Enter the Customer Name Properly...", vbInformation, "Sumathi Stores"
    txtcustname.SetFocus
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
    For i = 1 To MSGrid.Rows - 2
        rs.AddNew
        rs.Fields("billno") = txtbillno.Text
        rs.Fields("cid") = Val(c)
        rs.Fields("custname") = Trim(UCase(txtcustname.Text))
        rs.Fields("salesdate") = Format(DTPicker1.Value, "MM/DD/YYYY")
        rs.Fields("sno") = Val(MSGrid.TextMatrix(i, 0))
        
        If rs1.State = 1 Then rs1.Close
        rs1.Open "select itemcode from tbl_itemmaster where itemname='" & Trim(MSGrid.TextMatrix(i, 1)) & "'", db, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            icode = rs1.Fields("itemcode")
        End If
        rs1.Close

        rs.Fields("itemcode") = Trim(icode)
        rs.Fields("itemname") = Trim(MSGrid.TextMatrix(i, 1))
        rs.Fields("itemrate") = Format(Val(MSGrid.TextMatrix(i, 2)), "0.00")
        rs.Fields("qty") = Format(Val(MSGrid.TextMatrix(i, 3)), "0.00")
        rs.Fields("qtytype") = Trim(MSGrid.TextMatrix(i, 4))
        rs.Fields("amount") = Format(Val(MSGrid.TextMatrix(i, 5)), "0.00")
        rs.Fields("gridtotqty") = Format(Val(txtgridtotqty.Text), "0.00")
        rs.Fields("gridtotamt") = Format(Val(txtgridtotamt.Text), "0.00")
        rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
        rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
        rs.Fields("discount") = Format(Val(txtdiscount.Text), "0.00")
        rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
        rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
        rs.Update
    Next i
    rs.Close

    '====================Sales Balance====================
    If rs.State = 1 Then rs.Close
    rs.Open "select * from tbl_salesbalance", db, adOpenDynamic, adLockOptimistic
    rs.AddNew
        rs.Fields("billno") = Val(txtbillno.Text)
        rs.Fields("salesdate") = Format(DTPicker1.Value, "MM/DD/YYYY")
        rs.Fields("cid") = Val(c)
        rs.Fields("custname") = Trim(UCase(txtcustname.Text))
        rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
        rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
        rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
        rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
        rs.Fields("baldesc") = txttotamt.Text & "-" & txtpayamt.Text
    rs.Update
    rs.Close
    '====================Sales Balance====================

    MsgBox "Successfully Saved...", vbInformation, "Sumathi Stores"
End If

db.Execute "delete from tbl_sales where billno=" & Val(txtbillno.Text) & " isdel=false"

Call cmdclear_Click
End Sub

Private Sub cmdnext_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) + 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lbldeleted.Visible = False
    txtbillno.Text = ""
    txtcustname.Text = ""
    txtgridtotamt.Text = ""
    txtgridtotqty.Text = ""
    txtobalance.Text = ""
    txttotamt.Text = ""
    txtdiscount.Text = ""
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
        txtbillno.Text = rs.Fields("billno")
        lbldeleted.Visible = True
        GoTo s:
    Else
        txtbillno.Text = rs.Fields("billno")
    End If

    txtbillno.Text = rs.Fields("billno")
    txtcustname.Text = rs.Fields("custname")
    DTPicker1.Value = rs.Fields("salesdate")
    txtgridtotqty.Text = Format(rs.Fields("gridtotqty"), "0.00")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtobalance.Text = Format(rs.Fields("obalance"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtdiscount.Text = Format(rs.Fields("discount"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    txtbalamt.Text = Format(rs.Fields("balamt"), "0.00")

    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("sno")
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 2) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 3) = rs.Fields("qty")
        MSGrid.TextMatrix(i, 4) = rs.Fields("qtytype")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("amount"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True

    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
Else
    Call cmdclear_Click
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If

s:
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True
End Sub

Private Sub cmdprevious_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales where billno=" & Val(txtbillno.Text) - 1, db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    lbldeleted.Visible = False
    txtbillno.Text = ""
    txtcustname.Text = ""
    txtgridtotamt.Text = ""
    txtgridtotqty.Text = ""
    txtobalance.Text = ""
    txttotamt.Text = ""
    txtdiscount.Text = ""
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
        txtbillno.Text = rs.Fields("billno")
        lbldeleted.Visible = True
        GoTo s:
    Else
        txtbillno.Text = rs.Fields("billno")
    End If

    txtbillno.Text = rs.Fields("billno")
    txtcustname.Text = rs.Fields("custname")
    DTPicker1.Value = rs.Fields("salesdate")
    txtgridtotqty.Text = Format(rs.Fields("gridtotqty"), "0.00")
    txtgridtotamt.Text = Format(rs.Fields("gridtotamt"), "0.00")
    txtobalance.Text = Format(rs.Fields("obalance"), "0.00")
    txttotamt.Text = Format(rs.Fields("totamt"), "0.00")
    txtdiscount.Text = Format(rs.Fields("discount"), "0.00")
    txtpayamt.Text = Format(rs.Fields("payamt"), "0.00")
    txtbalamt.Text = Format(rs.Fields("balamt"), "0.00")

    i = 1
    While Not rs.EOF
        MSGrid.TextMatrix(i, 0) = rs.Fields("sno")
        MSGrid.TextMatrix(i, 1) = rs.Fields("itemname")
        MSGrid.TextMatrix(i, 2) = Format(rs.Fields("itemrate"), "0.00")
        MSGrid.TextMatrix(i, 3) = rs.Fields("qty")
        MSGrid.TextMatrix(i, 4) = rs.Fields("qtytype")
        MSGrid.TextMatrix(i, 5) = Format(rs.Fields("amount"), "0.00")
        i = i + 1
        MSGrid.Rows = MSGrid.Rows + 1
        rs.MoveNext
    Wend

    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True

    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
Else
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 1
    MSGrid.SetFocus
End If

s:
    CmdSave.Enabled = False
    CmdModify.Enabled = True
    CmdDelete.Enabled = True
End Sub

Private Sub CmdPrint_Click()
'rtfbill.SelStart = 0
'rtfbill.SelLength = Len(rtfbill.Text)
'
'rtfbill.SelPrint Printer.hdc

PrintRTF rtfbill, 180, 180, AnInch, AnInch

Frame2.Visible = False
'Unload Me
Call cmdclear_Click
End Sub

Private Sub CmdSave_Click()
If rs1.State = 1 Then rs1.Close
rs1.Open "select cid from tbl_custmaster where customername='" & Trim(txtcustname.Text) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    c = rs1.Fields("cid")
End If
rs1.Close

If txtcustname.Text = "" Then
    MsgBox "Enter the Customer Name Properly...", vbInformation, "Sumathi Stores"
    txtcustname.SetFocus
Else
    If Not MSGrid.TextMatrix(1, 1) = "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_sales", db, adOpenDynamic, adLockOptimistic
        For i = 1 To MSGrid.Rows - 2
            rs.AddNew
            rs.Fields("billno") = txtbillno.Text
            rs.Fields("cid") = Val(c)
            rs.Fields("custname") = Trim(UCase(txtcustname.Text))
            rs.Fields("salesdate") = Format(DTPicker1.Value, "MM/DD/YYYY")
            rs.Fields("sno") = Val(MSGrid.TextMatrix(i, 0))
            
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select itemcode from tbl_itemmaster where itemname='" & Trim(MSGrid.TextMatrix(i, 1)) & "'", db, adOpenDynamic, adLockOptimistic
            If Not rs1.EOF Then
                icode = rs1.Fields("itemcode")
            End If
            rs1.Close
    
            rs.Fields("itemcode") = Trim(icode)
            rs.Fields("itemname") = Trim(MSGrid.TextMatrix(i, 1))
            rs.Fields("itemrate") = Format(Val(MSGrid.TextMatrix(i, 2)), "0.00")
            rs.Fields("qty") = Format(Val(MSGrid.TextMatrix(i, 3)), "0.00")
            rs.Fields("qtytype") = Trim(MSGrid.TextMatrix(i, 4))
            rs.Fields("amount") = Format(Val(MSGrid.TextMatrix(i, 5)), "0.00")
            rs.Fields("gridtotqty") = Format(Val(txtgridtotqty.Text), "0.00")
            rs.Fields("gridtotamt") = Format(Val(txtgridtotamt.Text), "0.00")
            rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
            rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
            rs.Fields("discount") = Format(Val(txtdiscount.Text), "0.00")
            rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
            rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
            rs.Update
        Next i
        rs.Close
    
        '====================Sales Balance====================
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tbl_salesbalance", db, adOpenDynamic, adLockOptimistic
        rs.AddNew
            rs.Fields("billno") = Val(txtbillno.Text)
            rs.Fields("salesdate") = Format(DTPicker1.Value, "MM/DD/YYYY")
            rs.Fields("cid") = Val(c)
            rs.Fields("custname") = Trim(UCase(txtcustname.Text))
            rs.Fields("balamt") = Format(Val(txtbalamt.Text), "0.00")
            rs.Fields("obalance") = Format(Val(txtobalance.Text), "0.00")
            rs.Fields("totamt") = Format(Val(txttotamt.Text), "0.00")
            rs.Fields("payamt") = Format(Val(txtpayamt.Text), "0.00")
            rs.Fields("baldesc") = txttotamt.Text & "-" & txtpayamt.Text
        rs.Update
        rs.Close
        '====================Sales Balance====================
    
        MsgBox "Successfully Saved...", vbInformation, "Sumathi Stores"
    End If
End If
Call CmdBill_Click
'Call cmdclear_Click
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then   'esc Key for Clear
    Call cmdclear_Click
End If
End Sub

Private Sub Combo1_Click()
'rtfbill.SelStart = InStr(rtfbill.Text, "SUMATHI")
'rtfbill.SelLength = Len("SUMATHI")
'rtfbill.SelText = "SUMATHI"
rtfbill.SelFontName = Combo1.Text
rtfbill.SelFontSize = 20
rtfbill.SelBold = True
End Sub

Private Sub txtcustname_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    MSGrid.Col = 1
    MSGrid.Row = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub txtcustname_LostFocus()
If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_salesbalance where custname='" & Trim(txtcustname.Text) & "' and isdel=true order by id", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtobalance.Text = Format(Val(rs.Fields("balamt")), "0.00")
Else
    txtobalance.Text = "0.00"
End If
rs.Close

MSGrid.Col = 1
MSGrid.Row = 1
MSGrid.SetFocus
End Sub

Private Sub txtdiscount_Click()
txtdiscount.SetFocus
txtdiscount.SelStart = 0
txtdiscount.SelLength = Len(txtdiscount.Text)    'select the text
End Sub

Private Sub txtdiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdiscount = Format(Val(txtdiscount.Text), "0.00")
    
'    txtpayamt.SetFocus
'    txtpayamt.SelStart = 0
'    txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
    CmdSave.SetFocus   'cursor navigation to the Save Button
End If
End Sub

Private Sub txtdiscount_LostFocus()
txttotamt.Text = 0
txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text) - Val(txtdiscount.Text), "0.00")
txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")

tsround = Round(Val(txtbalamt.Text)) - Val(txtbalamt.Text)
If tsround = "-0.5" Then
    txtbalamt.Text = Format(Val(txtbalamt.Text) + 0.5, "0.00")
End If
End Sub

Private Sub txtpayamt_Click()
txtpayamt.SetFocus
txtpayamt.SelStart = 0
txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
End Sub

Private Sub txtpayamt_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 118 Then 'F7 for Cursor move to Discount textbox
    txtpayamt.Text = ""
    txtpayamt.Text = "0.00"

    txtdiscount.SetFocus
    txtdiscount.SelStart = 0
    txtdiscount.SelLength = Len(txtdiscount.Text)    'select the text
End If
End Sub

Private Sub txtpayamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpayamt.Text = Format(Val(txtpayamt.Text), "0.00")
    
    If CmdSave.Enabled = True Then
        CmdSave.SetFocus
    Else
        CmdModify.SetFocus
    End If
End If
End Sub

Private Sub txtpayamt_LostFocus()
txttotamt.Text = 0
txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text) - Val(txtdiscount.Text), "0.00")
txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")

tsround = Round(Val(txtbalamt.Text)) - Val(txtbalamt.Text)
If tsround = "-0.5" Then
    txtbalamt.Text = Format(Val(txtbalamt.Text) + 0.5, "0.00")
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSGrid.Col = 1
    MSGrid.Row = 1
    MSGrid.SetFocus
End If
End Sub

Private Sub Form_Load()
'----------------------------Printer Code-----------------------------------------
Dim PrintableWidth As Long
Dim PrintableHeight As Long
Dim X As Single

'' Initialize Form and Command button
'Me.Caption = "Rich Text Box WYSIWYG Printing Example"
'Command1.Move 10, 10, 600, 380
'Command1.Caption = "&Print"

'' Set the font of the RTF to a TrueType font for best results
'RichTextBox1.SelFontName = "Arial"
'RichTextBox1.SelFontSize = 10

'initialize the printer object
X = Printer.TwipsPerPixelX
Printer.Orientation = vbPRORPortrait  'vbPRORLandscape

' Tell the RTF to base it's display off of the printer
Call WYSIWYG_RTF(rtfbill, QuarterInch, QuarterInch, QuarterInch, QuarterInch, PrintableWidth, PrintableHeight) '1440 Twips=1 Inch

'' Set the form width to match the line width
'Me.Width = PrintableWidth + 200
'Me.Height = PrintableHeight + 800
'----------------------------Printer Code-------------------------------------------

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
Label11.BackColor = RGB(35, 29, 29)
Label12.BackColor = RGB(35, 29, 29)
MSGrid.BackColorBkg = RGB(35, 29, 29)
MSGrid1.BackColorBkg = RGB(35, 29, 29)

Call connect

DTPicker1.Value = Date

If rs.State = 1 Then rs.Close
rs.Open "select customername from tbl_custmaster order by customername", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    txtcustname.AddItem rs.Fields("customername")
    rs.MoveNext
Wend
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select itemname from tbl_itemmaster order by itemname", db, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    cmb_itemlist.AddItem rs.Fields("itemname")
    rs.MoveNext
Wend
rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select * from tbl_sales order by billno", db, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    rs.MoveLast
    txtbillno.Text = Val(rs.Fields("billno")) + 1
Else
    txtbillno.Text = 1
End If
rs.Close

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
MSGrid.TextMatrix(1, 0) = "1"

MSGrid.Row = 1
MSGrid.Col = 1
MSGrid.CellBackColor = RGB(117, 145, 233)

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

If KeyCode = 117 Then 'F6 Key for Delete the row
    If Not MSGrid.Rows = 1 Then
        txtgridtotqty.Text = Format(Val(txtgridtotqty.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 3)), "0.00")
        txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) - Val(MSGrid.TextMatrix(MSGrid.Row, 5)), "0.00")
        txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text), "0.00")
        txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")

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
    txtdiscount.SetFocus
    txtdiscount.SelStart = 0
    txtdiscount.SelLength = Len(txtdiscount.Text)    'select the text
End If

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

If MSGrid.Col = 1 Or MSGrid.Col = 2 Or MSGrid.Col = 3 Then      'Itemname, Itemrate and Quantity grid coloumn only edited
    Select Case KeyAscii
    Case 8          ' 8 keyascii is for Back Space key
        If MSGrid.Col = 2 Or MSGrid.Col = 3 Then
            'If Not MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = "" Then MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = Mid(Trim(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)), 1, (Len(MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col)) - 1))
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = ""
        End If
'    Case 32         ' 32 keyascii is for space bar key
'        If MSGrid.Col = 1 Then
'            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
'
'            Me.Frame1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top + Me.MSGrid.CellHeight
'            Me.Frame1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
'            Frame1.Visible = True
'            cmb_itemlist.SetFocus
'        End If
    Case 46         ' 46 keyascii is for dot symbol
        If MSGrid.Col = 2 Or MSGrid.Col = 3 Then 'For Itemrate and Qty Coloumn
            MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
        End If
    Case 48 To 57   ' 48-57 keyascii is for number from 0 to 9
        MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)
    Case 65 To 90   ' 65-90 keyascii is for Caps A to Z
        If MSGrid.Col = 1 Then
'            'MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)

            Me.Frame1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top '+ Me.MSGrid.CellHeight
            Me.Frame1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
            Frame1.Visible = True
            cmb_itemlist.SetFocus
            cmb_itemlist.Text = Chr(KeyAscii)
        End If
    Case 97 To 122  ' 97-122 keyascii is for small a to z
        If MSGrid.Col = 1 Then
'            'MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) = MSGrid.TextMatrix(MSGrid.Row, MSGrid.Col) & Chr(KeyAscii)

            Me.Frame1.Top = Me.MSGrid.CellTop + Me.MSGrid.Top '+ Me.MSGrid.CellHeight
            Me.Frame1.Left = Me.MSGrid.CellLeft + Me.MSGrid.Left
            Frame1.Visible = True
            cmb_itemlist.SetFocus
            cmb_itemlist.Text = Chr(KeyAscii)
        End If
    Case 13         ' 13 keyascii is for enter key
        If MSGrid.Col = 1 Then      ' Itemname
            If MSGrid.TextMatrix(MSGrid.Row, 1) <> "" Then
            
            Else
                Frame1.Visible = False
                MSGrid.CellBackColor = vbWhite
                
'                txtpayamt.SetFocus
'                txtpayamt.SelStart = 0
'                txtpayamt.SelLength = Len(txtpayamt.Text)    'select the text
                'cmdcontinue.SetFocus   'cursor navigation to the Continue Button
                CmdSave.SetFocus   'cursor navigation to the Save Button
            End If
        End If
        
        If MSGrid.Col = 2 Then      ' Itemrate
            If MSGrid.TextMatrix(MSGrid.Row, 2) <> "" Then
                MSGrid.TextMatrix(MSGrid.Row, 2) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 2)), "0.00")
                MSGrid.TextMatrix(MSGrid.Row, 3) = ""
                MSGrid.Row = MSGrid.Row
                MSGrid.Col = 3
            End If
        End If

        If MSGrid.Col = 3 Then      ' Quantity
            If MSGrid.TextMatrix(MSGrid.Row, 3) <> "" Then
                MSGrid.TextMatrix(MSGrid.Row, 5) = Format(Val(MSGrid.TextMatrix(MSGrid.Row, 2)) * Val(MSGrid.TextMatrix(MSGrid.Row, 3)), "0.00")

                txtgridtotamt.Text = 0
                txtgridtotqty.Text = 0
                For i = 1 To MSGrid.Rows - 1
                    txtgridtotamt.Text = Format(Val(txtgridtotamt.Text) + Val(MSGrid.TextMatrix(i, 5)), "0.00")   'Grid Total bill amount calculation
                    txtgridtotqty.Text = Val(txtgridtotqty.Text) + Val(MSGrid.TextMatrix(i, 3))   'Total quantity calculation
                Next i
                txttotamt.Text = Format(Val(txtgridtotamt.Text) + Val(txtobalance.Text), "0.00")
                txtbalamt.Text = Format(Val(txttotamt.Text) - Val(txtpayamt.Text), "0.00")

                If MSGrid.TextMatrix(MSGrid.Rows - 1, 1) = "" Then
                    MSGrid.RemoveItem MSGrid.Rows - 1  'Removing the extra row in the main grid
                End If

                MSGrid.Rows = MSGrid.Rows + 1   'One row will incremented i.e., added one row
                MSGrid.Row = MSGrid.Row + 1     'cursor position changed to the newlly created row
                MSGrid.Col = 1                  'cursor position changed to the first coloumn of that newly created row
                
                MSGrid.TextMatrix(MSGrid.Rows - 1, 0) = MSGrid.Rows - 1 '
                On Error Resume Next
                SendKeys "{DOWN}"   'For Windows 7 make your project as exe. Then right click -> propertirs
                                    'then select compatibility tab then select windows xp sp2. Now u run the exe file, it will
                                    'work properly
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

If MSGrid.Col = 1 Then
    Frame1.Visible = False
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
    rs1.Open "select * from tbl_itemmaster where itemcode='" & Trim(MSGrid1.TextMatrix(MSGrid1.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        MSGrid.TextMatrix(MSGrid.Rows - 1, 1) = rs1.Fields("itemname")
        MSGrid.TextMatrix(MSGrid.Rows - 1, 2) = rs1.Fields("rate")
        MSGrid.TextMatrix(MSGrid.Rows - 1, 4) = rs1.Fields("qty_type")
        MSGrid.Row = MSGrid.Rows - 1
        MSGrid.Col = 3  ' Grid entry was changed to Qty coloumn
        MSGrid.SetFocus
    End If
    rs1.Close
End If
End Sub

Private Sub MSGrid1_Click()
MSGrid1.CellBackColor = vbWhite
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from tbl_itemmaster where itemcode='" & Trim(MSGrid1.TextMatrix(MSGrid1.Row, 0)) & "'", db, adOpenDynamic, adLockOptimistic
If Not rs1.EOF Then
    MSGrid.TextMatrix(MSGrid.Rows - 1, 1) = rs1.Fields("itemname")
    MSGrid.TextMatrix(MSGrid.Rows - 1, 2) = rs1.Fields("rate")
    MSGrid.TextMatrix(MSGrid.Rows - 1, 4) = rs1.Fields("qty_type")
    MSGrid.Row = MSGrid.Rows - 1
    MSGrid.Col = 3  ' Grid entry was changed to Qty coloumn
    MSGrid.SetFocus
End If
rs1.Close
End Sub

Private Sub MSGrid1_LeaveCell()
MSGrid1.Row = MSGrid1.Row
MSGrid1.Col = MSGrid1.Col
MSGrid1.CellBackColor = vbWhite
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
