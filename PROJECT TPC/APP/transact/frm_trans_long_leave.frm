VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_long_leave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LONG LEAVE"
   ClientHeight    =   7365
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   3465
      Left            =   4680
      TabIndex        =   30
      Top             =   2820
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   6112
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectColor  =   9895934
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3180
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   720
      Width           =   3855
   End
   Begin VB.Frame fra_entry 
      Height          =   4035
      Left            =   240
      TabIndex        =   13
      Top             =   1710
      Width           =   11295
      Begin VB.TextBox txt_pph21_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2760
         Width           =   1245
      End
      Begin VB.TextBox txt_pph21_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5745
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         Top             =   2760
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DTPicker_Trans 
         Height          =   315
         Left            =   4440
         TabIndex        =   0
         Top             =   420
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   96075779
         CurrentDate     =   41331
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   4440
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3150
         Width           =   3765
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1140
         Width           =   3495
      End
      Begin prj_tpc.vbButton cmdBrowse 
         Height          =   315
         Left            =   5820
         TabIndex        =   25
         Top             =   780
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txt_employee_code 
         Height          =   315
         Left            =   6240
         TabIndex        =   27
         Top             =   780
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   4440
         TabIndex        =   33
         Top             =   1410
         Width           =   3525
         Begin VB.OptionButton optMethod 
            Caption         =   "MONEY"
            Height          =   255
            Index           =   1
            Left            =   1770
            TabIndex        =   3
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optMethod 
            Caption         =   "LEAVE"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   2
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame fraValue 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   4350
         TabIndex        =   36
         Top             =   1800
         Width           =   3765
         Begin VB.TextBox txtValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   150
            Width           =   3495
         End
      End
      Begin VB.Frame fraLeave 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   4380
         TabIndex        =   34
         Top             =   1800
         Width           =   3705
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   150
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96075779
            CurrentDate     =   41330
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   315
            Left            =   2070
            TabIndex        =   6
            Top             =   150
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   96075779
            CurrentDate     =   41330
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TO"
            Height          =   195
            Left            =   1620
            TabIndex        =   35
            Top             =   180
            Width           =   270
         End
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
         Height          =   375
         Left            =   4455
         OleObjectBlob   =   "frm_trans_long_leave.frx":001C
         TabIndex        =   10
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Frame Frame10 
         Height          =   465
         Left            =   4455
         TabIndex        =   39
         Top             =   2250
         Width           =   3495
         Begin VB.OptionButton optPPh 
            Caption         =   "MANUAL"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   7
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton optPPh 
            Caption         =   "FORMULA"
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   8
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PPh 21"
         Height          =   195
         Left            =   3600
         TabIndex        =   40
         Top             =   2370
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TRANS. DATE"
         Height          =   195
         Left            =   3105
         TabIndex        =   37
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "WITHDRAWAL METHODS"
         Height          =   195
         Left            =   2280
         TabIndex        =   32
         Top             =   1650
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   3120
         TabIndex        =   31
         Top             =   3150
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "EMP. NAME"
         Height          =   195
         Left            =   3150
         TabIndex        =   29
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "EMP. CODE"
         Height          =   195
         Left            =   3120
         TabIndex        =   28
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   5850
      Width           =   11295
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   0
         Top             =   390
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":1F78
         PICN            =   "frm_trans_long_leave.frx":1F94
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdSave 
         Height          =   705
         Left            =   1500
         TabIndex        =   17
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":3026
         PICN            =   "frm_trans_long_leave.frx":3042
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdEdit 
         Height          =   705
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":40D4
         PICN            =   "frm_trans_long_leave.frx":40F0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdDelete 
         Height          =   705
         Left            =   3540
         TabIndex        =   19
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":5182
         PICN            =   "frm_trans_long_leave.frx":519E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdCancel 
         Height          =   705
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":6230
         PICN            =   "frm_trans_long_leave.frx":624C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   10020
         TabIndex        =   21
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_trans_long_leave.frx":72DE
         PICN            =   "frm_trans_long_leave.frx":72FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4515
      Left            =   240
      TabIndex        =   12
      Top             =   1230
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7964
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "TRANS DATE"
      Columns(0).DataField=   "trans_date"
      Columns(0).NumberFormat=   "yyyy-MM-dd"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMPLOYEE CODE"
      Columns(1).DataField=   "employee_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ID EMP. CODE"
      Columns(2).DataField=   "nik"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "EMPLOYEE NAME"
      Columns(3).DataField=   "employee_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "METHODS"
      Columns(4).DataField=   "methods"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DESCRIPTION"
      Columns(5).DataField=   "description"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "FLAG TYPE"
      Columns(6).DataField=   "flag_type"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "DATE FROM"
      Columns(7).DataField=   "date_from"
      Columns(7).NumberFormat=   "dd-MM-yyyy"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "DATE TO"
      Columns(8).DataField=   "date_to"
      Columns(8).NumberFormat=   "dd-MM-yyyy"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2223"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2143"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=5133"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=5054"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2487"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2408"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=6244"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=6165"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(42)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=513"
      Splits(0)._ColumnProps(48)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST OF SPESIAL LEAVE"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=0,.bgcolor=&H80000002&"
      _StyleDefs(8)   =   ":id=4,.fgcolor=&H80000009&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(9)   =   ":id=4,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=4,.fontname=Tahoma"
      _StyleDefs(11)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(24)  =   ":id=14,.fgcolor=&H80000002&"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(70)  =   "Named:id=33:Normal"
      _StyleDefs(71)  =   ":id=33,.parent=0"
      _StyleDefs(72)  =   "Named:id=34:Heading"
      _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=34,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=35:Footing"
      _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=36:Selected"
      _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=37:Caption"
      _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(81)  =   "Named:id=38:HighlightRow"
      _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=39:EvenRow"
      _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(85)  =   "Named:id=40:OddRow"
      _StyleDefs(86)  =   ":id=40,.parent=33"
      _StyleDefs(87)  =   "Named:id=41:RecordSelector"
      _StyleDefs(88)  =   ":id=41,.parent=34"
      _StyleDefs(89)  =   "Named:id=42:FilterBar"
      _StyleDefs(90)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1410
      OleObjectBlob   =   "frm_trans_long_leave.frx":838C
      TabIndex        =   23
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   780
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LONG LEAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   15
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_long_leave.frx":A2F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11850
   End
End
Attribute VB_Name = "frm_trans_long_leave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLongLeave As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsPPh As New ADODB.Recordset

Dim vParam As String
Dim vTglAwal As Date
Dim vTglAkhir As Date
Dim vSeq As Integer

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
    check_validate_exist_new = False

    SQL = "select count(employee_code) as rec_count from t_long_leave " & _
          "where employee_code = '" & Replace$(Trim$(txt_employee_code.Text), Chr$(39), Chr$(96)) & "' " & _
                "AND trans_date = '" & Format(DTPicker_Trans.Value, "yyyy-MM-dd") & "' " & _
                "AND flag_type = " & IIf(optMethod(0).Value, 0, 1) & ""
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly

    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    txt_nik.Text = ""
    If txt_nik.Enabled = True Then txt_nik.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False

    If Not txt_employee_code = rsLongLeave.Fields("employee_code").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

    'validasi title code
    If Trim(txt_nik.Text) = "" Then
        MsgBox "ID Emp. Code is empty!", vbOKOnly + vbInformation, headerMSG
        txt_nik.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

On Error GoTo Err
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If

    i = MsgBox("Are you sure want to delete data '" _
        & TDBGrid1.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub

    CnG.BeginTrans

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    vTglAwal = Format(TDBGrid1.Columns("date_from").Value, "yyyy-MM-dd")
    vTglAkhir = Format(TDBGrid1.Columns("date_to").Value, "yyyy-MM-dd")

    vTglAwal = DateValue(vTglAwal)
    vTglAkhir = DateValue(vTglAkhir)

'    strsql = "delete from m_day where dt between '" & vTglAwal & "' and '" & vTglAkhir & "'"
'    CnG.Execute strsql

    While vTglAwal <= vTglAkhir
        DoEvents
        
        SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
                "AND absent_status = 28"
        CnG.Execute SQL
        
        vTglAwal = vTglAwal + 1
    Wend
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    CnG.Execute "delete from t_long_leave " & _
                "where employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
                    "and flag_type = " & TDBGrid1.Columns("flag_type").Value & " " & _
                    "and date(trans_date) = '" & Format(TDBGrid1.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
    CnG.CommitTrans

    Call load_data
    int_mode = 0
    Call load_mode
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
Dim vFlagMethod As Integer
Dim vFlagPPh As Integer

On Error GoTo Err
    vSetData = 1

    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If

    With rsLongLeave
        vSeq = .Fields("seq").Value
        txt_employee_code.Text = .Fields("employee_code").Value
        txt_nik.Text = .Fields("nik").Value
        txt_employee_name.Text = .Fields("employee_name").Value
        
        vFlagMethod = .Fields("flag_type").Value
        optMethod(0).Value = IIf(vFlagMethod = 0, 1, 0)
        optMethod(1).Value = IIf(vFlagMethod = 0, 0, 1)

        If optMethod(0).Value Then
            dtFrom.Value = .Fields("date_from").Value
            dtTo.Value = .Fields("date_to").Value
        Else
            txtValue.Text = FormatNumber(.Fields("value").Value)
        End If
        
        vFlagPPh = .Fields("flag_pph21").Value
        optPPh(0).Value = IIf(vFlagPPh = 0, 1, 0)
        optPPh(1).Value = IIf(vFlagPPh = 0, 0, 1)

        If optPPh(0).Value Then
            txt_pph21_value.Text = FormatNumber(.Fields("pph21_value").Value)
        Else
            TDBCombo_pph21.Text = .Fields("pph21_code").Value
            txt_pph21_name.Text = IIf(IsNull(.Fields("pph21_name").Value), "", .Fields("pph21_name").Value)
        End If

        txt_description.Text = .Fields("description").Value

    End With

    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdEdit_Click()
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdNew_Click()
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()
Dim vParameter As String
On Error GoTo Err

    '-----------------------------------------------------------------------------------------------------------------------------------------
    If optMethod(0).Value Then
        SQL = "SELECT a.employee_code FROM h_attendance a JOIN t_long_leave b on a.employee_code = b.employee_code " & _
                "WHERE date(a.att_date) BETWEEN '" & Format(dtFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtTo.Value, "yyyy-MM-dd") & "' " & _
                    "AND absent_status = 28"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        If rscari.RecordCount > 0 Then
            MsgBox "Data Found..", vbExclamation, headerMSG
            Exit Sub
        End If
        rscari.Close
    End If
    '-----------------------------------------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------------------------------------
    If optMethod(0).Value Then
        SQL = "SELECT a.employee_code FROM h_attendance a JOIN t_long_leave b on a.employee_code = b.employee_code " & _
                "WHERE date(a.att_date) BETWEEN '" & Format(dtFrom.Value, "yyyy-MM-dd") & "' AND '" & Format(dtTo.Value, "yyyy-MM-dd") & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        If rscari.RecordCount > 0 Then
            MsgBox "This Employee Has Already Absent On Between Date..", vbExclamation, headerMSG
            Exit Sub
        End If
        rscari.Close
    End If
    '-----------------------------------------------------------------------------------------------------------------------------------------

    CnG.BeginTrans
    
    SQL = "SELECT COUNT(employee_code) FROM t_long_leave WHERE employee_code = '" & txt_employee_code.Text & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vSeq = IIf(IsNull(rscari.Fields(0).Value), 0, rscari.Fields(0).Value)
    End If
    rscari.Close
    
    vSeq = vSeq + 1
    SQL = "INSERT INTO t_long_leave(trans_date,seq,employee_code,flag_type,date_from,date_to," & _
                "value,description,flag_pph21,pph21_code,pph21_value,entry_date,entry_user) " & _
            "VALUES " & _
            "('" & Format(DTPicker_Trans.Value, "yyyy-MM-dd") & "'," & vSeq & ",'" & Trim(txt_employee_code.Text) & "'," & _
            "" & IIf(optMethod(0).Value, 0, 1) & "," & _
            "'" & Format(dtFrom.Value, "yyyy-MM-dd") & "','" & Format(dtTo.Value, "yyyy-MM-dd") & "'," & _
            "'" & DropAllComma(txtValue.Text) & "','" & Trim(txt_description.Text) & "'," & _
            "" & IIf(optPPh(0).Value, 0, 1) & ",'" & TDBCombo_pph21.Text & "','" & DropAllComma(txt_pph21_value.Text) & "'," & _
            "now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If optMethod(0).Value Then
        vTglAwal = Format(dtFrom.Value, "yyyy-MM-dd")
        vTglAkhir = Format(dtTo.Value, "yyyy-MM-dd")

        vTglAwal = DateValue(vTglAwal)
        vTglAkhir = DateValue(vTglAkhir)

    '    strsql = "delete from m_day where dt between '" & vTglAwal & "' and '" & vTglAkhir & "'"
    '    CnG.Execute strsql

        While vTglAwal <= vTglAkhir
            DoEvents
                        
                SQL = "DELETE FROM h_attendance WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                        "AND date(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "' " & _
                        "AND absent_status = 28"
                CnG.Execute SQL
                
            If optMethod(0) Then
                SQL = "INSERT INTO h_attendance (employee_code, att_date, absent_status, description, entry_date) " & _
                      "VALUES (" & _
                        "'" & txt_employee_code & "','" & Format(vTglAwal, "yyyy-MM-dd") & "'," & _
                        "28,'" & txt_description.Text & "',now())"
                CnG.Execute SQL
            End If

            vTglAwal = vTglAwal + 1
        Wend
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err
    
    vTglAwal = Format(TDBGrid1.Columns("date_from").Value, "yyyy-MM-dd")
    vTglAkhir = Format(TDBGrid1.Columns("date_to").Value, "yyyy-MM-dd")

    vTglAwal = DateValue(vTglAwal)
    vTglAkhir = DateValue(vTglAkhir)

'    strsql = "delete from m_day where dt between '" & vTglAwal & "' and '" & vTglAkhir & "'"
'    CnG.Execute strsql

    While vTglAwal <= vTglAkhir
        DoEvents

        SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(vTglAwal, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
                "AND absent_status = 28"
        CnG.Execute SQL

        vTglAwal = vTglAwal + 1
    Wend
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    CnG.Execute "delete from t_long_leave " & _
                "where employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
                    "and flag_type = " & TDBGrid1.Columns("flag_type").Value & " " & _
                    "and date(trans_date) = '" & Format(TDBGrid1.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
    Call insert_new_data
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdSave_Click()
    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_new Then
            Call check_invalid: Exit Sub
        End If
        Call insert_new_data
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_edit Then
            Call check_invalid: Exit Sub
        End If
        Call edit_old_data
    End If

    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    cmdNew.Enabled = a And blnUser_Add
    cmdSave.Enabled = b
    cmdEdit.Enabled = c And blnUser_Edit
    cmdDelete.Enabled = d And blnUser_Delete
    cmdCancel.Enabled = e
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            Ctr.Value = Now
        End If
    Next
End Sub

Private Sub set_enabled_control(ByVal i As Boolean)
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            Ctr.Enabled = i
        ElseIf TypeOf Ctr Is TDBCombo Then
            Ctr.Enabled = i
        ElseIf TypeOf Ctr Is DTPicker Then
            Ctr.Value = Now
            Ctr.Enabled = i
        End If
    Next
End Sub

Private Sub set_new_data()
'    cbo_shiftable.ListIndex = 0
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        fra_entry.Visible = True
        txt_nik.Enabled = True
        TDBGrid1.Enabled = False
        Call set_new_data

        If txt_nik.Enabled = True Then
            txt_nik.SetFocus
        End If

    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry.Visible = False
        TDBGrid1.Enabled = True

    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data

        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If

        txt_nik.Enabled = False
        fra_entry.Visible = True
        TDBGrid1.Enabled = False
    End If
End Sub

Private Sub load_mode()
    If int_mode = 1 Then        ' new
        Call set_buttons_enable(False, True, False, False, True, False, False)
    ElseIf int_mode = 0 Then    ' view
        Call set_buttons_enable(True, False, True, True, False, True, True)
    ElseIf int_mode = 2 Then    ' edit/revise
        Call set_buttons_enable(False, True, False, False, True, False, False)
    End If

    Call set_data_mode
End Sub

Private Sub Form_Load()
    oClause = ""
    Call createGridKar
    Call load_data_company
    Timer1.Enabled = True

    optMethod(0).Value = True
    optPPh(0).Value = True
    txt_pph21_value.Text = FormatNumber(0)
    dtFrom.Value = Now
    dtTo.Value = Now

    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub optMethod_Click(Index As Integer)
    If Index = 0 Then
        fraLeave.Visible = True
        fraValue.Visible = False
    Else
        fraLeave.Visible = False
        fraValue.Visible = True
    End If
End Sub

Private Sub optPPh_Click(Index As Integer)
    If Index = 0 Then
        txt_pph21_value.Visible = True
        txt_pph21_value.Text = FormatNumber("0")

        TDBCombo_pph21.Visible = False
        txt_pph21_name.Visible = False
    Else
        txt_pph21_value.Visible = False

        Call load_data_pph21

        TDBCombo_pph21.Visible = True
        txt_pph21_name.Visible = True
    End If
End Sub

Private Sub load_data_pph21()
    If rsPPh.State Then rsPPh.Close
    SQL = "select * from m_pph21 order by pph21_code"
    rsPPh.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_pph21.RowSource = rsPPh
End Sub

Private Sub TDBCombo_pph21_ItemChange()
    If TDBCombo_pph21.ApproxCount > 0 Then
        TDBCombo_pph21.Text = TDBCombo_pph21.Columns("pph21_code").Value
        txt_pph21_name.Text = TDBCombo_pph21.Columns("pph21_name").Value

    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value

        Call load_data
    End If
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsLongLeave.Filter = adFilterNone
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer

    For Each Col In Cols
        If Trim(Col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If

            tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
        End If
    Next Col
    getFilter = tmp
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frm_trans_special_leave = Nothing
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields

    rsLongLeave.Filter = getFilter()
    TDBGrid1.Col = i
    TDBGrid1.EditActive = True

    TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
    If TDBGrid1.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid1.Col = i
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data()
    If rsLongLeave.State Then rsLongLeave.Close
    SQL = "select a.*, b.nik, b.company_code, b.employee_name, c.pph21_name," & _
            "CASE WHEN a.flag_type = 0 THEN 'LEAVE' ELSE 'MONEY' END methods " & _
          "from t_long_leave a join m_employee b on a.employee_code = b.employee_code " & _
            "left join m_pph21 c on a.pph21_code = c.pph21_code " & _
          "where b.company_code = '" & TDBCombo_company.Text & "'" & oClause
    rsLongLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid1.DataSource = rsLongLeave
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

    x = x + 1

    If x Mod 2 <> 1 And vSubject = TDBGrid1.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " ASC"
    End If

    vSubject = TDBGrid1.Columns(ColIndex).DataField
    Call load_data

End Sub


Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With

End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear

        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))

        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If

        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
                txt_nik.Text = rs!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else

        End If
        rs.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub LynxGrid2_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid2.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid2_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub txt_nik_Change()
    If txt_nik.Text = "" Then
        txt_employee_code.Text = ""
        txt_employee_name.Text = ""
    End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub


Private Sub txtValue_Validate(Cancel As Boolean)
    If Not Trim(txtValue) = "" Then
        txtValue = FormatNumber(DropAllComma(txtValue))
    End If
End Sub

Private Sub txt_pph21_value_Validate(Cancel As Boolean)
    If Not Trim(txt_pph21_value) = "" Then
        txt_pph21_value = FormatNumber(DropAllComma(txt_pph21_value))
    End If
End Sub
