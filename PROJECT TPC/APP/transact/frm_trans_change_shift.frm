VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_change_shift 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHANGE SHIFT"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13155
   Icon            =   "frm_trans_change_shift.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   1500
      TabIndex        =   11
      Top             =   2070
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5159
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
      GridLines       =   2
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fra_entry 
      Height          =   2955
      Left            =   90
      TabIndex        =   23
      Top             =   2190
      Visible         =   0   'False
      Width           =   12705
      Begin VB.CheckBox chkLate 
         Caption         =   "NOT INCL. LATE"
         Height          =   285
         Left            =   5160
         TabIndex        =   40
         Top             =   2130
         Width           =   2505
      End
      Begin VB.CheckBox chkTrans 
         Caption         =   "NOT INCL. TRANSPORT"
         Height          =   285
         Left            =   5160
         TabIndex        =   39
         Top             =   1860
         Width           =   2505
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1500
         Width           =   5505
      End
      Begin VB.TextBox txt_shift_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   6810
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         Top             =   810
         Width           =   3825
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_shift 
         Height          =   375
         Left            =   5160
         OleObjectBlob   =   "frm_trans_change_shift.frx":058A
         TabIndex        =   26
         Top             =   810
         Width           =   1605
      End
      Begin TDBTime6Ctl.TDBTime ttout 
         Height          =   285
         Left            =   8940
         TabIndex        =   28
         Top             =   1170
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         Caption         =   "frm_trans_change_shift.frx":24E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_trans_change_shift.frx":254E
         Spin            =   "frm_trans_change_shift.frx":259E
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "00:00"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0
      End
      Begin TDBTime6Ctl.TDBTime ttin 
         Height          =   285
         Left            =   5160
         TabIndex        =   27
         Top             =   1170
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         Caption         =   "frm_trans_change_shift.frx":25C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_trans_change_shift.frx":2632
         Spin            =   "frm_trans_change_shift.frx":2682
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "00:00"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5160
         TabIndex        =   32
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   95551491
         CurrentDate     =   40823
      End
      Begin prj_tpc.vbButton cmdSave 
         Height          =   705
         Left            =   8730
         TabIndex        =   36
         Top             =   1980
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
         MICON           =   "frm_trans_change_shift.frx":26AA
         PICN            =   "frm_trans_change_shift.frx":26C6
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
         Left            =   9720
         TabIndex        =   38
         Top             =   1980
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
         MICON           =   "frm_trans_change_shift.frx":3758
         PICN            =   "frm_trans_change_shift.frx":3774
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         Caption         =   "REMARK :"
         Height          =   210
         Left            =   4320
         TabIndex        =   35
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "DATE :"
         Height          =   195
         Left            =   4560
         TabIndex        =   33
         Top             =   510
         Width           =   525
      End
      Begin VB.Label Label14 
         Caption         =   "TIME OUT :"
         Height          =   195
         Left            =   7950
         TabIndex        =   31
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "TIME IN :"
         Height          =   195
         Left            =   4410
         TabIndex        =   30
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SHIFT :"
         Height          =   195
         Left            =   3870
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
   End
   Begin prj_tpc.vbButton cmdSetIn 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   8400
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Set Time In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_trans_change_shift.frx":4806
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdBrowse 
      Height          =   285
      Left            =   3060
      TabIndex        =   16
      Top             =   1770
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "..."
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
      MICON           =   "frm_trans_change_shift.frx":4822
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_employee_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3420
      TabIndex        =   13
      Top             =   1770
      Width           =   3495
   End
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      TabIndex        =   12
      Top             =   1770
      Width           =   1515
   End
   Begin VB.TextBox txt_department_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1410
      Width           =   3855
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1050
      Width           =   3855
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Log_Att 
      Height          =   2655
      Left            =   90
      TabIndex        =   1
      Top             =   5700
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "EMPLOYEE CODE"
      Columns(0).DataField=   "employee_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DATE"
      Columns(1).DataField=   "date_att"
      Columns(1).NumberFormat=   "yyyy-MM-dd"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TIME"
      Columns(2).DataField=   "time_att"
      Columns(2).NumberFormat=   "hh:mm:ss"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TYPE"
      Columns(3).DataField=   "type"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ATT DATE"
      Columns(4).DataField=   "att_date"
      Columns(4).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3969"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3889"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=4471"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=4392"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=3678"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3598"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos(0).PreviewInitHeight=   -1
      PrintInfos(0).PreviewInitScreenFill=   -1
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST OF LOG ATTENDANCE"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Named:id=33:Normal"
      _StyleDefs(55)  =   ":id=33,.parent=0"
      _StyleDefs(56)  =   "Named:id=34:Heading"
      _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   ":id=34,.wraptext=-1"
      _StyleDefs(59)  =   "Named:id=35:Footing"
      _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=36:Selected"
      _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=37:Caption"
      _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(65)  =   "Named:id=38:HighlightRow"
      _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker DTPicker_from 
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      Top             =   690
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   95551491
      CurrentDate     =   39270
   End
   Begin prj_tpc.vbButton cmdSearch 
      Height          =   465
      Left            =   7080
      TabIndex        =   3
      Top             =   1590
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&View"
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
      MICON           =   "frm_trans_change_shift.frx":483E
      PICN            =   "frm_trans_change_shift.frx":485A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1500
      OleObjectBlob   =   "frm_trans_change_shift.frx":58EC
      TabIndex        =   6
      Top             =   1050
      Width           =   1545
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_department 
      Height          =   375
      Left            =   1500
      OleObjectBlob   =   "frm_trans_change_shift.frx":7852
      TabIndex        =   10
      Top             =   1410
      Width           =   1545
   End
   Begin VB.TextBox txt_employee_code 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6780
      TabIndex        =   15
      Top             =   1770
      Visible         =   0   'False
      Width           =   255
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Att 
      Height          =   2955
      Left            =   90
      TabIndex        =   17
      Top             =   2190
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   5212
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "EMPLOYEE CODE"
      Columns(0).DataField=   "employee_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DATE"
      Columns(1).DataField=   "att_date"
      Columns(1).NumberFormat=   "dd-MM-yyyy"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMP. CODE"
      Columns(2).DataField=   "nik"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "EMP. NAME"
      Columns(3).DataField=   "employee_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "SHIFT CODE"
      Columns(4).DataField=   "shift_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "SHIFT NAME"
      Columns(5).DataField=   "shift_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "STATUS"
      Columns(6).DataField=   "status"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "TITLE CODE"
      Columns(7).DataField=   "title_code"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "JOB TITLE"
      Columns(8).DataField=   "title_name"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "TIME IN"
      Columns(9).DataField=   "time_in"
      Columns(9).NumberFormat=   "hh:mm"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "TIME OUT"
      Columns(10).DataField=   "time_out"
      Columns(10).NumberFormat=   "hh:mm"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "ENTRY DATE"
      Columns(11).DataField=   "entry_date"
      Columns(11).NumberFormat=   "yyyy-MM-dd"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "DEPT. CODE"
      Columns(12).DataField=   "department_code"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "DEPT. NAME"
      Columns(13).DataField=   "department_name"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "DIV. CODE"
      Columns(14).DataField=   "division_code"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "DIV. NAME"
      Columns(15).DataField=   "division_name"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "DESCRIPTION"
      Columns(16).DataField=   "descriptioon"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "STATUS NAME"
      Columns(17).DataField=   "absent_name"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "ATT DATE"
      Columns(18).DataField=   "tgl_absen"
      Columns(18).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "GROUP CODE"
      Columns(19).DataField=   "group_code"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "FLAG TRANSPORT"
      Columns(20).DataField=   "flag_transport"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "FLAG INC LATE"
      Columns(21).DataField=   "flag_inc_late"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "SHIFT"
      Columns(22).DataField=   "shift"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "FLAG DAY OVER"
      Columns(23).DataField=   "flag_day_over"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   24
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=24"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1931"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1852"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=3731"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3651"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=1958"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1879"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(6).Width=1852"
      Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=185"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=106"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(42)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2963"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2884"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(8)._MinWidth=10"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=1984"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=1905"
      Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=513"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(55)=   "Column(9)._MinWidth=54215968"
      Splits(0)._ColumnProps(56)=   "Column(10).Width=2011"
      Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=1931"
      Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(61)=   "Column(10)._MinWidth=54215968"
      Splits(0)._ColumnProps(62)=   "Column(11).Width=2223"
      Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=2143"
      Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(67)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(70)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(71)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(73)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(76)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(77)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(78)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(79)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(80)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(82)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(83)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(85)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(88)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(89)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(90)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(91)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(92)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(94)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(95)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(96)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(97)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(98)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(100)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(101)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(102)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(103)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(104)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(106)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(107)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(108)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(109)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(110)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(113)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(114)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(115)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(116)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(117)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(118)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(119)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(120)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(121)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(122)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(124)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(125)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(126)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(127)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(128)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(129)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(130)=   "Column(22)._ColStyle=516"
      Splits(0)._ColumnProps(131)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(132)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(133)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(134)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(136)=   "Column(23)._ColStyle=516"
      Splits(0)._ColumnProps(137)=   "Column(23).Order=24"
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
      Caption         =   "LIST OF ATTENDANCE"
      AllowArrows     =   0   'False
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=90,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=94,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=98,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=102,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=99,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=100,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=101,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=110,.parent=13,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=107,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=108,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=109,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=28,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=46,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=43,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=44,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=45,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=58,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=55,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=56,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=57,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=70,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=67,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=68,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=69,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=78,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=75,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=76,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=77,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=82,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=106,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=118,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=122,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=126,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=130,.parent=13"
      _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=17"
      _StyleDefs(130) =   "Named:id=33:Normal"
      _StyleDefs(131) =   ":id=33,.parent=0"
      _StyleDefs(132) =   "Named:id=34:Heading"
      _StyleDefs(133) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(134) =   ":id=34,.wraptext=-1"
      _StyleDefs(135) =   "Named:id=35:Footing"
      _StyleDefs(136) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(137) =   "Named:id=36:Selected"
      _StyleDefs(138) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(139) =   "Named:id=37:Caption"
      _StyleDefs(140) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(141) =   "Named:id=38:HighlightRow"
      _StyleDefs(142) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(143) =   "Named:id=39:EvenRow"
      _StyleDefs(144) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(145) =   "Named:id=40:OddRow"
      _StyleDefs(146) =   ":id=40,.parent=33"
      _StyleDefs(147) =   "Named:id=41:RecordSelector"
      _StyleDefs(148) =   ":id=41,.parent=34"
      _StyleDefs(149) =   "Named:id=42:FilterBar"
      _StyleDefs(150) =   ":id=42,.parent=33"
   End
   Begin prj_tpc.vbButton cmdSetOut 
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   8400
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Set Time Out"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_trans_change_shift.frx":97BB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdChangeShift 
      Height          =   285
      Left            =   1470
      TabIndex        =   20
      Top             =   5190
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Change Shift"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_trans_change_shift.frx":97D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker_to 
      Height          =   315
      Left            =   5100
      TabIndex        =   21
      Top             =   690
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   95551491
      CurrentDate     =   39270
   End
   Begin prj_tpc.vbButton cmdAddShift 
      Height          =   285
      Left            =   210
      TabIndex        =   24
      Top             =   5190
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Add Shift"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_trans_change_shift.frx":97F3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   11610
      TabIndex        =   37
      Top             =   8850
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
      MICON           =   "frm_trans_change_shift.frx":980F
      PICN            =   "frm_trans_change_shift.frx":982B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker_Periode 
      Height          =   315
      Left            =   1500
      TabIndex        =   41
      Top             =   690
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM-yyyy"
      Format          =   95551491
      CurrentDate     =   39270
   End
   Begin prj_tpc.vbButton cmdDeleteShift 
      Height          =   285
      Left            =   2730
      TabIndex        =   42
      Top             =   5190
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Delete Shift"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_trans_change_shift.frx":A8BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4650
      TabIndex        =   22
      Top             =   750
      Width           =   405
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EMPLOYEE*"
      Height          =   195
      Left            =   450
      TabIndex        =   14
      Top             =   1770
      Width           =   930
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "COMPANY*"
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Label60 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERIODE*"
      Height          =   195
      Left            =   585
      TabIndex        =   4
      Top             =   750
      Width           =   780
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE SHIFT"
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
      Left            =   300
      TabIndex        =   0
      Top             =   150
      Width           =   4365
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   -750
      Picture         =   "frm_trans_change_shift.frx":A8D9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_trans_change_shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmployee As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDepartment As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset
Dim rsLogAtt As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsShift As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim i_lang As Integer
Dim v_company As String
Dim dst As String

Public v_tt_in, v_tt_out, v_absen_status As String
Dim tglAwal, tglAkhir As Date

Dim i As Integer
Dim editTrans As Boolean

Dim vParam As String

Private Sub cmdChangeShift_Click()
    fra_entry.Visible = True
    chkTrans.Value = 0
    
    Call showData
End Sub

Private Sub cmdAddShift_Click()
    fra_entry.Visible = True
    chkTrans.Value = 0
    
    Call clear_view_data
    
    DTPicker1.Enabled = True
    
    If Not (TDBGrid_Att.ApproxCount > 0 And TDBGrid_Att.Bookmark > 0) Then
        DTPicker1.Value = Format(frm_list_manual_att.TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
    Else
        DTPicker1.Value = Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
    End If
    editTrans = False
End Sub

Private Sub cmdDeleteShift_Click()
Dim i As Integer
    If Not (TDBGrid_Att.ApproxCount > 0 And TDBGrid_Att.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Are you sure want to delete data ''" _
        & Format(TDBGrid_Att.Columns("tgl_absen").Value, "yyyy-MM-dd") & "' -" _
        & TDBGrid_Att.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    SQL = "DELETE FROM h_attendance " & _
          "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
            "AND att_date = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "'"
    CnG.Execute SQL
    
    Call load_data_att
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" And _
               Not LCase(Ctr.name) = "txt_nik" And _
               Not LCase(Ctr.name) = "txt_employee_code" And _
               Not LCase(Ctr.name) = "txt_employee_name" And _
               Not LCase(Ctr.name) = "txt_department_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" And _
               Not LCase(Ctr.name) = "tdbcombo_department" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            If Not LCase(Ctr.name) = "dtpicker_from" And _
               Not LCase(Ctr.name) = "dtpicker_to" Then Ctr.Value = Now
        End If
    Next
End Sub

Private Sub showData()
Dim vFlagTransport As String
Dim vFlagLate As String
    If TDBGrid_Att.ApproxCount > 0 Then
        DTPicker1.Value = Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd hh:nn:ss")
        
        TDBCombo_shift.Text = TDBGrid_Att.Columns("shift").Value
        txt_shift_name.Text = TDBGrid_Att.Columns("shift_name").Value
        
        ttin.Value = Format(TDBGrid_Att.Columns("time_in").Value, "hh:mm")
        ttout.Value = Format(TDBGrid_Att.Columns("time_out").Value, "hh:mm")
        
        txt_description = TDBGrid_Att.Columns("description").Value
        vFlagTransport = TDBGrid_Att.Columns("flag_transport").Value
        If vFlagTransport = "" Then
            chkTrans.Value = 0
        Else
            If vFlagTransport < 0 Then
                chkTrans.Value = 1
            Else
                chkTrans.Value = TDBGrid_Att.Columns("flag_transport").Value
            End If
        End If
        
        vFlagLate = TDBGrid_Att.Columns("flag_inc_late").Value
        If vFlagLate = "" Then
            chkLate.Value = 0
        Else
            If vFlagLate < 0 Then
                chkLate.Value = 1
            Else
                chkLate.Value = TDBGrid_Att.Columns("flag_inc_late").Value
            End If
        End If

        DTPicker1.Enabled = False
        editTrans = True
    End If
End Sub

Private Sub set_new_data()
    DTPicker_Periode.Value = Now
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
End Sub

Private Sub cmdSearch_Click()
    If TDBCombo_company.Text = "" Then
        MsgBox "Company Is Empty!", vbExclamation, headerMSG
        TDBCombo_company.SetFocus
        Exit Sub
    End If

    If txt_nik.Text = "" Then
        MsgBox "Employee Not Selected!", vbExclamation, headerMSG
        txt_nik.SetFocus
        Exit Sub
    End If
    
    vModeLoad = 0
    Call load_data_att
    Call load_data_shift
End Sub

Private Sub cmdSetIn_Click()
On Error GoTo Err
    If Not (TDBGrid_Log_Att.ApproxCount > 0 And TDBGrid_Log_Att.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If

    i = MsgBox("Are you sure want to set '" _
        & Format(TDBGrid_Log_Att.Columns("att_date").Value, "yyyy-MM-dd hh:mm:ss") & "' as Time In?", vbYesNo + vbQuestion, headerMSG)
        
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    
    SQL = "UPDATE h_attendance SET time_in = '" & Format(TDBGrid_Log_Att.Columns("att_date").Value, "yyyy-MM-dd hh:mm:ss") & "' " & _
          "WHERE att_date = '" & Format(TDBGrid_Att.Columns("tgl_absen").Value, "yyyy-MM-dd hh:mm:ss") & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans
    
    Call load_data_att
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdSetOut_Click()
On Error GoTo Err
    If Not (TDBGrid_Log_Att.ApproxCount > 0 And TDBGrid_Log_Att.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If

    i = MsgBox("Are you sure want to set '" _
        & Format(TDBGrid_Log_Att.Columns("att_date").Value, "yyyy-MM-dd hh:mm:ss") & "' as Time Out?", vbYesNo + vbQuestion, headerMSG)
        
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    
    SQL = "UPDATE h_attendance SET time_out = '" & Format(TDBGrid_Log_Att.Columns("att_date").Value, "yyyy-MM-dd hh:mm:ss") & "' " & _
          "WHERE att_date = '" & Format(TDBGrid_Att.Columns("tgl_absen").Value, "yyyy-MM-dd hh:mm:ss") & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans
    
    Call load_data_att
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_from, DTPicker_to)
End Sub

Private Sub Form_Load()
    Call load_data_company
    Call createGridKar

'    Call load_data_user_access(Me)
    Call set_new_data

    oClause = ""
    
    timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_trans_change_shift = Nothing
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value

        If LOGIN_LEVEL = 100 Then
            Call load_data_department
        Else
            If DEPARTMENT_CODE <> "" Then
                SQL = "SELECT department_name FROM m_department " & _
                        "WHERE company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND department_code = '" & DEPARTMENT_CODE & "'"
                rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

                If rs.RecordCount > 0 Then
                    TDBCombo_department.Text = DEPARTMENT_CODE
                    txt_department_name.Text = rs!department_name
                    TDBCombo_department.Enabled = False
                    txt_department_name.Enabled = False
                End If
                rs.Close
            Else
                Call load_data_department
            End If
        End If
    End If
End Sub

Private Sub TDBCombo_department_Change()
    If TDBCombo_department.Text = "" Then
        txt_department_name.Text = ""
    End If
End Sub

Private Sub TDBCombo_department_ItemChange()
    If TDBCombo_department.ApproxCount > 0 Then
        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
        txt_department_name.Text = TDBCombo_department.Columns("department_name").Value
    End If
End Sub

Private Sub TDBCombo_shift_ItemChange()
    If TDBCombo_shift.ApproxCount > 0 Then
        TDBCombo_shift.Text = TDBCombo_shift.Columns("shift_code").Value
        txt_shift_name.Text = TDBCombo_shift.Columns("shift_name").Value
        
        SQL = "SELECT start_time, end_time FROM m_shift WHERE shift_code = '" & TDBCombo_shift.Text & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            ttin.Value = Format(rscari!start_time, "hh:nn")
            ttout.Value = Format(rscari!end_time, "hh:nn")
        End If
        rscari.Close
    End If
End Sub

Public Sub load_data_company()
    TDBCombo_company.Text = "": txt_company_name = ""

    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_company.RowSource = rsCompany
End Sub

Public Sub load_data_department()
    TDBCombo_department.Text = "": txt_department_name.Text = ""

    If rsDepartment.State Then rsDepartment.Close
    SQL = "select * from m_department where company_code = '" & TDBCombo_company.Text & "' order by company_code"
    rsDepartment.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_department.RowSource = rsDepartment
End Sub

Public Sub load_data_shift()
    If rsShift.State Then rsShift.Close
    SQL = "select * from m_shift order by shift_code"
    rsShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_shift.RowSource = rsShift
End Sub

Public Sub load_data_att()
    If rsAtt.State Then rsAtt.Close
    
    If vModeLoad = 0 Then
        SQL = "select a.att_date,a.employee_code,b.nik,b.employee_name,a.status,e.absent_name,b.title_code,c.title_name, " & _
                    "a.time_in,a.time_out,a.entry_date,b.division_code,d.division_name,a.description, " & _
                    "b.department_code, f.department_name, " & _
                    "CASE WHEN a.shift_code = 'S01' THEN 'M' " & _
                        "WHEN a.shift_code = 'S02' THEN 'A' " & _
                        "WHEN a.shift_code = 'S03' THEN 'N' " & _
                        "WHEN a.shift_code = 'ST01' THEN 'DAILY' " & _
                        "WHEN a.shift_code = 'OFF' THEN 'OFF' END shift_code," & _
                    "g.shift_name, a.att_date AS tgl_absen, g.group_code, a.flag_transport, a.flag_inc_late, a.shift_code shift, g.flag_day_over " & _
                  "from h_attendance a join m_employee b on a.employee_code = b.employee_code " & _
                    "left join m_title c on b.title_code = c.title_code " & _
                    "left join m_division d on b.division_code = d.division_code and b.department_code = d.department_code and b.company_code = d.company_code " & _
                    "join m_absent_status e on a.status = e.absent_code " & _
                    "left join m_department f on b.department_code = f.department_code and b.company_code = f.company_code " & _
                    "left join m_shift g on a.shift_code = g.shift_code " & _
              "WHERE a.employee_code = '" & txt_employee_code.Text & "' " & _
                "AND (date(a.att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & oClause
    Else
        SQL = "select a.att_date,a.employee_code,b.nik,b.employee_name,a.status,e.absent_name,b.title_code,c.title_name, " & _
                    "a.time_in,a.time_out,a.entry_date,b.division_code,d.division_name,a.description, " & _
                    "b.department_code, f.department_name, " & _
                    "CASE WHEN a.shift_code = 'S01' THEN 'M' " & _
                        "WHEN a.shift_code = 'S02' THEN 'A' " & _
                        "WHEN a.shift_code = 'S03' THEN 'N' " & _
                        "WHEN a.shift_code = 'ST01' THEN 'DAILY' " & _
                        "WHEN a.shift_code = 'OFF' THEN 'OFF' END shift_code," & _
                    "g.shift_name, a.att_date AS tgl_absen, g.group_code, a.flag_transport, a.flag_inc_late, a.shift_code shift, g.flag_day_over " & _
                  "from h_attendance a join m_employee b on a.employee_code = b.employee_code " & _
                    "left join m_title c on b.title_code = c.title_code " & _
                    "left join m_division d on b.division_code = d.division_code and b.department_code = d.department_code and b.company_code = d.company_code " & _
                    "join m_absent_status e on a.status = e.absent_code " & _
                    "left join m_department f on b.department_code = f.department_code and b.company_code = f.company_code " & _
                    "left join m_shift g on a.shift_code = g.shift_code " & _
              "WHERE a.employee_code = '" & txt_employee_code.Text & "' " & _
                "AND date(a.att_date) = '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "'"

    End If
    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Att.DataSource = rsAtt
End Sub

Private Sub load_data_log_att()
    If rsLogAtt.State Then rsLogAtt.Close
'    SQL = "SELECT b.*, DATE(b.att_date) AS date_att, TIME(b.att_date) AS time_att," & _
'             "CASE WHEN b.flag_io = 0 THEN 'IN' ELSE 'OUT' END TYPE " & _
'          "FROM h_attendance a JOIN h_log_attendance b ON DATE(a.att_date) = DATE(b.att_date) AND a.employee_code = b.employee_code " & _
'            "JOIN m_shift c ON a.shift_code = c.shift_code AND a.group_code = c.group_code " & _
'          "WHERE b.enrollnumber = '" & Val(TDBGrid_Att.Columns("nik").Value) & "' " & _
'            "AND CASE WHEN c.flag_day_over = 1 OR WEEKDAY('" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "') = 5 OR WEEKDAY('" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "') = 6 THEN " & _
'                    "(date(a.att_date) BETWEEN '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "' AND ADDDATE('" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "',1)) " & _
'                "ELSE date(a.att_date) = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "' END " & oClause

    SQL = "SELECT a.*, DATE(a.att_date) AS date_att, TIME(a.att_date) AS time_att," & _
             "CASE WHEN a.flag_io = 0 THEN 'IN' ELSE 'OUT' END TYPE " & _
          "FROM h_log_attendance a " & _
          "WHERE (a.enrollnumber = '" & Val(TDBGrid_Att.Columns("nik").Value) & "' " & _
                "OR a.employee_code = '" & Val(TDBGrid_Att.Columns("employee_code").Value) & "') " & _
            "AND CASE WHEN '" & Val(TDBGrid_Att.Columns("flag_day_over").Value) & "' = 1 OR WEEKDAY('" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "') = 5 OR WEEKDAY('" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "') = 6 THEN " & _
                    "(date(a.att_date) BETWEEN '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "' AND ADDDATE('" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "',1)) " & _
                "ELSE date(a.att_date) = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "' END " & _
          "ORDER BY a.att_date, a.flag_io"
    rsLogAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Log_Att.DataSource = rsLogAtt
End Sub

Private Sub TDBGrid_Att_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not IIf(TDBGrid_Att.Bookmark = "", 0, TDBGrid_Att.Bookmark) > 0 Then
        Set TDBGrid_Log_Att.DataSource = Nothing
        Exit Sub
    End If
    Call load_data_log_att
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False

    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid_Att.Columns
        Col.FilterText = ""
    Next Col
    rsAtt.Filter = adFilterNone
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

Private Sub grid_filter()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid_Att.Columns
    i = TDBGrid_Att.Col
    TDBGrid_Att.HoldFields

    rsAtt.Filter = getFilter()
    TDBGrid_Att.Col = i
    TDBGrid_Att.EditActive = True

    TDBGrid_Att.SelStart = Len(TDBGrid_Att.Columns(i).FilterText)
    If TDBGrid_Att.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid_Att.Col = i
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
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

        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "company_code = '" & COMPANY_CODE & "'", "department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))

        If LOGIN_LEVEL = 100 Then
            If TDBCombo_department = "" Then
                SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%')"
            Else
                SQL = "select nik,employee_name,employee_code " & _
                         "from m_employee " & _
                         "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' AND department_code = '" & TDBCombo_department.Text & "' " & _
                            "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                                "OR employee_name LIKE '%" & txt_nik.Text & "%')"
            End If
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee " & _
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

Private Sub cmdSave_Click()
Dim rsemp As New ADODB.Recordset
Dim aa As Integer
Dim a, b, c, d As String

'On Error GoTo Err
    tglAwal = DTPicker1.Value
    
    If ttin.Value = ttout.Value Then
        MsgBox "Invalid Check In Time...!!!", vbExclamation, headerMSG
        Exit Sub
    End If

    If Format(ttout.Value, "hh:mm") = "" Or Format(ttout.Value, "hh:mm") = "" Then
        MsgBox "Invalid Format Check In or Check Out Time....!!!", vbExclamation, headerMSG
        Exit Sub
    End If
    
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    SQL = "SELECT employee_code FROM m_employee WHERE employee_code = '" & txt_employee_code.Text & "'"
    rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rs2.RecordCount = 0 Then
        MsgBox "Invalid Employee Code...!!", vbCritical, headerMSG
        rs2.Close
        Exit Sub
    End If
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False Then
        SQL = "SELECT att_date,shift_code,shift_number,status FROM h_attendance WHERE employee_code = '" & txt_employee_code.Text & "' " _
            & "AND att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
        If rs2.RecordCount > 0 Then
            Dim v_shift As String
            v_shift = IIf(rs2!Status = "P", "PRESENT", IIf(rs2!Status = "A", "ALPHA", IIf(rs2!Status = "DT", "DUTY", _
                    IIf(rs2!Status = "OF", "OFF", IIf(rs2!Status = "PR", "PERMISSION", "Sakit Ada Surat Dokter")))))
            
            MsgBox "This Employee Is Already Exist With Status " & v_shift, vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Call InsertPresent
    
    MsgBox "Save Data Succesfully!", vbInformation, "Sukses!"
    
    Call clear_view_data
    
    fra_entry.Visible = False
    
    load_data_att
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub InsertPresent()
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String
Dim rsemp As New ADODB.Recordset
Dim vFlagRollable As Integer
Dim vFlagDayOver As Integer
Dim vGroupCode As String
    
    SQL = "SELECT group_code,flag_day_over FROM m_shift WHERE shift_code = '" & TDBCombo_shift.Text & "'"
    rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs2.RecordCount > 0 Then
        vFlagDayOver = rs2!flag_day_over
        vGroupCode = rs2!group_code
    End If
    rs2.Close
    
    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
    SQL = "Select CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(max_break_out)) as datetime) max_break_out," _
            & "CAST(concat('" & Format(tglAwal, "yyyy-MM-dd") & "',' ', time(min_break_in)) as datetime) min_break_in," _
            & "curdate() tglserver " _
            & "from m_shift where shift_code = '" & TDBCombo_shift.Text & "' " & _
                "AND group_code = '" & vGroupCode & "'"
'        & "from m_shift where shift_code = '" & txtkdshift.Text & "'"
    rs2.Open SQL, CnG, adOpenDynamic, adLockReadOnly
    
    start_time = Format(rs2!start_time, "yyyy-MM-dd hh:mm:ss")
    
    If vFlagDayOver = 1 Then
        end_time = Format(DateAdd("d", 1, rs2!end_time), "yyyy-MM-dd hh:mm:ss")
    Else
        end_time = Format(rs2!end_time, "yyyy-MM-dd hh:mm:ss")
    End If
    
    time_in = Format(tglAwal, "yyyy-MM-dd") & " " & Format(ttin.Value, "hh:mm") & ":00"
    
    If Format(ttout.Value, "hh:mm:ss") < Format(ttin.Value, "hh:mm:ss") Then
        time_out = Format(tglAwal + 1, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    Else
        time_out = Format(tglAwal, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    End If
    
    rs2.Close
    
    If editTrans = False Then
        SQL = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "' AND employee_code = '" & txt_employee_code.Text & "'"
        CnG.Execute SQL
           
        If ttin.ValueIsNull = True Then
            SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                    "shift_number,start_time,end_time,time_out,description,entry_date,userinput," & _
                    "absent_status,flag_present,flag_duty, flag_transport, flag_inc_late, flag_manual) " & _
                  "VALUES (" & _
                    "'" & txt_employee_code.Text & "','" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "','" & vGroupCode & "'," & _
                    "'" & TDBCombo_shift.Text & "','P'," & _
                    "1,'" & start_time & "','" & end_time & "','" & time_out & "'," & _
                    "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',0,1,0,'" & IIf(chkTrans.Value, -1, 0) & "'," & chkLate.Value & ",1)"
        ElseIf ttout.ValueIsNull = True Then
            SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                    "shift_number,start_time,end_time,time_in,description,entry_date,userinput," & _
                    "absent_status,flag_present,flag_duty, flag_transport, flag_inc_late, flag_manual) " & _
                  "VALUES (" & _
                    "'" & txt_employee_code.Text & "','" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "','" & vGroupCode & "'," & _
                    "'" & TDBCombo_shift.Text & "','P'," & _
                    "1,'" & start_time & "','" & end_time & "','" & time_in & "'," & _
                    "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',0,1,0,'" & IIf(chkTrans.Value, -1, 0) & "'," & chkLate.Value & ",1)"
        Else
            SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                    "shift_number,start_time,end_time,time_in,time_out,description,entry_date,userinput," & _
                    "absent_status,flag_present,flag_duty, flag_transport, flag_inc_late, flag_manual) " & _
                  "VALUES (" & _
                    "'" & txt_employee_code.Text & "','" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "','" & vGroupCode & "'," & _
                    "'" & TDBCombo_shift.Text & "','P'," & _
                    "1,'" & start_time & "','" & end_time & "','" & time_in & "','" & time_out & "'," & _
                    "'" & txt_description.Text & "',now(),'" & LOGIN_NAME & "',0,1,0,'" & IIf(chkTrans.Value, -1, 0) & "'," & chkLate.Value & ",1)"
        End If
    Else
        If ttin.ValueIsNull = True Then
            SQL = "UPDATE h_attendance set employee_code = '" & txt_employee_code.Text & "'," & _
                    "att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "',group_code = '" & vGroupCode & "'," & _
                    "shift_code = '" & TDBCombo_shift.Text & "'," & _
                    "status = 'P'," & _
                    "shift_number = 1,start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                    "flag_present = 1," & _
                    "time_out = '" & time_out & "'," & _
                    "description = '" & txt_description.Text & "',edit_date = now(),useredit = '" & LOGIN_NAME & "',absent_status = 0, flag_present = 1, flag_duty = 0, " & _
                    "flag_transport = '" & IIf(chkTrans.Value, -1, 0) & "', " & _
                    "flag_inc_late = " & chkLate.Value & "," & _
                    "flag_manual = 1 " & _
                    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                        "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & "'"
        ElseIf ttout.ValueIsNull = True Then
            SQL = "UPDATE h_attendance set employee_code = '" & txt_employee_code.Text & "'," & _
                    "att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "',group_code = '" & vGroupCode & "'," & _
                    "shift_code = '" & TDBCombo_shift.Text & "'," & _
                    "status = 'P'," & _
                    "shift_number = 1,start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                    "time_in = '" & time_in & "',flag_present = 1," & _
                    "description = '" & txt_description.Text & "',edit_date = now(),useredit = '" & LOGIN_NAME & "',absent_status = 0, flag_present = 1, flag_duty = 0, " & _
                    "flag_transport = '" & IIf(chkTrans.Value, -1, 0) & "'," & _
                    "flag_inc_late = " & chkLate.Value & "," & _
                    "flag_manual = 1 " & _
                    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                        "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & "'"
        Else
            SQL = "UPDATE h_attendance set employee_code = '" & txt_employee_code.Text & "'," & _
                    "att_date = '" & Format(tglAwal, "yyyy-MM-dd hh:mm:ss") & "',group_code = '" & vGroupCode & "'," & _
                    "shift_code = '" & TDBCombo_shift.Text & "'," & _
                    "status = 'P'," & _
                    "shift_number = 1,start_time = '" & start_time & "',end_time = '" & end_time & "'," & _
                    "time_in = '" & time_in & "',flag_present = 1," & _
                    "time_out = '" & time_out & "'," & _
                    "description = '" & txt_description.Text & "',edit_date = now(),useredit = '" & LOGIN_NAME & "',absent_status = 0, flag_present = 1, flag_duty = 0, " & _
                    "flag_transport = '" & IIf(chkTrans.Value, -1, 0) & "'," & _
                    "flag_inc_late = " & chkLate.Value & "," & _
                    "flag_manual = 1 " & _
                    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                        "AND att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd hh:mm:ss") & "'"
        End If
'                & "AND shift_code = '" & txtkdshift.Text & "'"
    End If
    CnG.Execute SQL
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    CnG.Execute "call spg_leave_periode2 ('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "')"

End Sub

Private Sub cmdCancel_Click()
    clear_view_data
    fra_entry.Visible = False
End Sub

Private Sub TDBGrid_Att_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Att_HeadClick(ByVal ColIndex As Integer)

    x = x + 1

    If x Mod 2 <> 1 And vSubject = TDBGrid_Att.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Att.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Att.Columns(ColIndex).DataField + " ASC"
    End If

    vSubject = TDBGrid_Att.Columns(ColIndex).DataField
    Call load_data_att

End Sub

Private Sub TDBGrid_Log_Att_HeadClick(ByVal ColIndex As Integer)

    x = x + 1

    If x Mod 2 <> 1 And vSubject = TDBGrid_Log_Att.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Log_Att.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Log_Att.Columns(ColIndex).DataField + " ASC"
    End If

    vSubject = TDBGrid_Log_Att.Columns(ColIndex).DataField
    Call load_data_log_att

End Sub
