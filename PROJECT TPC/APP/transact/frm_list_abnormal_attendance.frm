VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_list_abnormal_attendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIST UNKNOWN ATTENDANCE"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13020
   Icon            =   "frm_list_abnormal_attendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   1200
      TabIndex        =   9
      Top             =   1350
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
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1050
      Width           =   1305
   End
   Begin VB.TextBox txt_employee_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2910
      TabIndex        =   10
      Top             =   1050
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2550
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   0
      Top             =   690
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker_from 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1740
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyyy"
      Format          =   97583107
      CurrentDate     =   40794
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Att 
      Height          =   7035
      Left            =   90
      TabIndex        =   6
      Top             =   2280
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   12409
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
      Columns.Count   =   18
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=18"
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
      Caption         =   "LIST OF UNKNOWN ATTENDANCE"
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
      _StyleDefs(106) =   "Named:id=33:Normal"
      _StyleDefs(107) =   ":id=33,.parent=0"
      _StyleDefs(108) =   "Named:id=34:Heading"
      _StyleDefs(109) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(110) =   ":id=34,.wraptext=-1"
      _StyleDefs(111) =   "Named:id=35:Footing"
      _StyleDefs(112) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   "Named:id=36:Selected"
      _StyleDefs(114) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=37:Caption"
      _StyleDefs(116) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(117) =   "Named:id=38:HighlightRow"
      _StyleDefs(118) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(119) =   "Named:id=39:EvenRow"
      _StyleDefs(120) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(121) =   "Named:id=40:OddRow"
      _StyleDefs(122) =   ":id=40,.parent=33"
      _StyleDefs(123) =   "Named:id=41:RecordSelector"
      _StyleDefs(124) =   ":id=41,.parent=34"
      _StyleDefs(125) =   "Named:id=42:FilterBar"
      _StyleDefs(126) =   ":id=42,.parent=33"
   End
   Begin prj_tpc.vbButton cmdSearch 
      Height          =   465
      Left            =   4410
      TabIndex        =   7
      Top             =   1650
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
      MICON           =   "frm_list_abnormal_attendance.frx":058A
      PICN            =   "frm_list_abnormal_attendance.frx":05A6
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
      Left            =   1200
      OleObjectBlob   =   "frm_list_abnormal_attendance.frx":1638
      TabIndex        =   1
      Top             =   690
      Width           =   1305
   End
   Begin prj_tpc.vbButton cmdBrowse 
      Height          =   285
      Left            =   2550
      TabIndex        =   8
      Top             =   1050
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
      MICON           =   "frm_list_abnormal_attendance.frx":359E
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
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6330
      TabIndex        =   12
      Top             =   1050
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker DTPicker_to 
      Height          =   315
      Left            =   2940
      TabIndex        =   14
      Top             =   1740
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyyy"
      Format          =   97583107
      CurrentDate     =   40794
   End
   Begin MSComCtl2.DTPicker DTPicker_Periode 
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   1380
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM-yyyy"
      Format          =   97583107
      CurrentDate     =   40794
   End
   Begin VB.Label Label4 
      Caption         =   "PERIODE"
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   1410
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   15
      Top             =   1770
      Width           =   285
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "EMPLOYEE"
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   1080
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LIST UNKNOWN ATTENDANCE"
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
      Left            =   390
      TabIndex        =   4
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "DATE"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1770
      Width           =   645
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   0
      Picture         =   "frm_list_abnormal_attendance.frx":35BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_list_abnormal_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_from, DTPicker_to)
End Sub

Private Sub Form_Load()
    Call load_data_company
    
    Call createGridKar
    timer1.Enabled = True

    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
    DTPicker_Periode.Value = Now
End Sub

Public Sub load_data_company()
    TDBCombo_company.Text = "": txt_company_name = ""
    
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Public Sub load_data_att()
    If rsAtt.State Then rsAtt.Close
     
    SQL = "select a.att_date,a.employee_code,b.nik,b.employee_name,a.status,e.absent_name,b.title_code,c.title_name, " & _
                "a.time_in,a.time_out,a.entry_date,b.division_code,d.division_name,a.description, " & _
                "b.department_code, f.department_name, a.shift_code, g.shift_name " & _
              "from h_attendance a join m_employee b on a.employee_code = b.employee_code " & _
                "left join m_title c on b.title_code = c.title_code " & _
                "left join m_division d on b.division_code = d.division_code and b.department_code = d.department_code and b.company_code = d.company_code " & _
                "join m_absent_status e on a.status = e.absent_code " & _
                "left join m_department f on b.department_code = f.department_code and b.company_code = f.company_code " & _
                "left join m_shift g on a.shift_code = g.shift_code " & _
              "where (date(a.att_date) between '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' and '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & _
                "AND " & IIf(txt_nik.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                    "a.employee_code = '" & txt_employee_code.Text & "' AND b.company_code = '" & TDBCombo_company.Text & "'") & " " & _
                "AND (IFNULL((HOUR(SEC_TO_TIME(TIME_TO_SEC(TIMEDIFF(time_out,time_in)))) " & _
                        "+ (MINUTE(SEC_TO_TIME(TIME_TO_SEC(TIMEDIFF(time_out,time_in))))/60)),0) > 9.5 " & _
                "OR IFNULL((HOUR(SEC_TO_TIME(TIME_TO_SEC(TIMEDIFF(time_out,time_in)))) " & _
                        "+ (MINUTE(SEC_TO_TIME(TIME_TO_SEC(TIMEDIFF(time_out,time_in))))/60)),0) < 9.5)"
    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Att.DataSource = rsAtt
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
    End If
End Sub

Private Sub cmdSearch_Click()
    Call load_data_att
End Sub

Private Sub TDBGrid_Att_FilterChange()
    Call grid_filter
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
'On Error GoTo Err

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

