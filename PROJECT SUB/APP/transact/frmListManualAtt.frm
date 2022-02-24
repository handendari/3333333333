VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_list_manual_att 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANUAL ATTENDANCE"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "frmListManualAtt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin prj_absensi.LynxGrid LynxGrid1 
      Height          =   2925
      Left            =   6240
      TabIndex        =   2
      Top             =   1290
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   1650
      Width           =   3435
   End
   Begin VB.Frame Frame1 
      Height          =   7845
      Left            =   270
      TabIndex        =   18
      Top             =   2160
      Width           =   14865
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Shift 
         Height          =   5325
         Left            =   60
         TabIndex        =   19
         Top             =   150
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   9393
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "SHIFT CODE"
         Columns(0).DataField=   "shift_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "SHIFT NAME"
         Columns(1).DataField=   "shift_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FLAG DAY OVER"
         Columns(2).DataField=   "flag_day_over"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3545"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3466"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
         Caption         =   "LIST OF SHIFT"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Att 
         Height          =   7635
         Left            =   3750
         TabIndex        =   20
         Top             =   150
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   13467
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
         Columns(1).DataField=   "tgl"
         Columns(1).NumberFormat=   "yyyy-MM-dd"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ATT DATE"
         Columns(2).DataField=   "att_date"
         Columns(2).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "EMP. CODE"
         Columns(3).DataField=   "employee_code"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "EMP. NAME"
         Columns(4).DataField=   "employee_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "STATUS CODE"
         Columns(5).DataField=   "status"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TITLE CODE"
         Columns(6).DataField=   "title_code"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "JOB TITLE"
         Columns(7).DataField=   "title_name"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TIME IN"
         Columns(8).DataField=   "time_in"
         Columns(8).NumberFormat=   "hh:mm"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "TIME OUT"
         Columns(9).DataField=   "time_out"
         Columns(9).NumberFormat=   "hh:mm"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "SHIFT"
         Columns(10).DataField=   "shift_code"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "LATE"
         Columns(11).DataField=   "late"
         Columns(11).NumberFormat=   "HH:mm"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "EARLY"
         Columns(12).DataField=   "early"
         Columns(12).NumberFormat=   "HH:mm"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "STATUS"
         Columns(13).DataField=   "status"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "ENTRY DATE"
         Columns(14).DataField=   "entry_date"
         Columns(14).NumberFormat=   "yyyy-MM-dd"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "DESCRIPTION"
         Columns(15).DataField=   "description"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "STATUS NAME"
         Columns(16).DataField=   "absent_name"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "FLAG LATE"
         Columns(17).DataField=   "flag_inc_late"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "SHIFT CODE"
         Columns(18).DataField=   "shift"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "ENROLLNUMBER"
         Columns(19).DataField=   "enrollnumber"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   20
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=20"
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
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2355"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2275"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1667"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1588"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=4128"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=4048"
         Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=1614"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1535"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=185"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=106"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(44)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(46)=   "Column(7)._MinWidth=10"
         Splits(0)._ColumnProps(47)=   "Column(8).Width=1402"
         Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1323"
         Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(52)=   "Column(8)._MinWidth=54215968"
         Splits(0)._ColumnProps(53)=   "Column(9).Width=1640"
         Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=1561"
         Splits(0)._ColumnProps(56)=   "Column(9)._ColStyle=513"
         Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(58)=   "Column(9)._MinWidth=54215968"
         Splits(0)._ColumnProps(59)=   "Column(10).Width=1244"
         Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=1164"
         Splits(0)._ColumnProps(62)=   "Column(10)._ColStyle=513"
         Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(64)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=1693"
         Splits(0)._ColumnProps(67)=   "Column(11)._ColStyle=513"
         Splits(0)._ColumnProps(68)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(69)=   "Column(12).Width=1773"
         Splits(0)._ColumnProps(70)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(12)._WidthInPix=1693"
         Splits(0)._ColumnProps(72)=   "Column(12)._ColStyle=513"
         Splits(0)._ColumnProps(73)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(74)=   "Column(13).Width=2434"
         Splits(0)._ColumnProps(75)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(13)._WidthInPix=2355"
         Splits(0)._ColumnProps(77)=   "Column(13)._ColStyle=513"
         Splits(0)._ColumnProps(78)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(79)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(80)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(82)=   "Column(14)._ColStyle=516"
         Splits(0)._ColumnProps(83)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(85)=   "Column(15).Width=4366"
         Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=4286"
         Splits(0)._ColumnProps(88)=   "Column(15)._ColStyle=516"
         Splits(0)._ColumnProps(89)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(90)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(91)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(92)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(93)=   "Column(16)._ColStyle=516"
         Splits(0)._ColumnProps(94)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(95)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(96)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(97)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(98)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(99)=   "Column(17)._ColStyle=516"
         Splits(0)._ColumnProps(100)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(101)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(102)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(103)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(104)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(105)=   "Column(18)._ColStyle=516"
         Splits(0)._ColumnProps(106)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(107)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(108)=   "Column(19).Width=2725"
         Splits(0)._ColumnProps(109)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(110)=   "Column(19)._WidthInPix=2646"
         Splits(0)._ColumnProps(111)=   "Column(19)._ColStyle=516"
         Splits(0)._ColumnProps(112)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(113)=   "Column(19).Order=20"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=102,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=110,.parent=13,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=158,.parent=13,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=155,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=156,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=157,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=130,.parent=13,.alignment=2"
         _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=127,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=128,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=129,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=138,.parent=13,.alignment=2"
         _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=135,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=136,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=137,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=28,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=25,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=26,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=27,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=82,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=17"
         _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=86,.parent=13"
         _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=83,.parent=14"
         _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=84,.parent=15"
         _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=85,.parent=17"
         _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=114,.parent=13"
         _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=111,.parent=14"
         _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=112,.parent=15"
         _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=113,.parent=17"
         _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=94,.parent=13"
         _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=91,.parent=14"
         _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=92,.parent=15"
         _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=93,.parent=17"
         _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=146,.parent=13"
         _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=143,.parent=14"
         _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=144,.parent=15"
         _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=145,.parent=17"
         _StyleDefs(114) =   "Named:id=33:Normal"
         _StyleDefs(115) =   ":id=33,.parent=0"
         _StyleDefs(116) =   "Named:id=34:Heading"
         _StyleDefs(117) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(118) =   ":id=34,.wraptext=-1"
         _StyleDefs(119) =   "Named:id=35:Footing"
         _StyleDefs(120) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(121) =   "Named:id=36:Selected"
         _StyleDefs(122) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(123) =   "Named:id=37:Caption"
         _StyleDefs(124) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(125) =   "Named:id=38:HighlightRow"
         _StyleDefs(126) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(127) =   "Named:id=39:EvenRow"
         _StyleDefs(128) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(129) =   "Named:id=40:OddRow"
         _StyleDefs(130) =   ":id=40,.parent=33"
         _StyleDefs(131) =   "Named:id=41:RecordSelector"
         _StyleDefs(132) =   ":id=41,.parent=34"
         _StyleDefs(133) =   "Named:id=42:FilterBar"
         _StyleDefs(134) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2265
         Left            =   60
         TabIndex        =   21
         Top             =   5520
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   3995
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
         Columns(1).NumberFormat=   "dd/MM/yyyy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "TIME"
         Columns(2).DataField=   "time_att"
         Columns(2).NumberFormat=   "HH:mm:ss"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "ATT DATE"
         Columns(3).DataField=   "att_date"
         Columns(3).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2778"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2699"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2619"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2540"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=3678"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3598"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Advanced Filter"
      Height          =   1395
      Index           =   0
      Left            =   5130
      TabIndex        =   7
      Top             =   720
      Width           =   5025
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         DragMode        =   1  'Automatic
         Height          =   285
         Left            =   2370
         TabIndex        =   9
         Top             =   270
         Width           =   2535
      End
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         TabIndex        =   8
         Top             =   270
         Width           =   855
      End
      Begin prj_absensi.vbButton cmdBrowse_Emp 
         Height          =   285
         Left            =   2010
         TabIndex        =   10
         Top             =   270
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
         MICON           =   "frmListManualAtt.frx":058A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker_from 
         Height          =   315
         Left            =   1110
         TabIndex        =   11
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   85590019
         CurrentDate     =   40794
      End
      Begin MSComCtl2.DTPicker DTPicker_to 
         Height          =   315
         Left            =   2790
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   85590019
         CurrentDate     =   40794
      End
      Begin prj_absensi.vbButton cmdSearch 
         Height          =   705
         Left            =   4140
         TabIndex        =   13
         Top             =   600
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1244
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
         MICON           =   "frmListManualAtt.frx":05A6
         PICN            =   "frmListManualAtt.frx":05C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "EMPLOYEE"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   870
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
         Index           =   0
         Left            =   2430
         TabIndex        =   15
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DATE"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   750
         Width           =   435
      End
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   30
      Top             =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_absensi.vbButton cmdExit 
      Height          =   705
      Left            =   14100
      TabIndex        =   1
      Top             =   10050
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
      MICON           =   "frmListManualAtt.frx":1654
      PICN            =   "frmListManualAtt.frx":1670
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   10200
      TabIndex        =   3
      Top             =   1380
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prj_absensi.vbButton cmdDelete 
      Height          =   465
      Left            =   12900
      TabIndex        =   4
      Top             =   1560
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
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
      MICON           =   "frmListManualAtt.frx":2702
      PICN            =   "frmListManualAtt.frx":271E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_absensi.vbButton cmdEdit 
      Height          =   465
      Left            =   11550
      TabIndex        =   5
      Top             =   1560
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
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
      MICON           =   "frmListManualAtt.frx":37B0
      PICN            =   "frmListManualAtt.frx":37CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_absensi.vbButton cmdNew 
      Height          =   465
      Left            =   10200
      TabIndex        =   6
      Top             =   1560
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
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
      MICON           =   "frmListManualAtt.frx":485E
      PICN            =   "frmListManualAtt.frx":487A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_absensi.vbButton cmdPrev 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   840
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "<"
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
      MICON           =   "frmListManualAtt.frx":590C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1530
      TabIndex        =   23
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   85590019
      CurrentDate     =   40794
   End
   Begin prj_absensi.vbButton cmdRefresh 
      Height          =   495
      Left            =   3690
      TabIndex        =   24
      Top             =   750
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Refresh"
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
      MICON           =   "frmListManualAtt.frx":5928
      PICN            =   "frmListManualAtt.frx":5944
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_absensi.vbButton cmdNext 
      Height          =   285
      Left            =   3210
      TabIndex        =   25
      Top             =   840
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   ">"
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
      MICON           =   "frmListManualAtt.frx":69D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_absensi.vbButton cmd_reproccess 
      Height          =   675
      Left            =   10200
      TabIndex        =   26
      Top             =   690
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1191
      BTYPE           =   14
      TX              =   "Reproccess"
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
      MICON           =   "frmListManualAtt.frx":69F2
      PICN            =   "frmListManualAtt.frx":6A0E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1530
      OleObjectBlob   =   "frmListManualAtt.frx":7AA0
      TabIndex        =   27
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "DATE"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MANUAL ATTENDANCE"
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
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   150
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frmListManualAtt.frx":9EBA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16860
   End
End
Attribute VB_Name = "frm_list_manual_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Dim rs As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsShift As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset

Dim rsLogAtt As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim vMode As Integer
Dim vParam As String

Dim vAttDate As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    vMode = 0
    DTPicker1.Value = DTPicker1 + 1

    Call load_data_att
End Sub

Private Sub cmdPrev_Click()
    vMode = 0
    DTPicker1.Value = DTPicker1 - 1

    Call load_data_att
End Sub

Private Sub cmdSearch_Click()
    'validasi company
    If Trim(TDBCombo_company.Text) = "" Then
        MsgBox "Company not selected!", vbOKOnly + vbInformation, headerMSG
        TDBCombo_company.SetFocus
        Exit Sub
    End If
    
    vMode = 1
            
    Call load_data_att
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    vMode = 0
    Call load_data_att
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    
    Call createGridKar
    Call load_data_company
    
    oClause = ""
    vMode = 0
    
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
    
    timer1.Enabled = True
    
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_shift()
    If rsShift.State Then rsShift.Close
    SQL = "select * from m_shift " & _
          "WHERE flag_shift = 0 " & _
          "order by shift_code"
    rsShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    Set TDBGrid_Shift.DataSource = rsShift
End Sub

Public Sub load_data_att()
Dim vParameter As String
Dim vFlagRollable As Integer
Dim x As Integer
                            
    If rsAtt.State Then rsAtt.Close
    vParameter = IIf(txt_employee_code.Text <> "", _
                    "a.employee_code = '" & txt_employee_code.Text & "' AND (DATE(b.att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') ", _
                    "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND (DATE(b.att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') ")
                    
    SQL = "SELECT DISTINCT b.att_date,b.att_date tgl,a.employee_code," & _
                "a.employee_name, a.title_code, c.title_name, b.time_in, b.time_out, b.entry_date," & _
                "b.description, d.shift_name, b.shift_code, " & _
                "f_actual_status(IFNULL(b.flag_present,0), IFNULL(b.flag_duty,0), IFNULL(b.absent_status,0)) status, " & _
                "b.enrollnumber, " & _
                "f_get_late(b.start_time,b.time_in) late," & _
                "f_get_early(b.end_time,b.time_out) early " & _
          "FROM m_employee a JOIN h_attendance b ON a.employee_code = b.employee_code " & _
                "LEFT JOIN m_title c ON a.title_code = c.title_code " & _
                "LEFT JOIN m_shift d ON b.shift_code = d.shift_code " & _
                "LEFT JOIN td_shift f ON a.employee_code = f.employee_code "
                
    If vMode = 0 Then
        vParam = IIf(LOGIN_LEVEL <> 100, _
                IIf(COMPANY_ACCESS = 0, _
                    "a.company_code = '" & COMPANY_CODE & "' ", _
                            "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' "), "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' ")
                 
        SQL = SQL & _
                "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND date(b.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                    "AND b.shift_code = '" & TDBGrid_Shift.Columns("shift_code").Value & "' " & _
                "order by a.employee_name, b.att_date "
    Else
        SQL = SQL & _
                "where " & vParameter & " ORDER BY a.employee_name, b.att_date "
    End If
    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Att.DataSource = rsAtt
    
    cmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
    cmdEdit.Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    cmdDelete.Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    cmdRefresh.Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    
    Call load_data_user_access(Me)
    cmdNew.Enabled = blnUser_Add
    cmdEdit.Enabled = blnUser_Edit
    cmdDelete.Enabled = blnUser_Delete
End Sub

Private Sub load_data_log_att()
    If rsLogAtt.State Then rsLogAtt.Close
    SQL = "SELECT a.*, DATE(a.att_date) AS date_att, TIME(a.att_date) AS time_att," & _
             "CASE WHEN a.flag_io = 0 THEN 'IN' " & _
                  "WHEN a.flag_io = 1 THEN 'OUT' " & _
             "END " & _
          "FROM h_log_attendance a " & _
          "WHERE a.enrollnumber = '" & TDBGrid_Att.Columns("enrollnumber").Value & "' " & _
            "AND date(a.att_date) = '" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "' " & _
          "ORDER BY a.att_date, a.flag_io"
    rsLogAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid1.DataSource = rsLogAtt
End Sub

Private Sub showData()
Dim s As String
Dim vFlagLate As Integer
Dim vTimeIn, vTimeOut As String
    If rscari.State Then rscari.Close
    SQL = "SELECT start_time, end_time FROM m_working_day WHERE shift_code = '" & TDBGrid_Shift.Columns("shift_code").Value & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vTimeIn = Format(rscari!start_time, "HH:mm")
        vTimeOut = Format(rscari!end_time, "HH:mm")
    Else
        vTimeIn = "00:00"
        vTimeOut = "00:00"
    End If
    rscari.Close
    
    If TDBGrid_Att.ApproxCount > 0 Then
        With frm_trans_att_man
            .load_data_shift
            Call .set_data_shift(IIf(TDBGrid_Att.Columns("shift_code").Value = "", TDBGrid_Shift.Columns("shift_code").Value, TDBGrid_Att.Columns("shift_code").Value))
            .DTPicker1.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
            .txt_employee_code.Text = IIf(IsNull(TDBGrid_Att.Columns("employee_code").Value), txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
            .txt_employee_name.Text = IIf(IsNull(TDBGrid_Att.Columns("employee_name").Value), txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
                        
            .txt_title_code.Text = IIf(IsNull(TDBGrid_Att.Columns("title_code").Value), "", TDBGrid_Att.Columns("title_code").Value)
            .txt_title_name.Text = IIf(IsNull(TDBGrid_Att.Columns("title_name").Value), "", TDBGrid_Att.Columns("title_name").Value)
                        
            .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
            .DTPicker_to.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
            
            If TDBGrid_Att.Columns("att_date").Value = "" Then
                .DTPicker_from.Enabled = True
                .DTPicker_to.Enabled = True
                
                .ttin.Value = vTimeIn
                .ttout.Value = vTimeOut
            Else
                .DTPicker_from.Enabled = False
                .DTPicker_to.Enabled = False
                                
                .ttin.Value = Format(TDBGrid_Att.Columns("time_in").Value, "hh:mm")
                .ttout.Value = Format(TDBGrid_Att.Columns("time_out").Value, "hh:mm")
            End If
                        
            .txt_description = IIf(IsNull(TDBGrid_Att.Columns("description").Value), "", TDBGrid_Att.Columns("description").Value)
                                                
            .chk_all_employee.Enabled = False
                        
            vAttDate = TDBGrid_Att.Columns("att_date").Value
            If vAttDate = "" Then
                .editTrans = False
            Else
                .editTrans = True
            End If
            
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_att_man
        .load_data_shift
        Call .set_data_shift(TDBGrid_Shift.Columns("shift_code").Value)
        .DTPicker1.Value = DTPicker1.Value
        .DTPicker_from.Value = DTPicker1.Value
        .DTPicker_to.Value = DTPicker1.Value
        .vShiftCode = TDBGrid_Shift.Columns("shift_code").Value
        
        .chkSat.Value = 0
        .chkSun.Value = 0
        
        .editTrans = False
        .Show vbModal
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_list_manual_att = Nothing
End Sub

Private Sub TDBGrid_Att_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (TDBGrid_Att.Row + 1) > 0 And (TDBGrid_Att.Row + 1) <> LastRow Then
        'MsgBox "LETS..."
        Call load_data_log_att
    End If
End Sub

Private Sub TDBGrid_Shift_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call load_data_att
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value

        Call load_data_shift
        Call load_data_att
    End If
End Sub

Private Sub CmdNew_Click()
    If TDBGrid_Shift.ApproxCount = 0 Then
        MsgBox "Invalid Shift Code!" & Chr(13) & "Please Check Your Transaction Again...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    newData
End Sub

Private Sub cmdEdit_Click()
    If TDBGrid_Att.Columns("status").Text <> "Present" Then
        MsgBox "You couldn't edit this data because it's not Present status...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    showData
End Sub

Private Sub cmdRefresh_Click()
    vMode = 0
    Call load_data_att
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode_rs(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Att.Columns
            Col.FilterText = ""
        Next Col
        rsAtt.Filter = adFilterNone
    ElseIf SSTab1.Tab = 3 Then
        For Each Col In TDBGrid_Sum.Columns
            Col.FilterText = ""
        Next Col
        rsSumAtt.Filter = adFilterNone
    End If
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
            
            tmp = tmp & Col.DataField & " LIKE '*" & Col.FilterText & "*'"
        End If
    Next Col
    getFilter = tmp
End Function

Private Sub TDBGrid_Att_FilterChange()
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

Private Sub TDBGrid_Sum_FilterChange()
On Error GoTo Err

    Dim i As Integer
    
    Set Cols = TDBGrid_Sum.Columns
    i = TDBGrid_Sum.Col
    TDBGrid_Sum.HoldFields
    
    rsSumAtt.Filter = getFilter()
    TDBGrid_Sum.Col = i
    TDBGrid_Sum.EditActive = True
    
    TDBGrid_Sum.SelStart = Len(TDBGrid_Sum.Columns(i).FilterText)
    If TDBGrid_Sum.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid_Sum.Col = i
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim j As Integer
Dim item
    
On Error GoTo Err
    If TDBGrid_Att.Columns("status").Text <> "Present" Then
        MsgBox "You couldn't delete this data because it's not Present status...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid_Att.SelBookmarks
    i = MsgBox("Are you sure want to delete " _
        & SelBks.Count & " attendance's data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        SQL = "DELETE FROM h_attendance where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
            & "and att_date = '" & Format(TDBGrid_Att.Columns("att_date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM t_absent where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
                & "AND DATE(absent_date_from) = '" & Format(TDBGrid_Att.Columns("tgl").CellText(item), "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM t_check where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
                & "AND DATE(check_date) = '" & Format(TDBGrid_Att.Columns("tgl").CellText(item), "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM t_leave where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
                & "AND DATE(leave_date_from) = '" & Format(TDBGrid_Att.Columns("tgl").CellText(item), "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        i = i + 1
    Next
    CnG.CommitTrans
    Call load_data_att
    MsgBox i & " attendance's data are successfully deleted..", vbInformation, headerMSG
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub createGrid()
    LynxGrid1.ClearAll
    
    With LynxGrid1
       .AddColumn "ATT. DATE", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd", , , , , , True
       .AddColumn "EMP. CODE", 1300, lgAlignCenterCenter, , , , , , , True
       .AddColumn "EMP. NAME", 2500, , , , , , , , True
       .AddColumn "GROUP CODE", 800, lgAlignCenterCenter, , , , , , , , True
       .AddColumn "SHIFT CODE", 800, , , , , , , , , True
       .AddColumn "ATT. STATUS", 800, lgAlignCenterCenter, , , , , , , True
       .AddColumn "TIME IN", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
       .AddColumn "TIME OUT", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
       .AddColumn "DESCRIPTION", 3000, lgAlignCenterCenter, , , , , , , True
       .BackColorBkg = &HFCE1CB
       .Redraw = True
    End With
End Sub

Private Sub createGridKar()
    LynxGrid1.ClearAll
    
    With LynxGrid1
       .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
       .AddColumn "Name", 3000, , , , , , , , , True
       .BackColorBkg = &HFCE1CB
       .Redraw = True
    End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid1.Clear
         
        If rs.State Then rs.Close
        SQL = "select employee_name,employee_code " & _
                "from m_employee a " & _
                "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                    "AND flag_active <> 0 " & _
                    "AND (employee_code LIKE '%" & txt_employee_code.Text & "%' " & _
                    "OR employee_name LIKE '%" & txt_employee_code.Text & "%')"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid1.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid1.AddItem rs!employee_code & vbTab & rs!EMPLOYEE_NAME
                rs.MoveNext
            Wend
            LynxGrid1.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
'                TDBCombo1.SetFocus
            Else
                LynxGrid1.Visible = True
                LynxGrid1.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid1.Rows > 0 Then
            txt_employee_code.Text = LynxGrid1.CellText(LynxGrid1.Row, 0)
            txt_employee_name.Text = LynxGrid1.CellText(LynxGrid1.Row, 1)
        End If
        LynxGrid1.Visible = False
    End If
End Sub

Private Sub LynxGrid1_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid1.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid1_LostFocus()
    LynxGrid1.Visible = False
End Sub

Private Sub txt_employee_code_Change()
    If txt_employee_code.Text = "" Then
        txt_employee_name.Text = ""
    End If
End Sub

Private Sub txt_employee_code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Emp_Click()
    isiGridKar (1)
End Sub

Private Sub cmd_reproccess_Click()
Dim strsql As String
Dim vStatus As String

'On Error Resume Next

    If txt_employee_code.Text = "" Then
        MsgBox "Employee is empty...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    '+++++++++++++++++++++++++++++ Reproccess Attendance ++++++++++++++++++++++++++++++++
    SQL = "DELETE FROM h_attendance_reproccess"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_log_attendance_recover"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_log_attendance_reproccess"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_attendance_reproccess " & _
          "SELECT * FROM h_attendance " & _
          "WHERE IFNULL(flag_present,0) <> 1 " & _
              "AND (DATE(att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " & _
                  "AND '" & Format(DateAdd("d", 1, DTPicker_to.Value), "yyyy-MM-dd") & "') " & _
                  "AND employee_code = '" & txt_employee_code.Text & "'"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_log_attendance_recover(att_date, ip_address, enrollnumber, verifymode, " & _
                "flag_io, flag_attendance, entry_date) " & _
            "SELECT DISTINCT att_date, a.ip_address, a.enrollnumber, verifymode, flag_io, " & _
                "flag_attendance, entry_date " & _
            "FROM h_log_attendance a JOIN m_enroll_link b ON a.enrollnumber = b.enrollnumber AND a.ip_address = b.ip_address  " & _
            "WHERE (date(a.att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " & _
                "AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & _
                "AND b.employee_code = '" & txt_employee_code.Text & "' " & _
            "ORDER BY att_date, flag_io"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_attendance " & _
            "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " & _
                "AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
    CnG.Execute SQL
    
    If rs.State Then rs.Close
    SQL = "SELECT * from h_log_attendance_recover ORDER BY att_date, flag_io"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    ProgressBar1.Visible = True
    ProgressBar1.Value = 0
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            ProgressBar1.Max = rs.RecordCount
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            SQL = "INSERT INTO h_log_attendance_reproccess(att_date, ip_address, enrollnumber, verifymode, " & _
                            "flag_io, flag_attendance, entry_date) " & _
                    "VALUES (" & _
                        "'" & Format(rs!att_date, "yyyy-MM-dd hh:nn:ss") & "', '" & rs!ip_address & "', '" & IIf(IsNull(rs!enrollnumber), 0, rs!enrollnumber) & "', " & _
                        "'" & IIf(IsNull(rs!verifymode), 0, rs!verifymode) & "', '" & rs!flag_io & "', " & _
                        "'" & IIf(IsNull(rs!flag_attendance), 0, rs!flag_attendance) & "', " & _
                        "'" & Format(IIf(IsNull(rs!entry_date), 0, rs!entry_date), "yyyy-MM-dd hh:nn:ss") & "')"
            CnG.Execute SQL
        rs.MoveNext
        Wend
    End If
    rs.Close
    
    SQL = "DELETE a FROM h_attendance a JOIN h_attendance_reproccess b ON a.employee_code = b.employee_code AND DATE(a.att_date) = DATE(b.att_date)"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_attendance " & _
          "SELECT * FROM h_attendance_reproccess"
    CnG.Execute SQL
    
    Screen.MousePointer = vbDefault
    
    ProgressBar1.Visible = False
    MsgBox "Proccess Successfully!", vbInformation, headerMSG
End Sub

Private Sub TDBCombo_company_Change()
    If TDBCombo_company.Text = "" Then txt_company_name.Text = ""
End Sub
