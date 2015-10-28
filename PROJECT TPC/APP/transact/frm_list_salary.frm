VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_list_salary 
   Caption         =   "LIST SALARY"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frm_list_salary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5415
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9551
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "MONTH"
      Columns(0).DataField=   "month"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DATE FROM"
      Columns(1).DataField=   "date_from"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DATE TO"
      Columns(2).DataField=   "date_to"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "EMPLOYEE CODE"
      Columns(3).DataField=   "employee_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ID EMP. CODE"
      Columns(4).DataField=   "nik"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "EMPLOYEE NAME"
      Columns(5).DataField=   "employee_name"
      Columns(5).NumberFormat=   "General Number"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "BASIC SALARY"
      Columns(6).DataField=   "basic_salary"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "LEAVE (HRS)"
      Columns(7).DataField=   "leave"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "RECEIVED SALARY"
      Columns(8).DataField=   "received_salary"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "ATTENDANT"
      Columns(9).DataField=   "attendant"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "OVERTIME"
      Columns(10).DataField=   "overtime"
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "POSITION ALLOW."
      Columns(11).DataField=   "position_allow"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "OTHER ALLOW."
      Columns(12).DataField=   "other_allow"
      Columns(12).NumberFormat=   "Standard"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "SHIFT ALLOW."
      Columns(13).DataField=   "shift_allow"
      Columns(13).NumberFormat=   "Standard"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "MEAL ALLOW."
      Columns(14).DataField=   "meal_allow"
      Columns(14).NumberFormat=   "Standard"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "TRANSPORT ALLOW."
      Columns(15).DataField=   "transport_allow"
      Columns(15).NumberFormat=   "Standard"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "GROSS INCOME"
      Columns(16).DataField=   "gross_income"
      Columns(16).NumberFormat=   "Standard"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "JMS"
      Columns(17).DataField=   "jms"
      Columns(17).NumberFormat=   "Standard"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "TAX STATUS"
      Columns(18).DataField=   "tax_status"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "INCOME TAX"
      Columns(19).DataField=   "income_tax"
      Columns(19).NumberFormat=   "Standard"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "TAX ADJ."
      Columns(20).DataField=   "tax_adj"
      Columns(20).NumberFormat=   "Standard"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "COOP. CONTR."
      Columns(21).DataField=   "coop_contr"
      Columns(21).NumberFormat=   "Standard"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "COOP. INSTALL"
      Columns(22).DataField=   "coop_install"
      Columns(22).NumberFormat=   "Standard"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "ACTUAL RECEIVED"
      Columns(23).DataField=   "actual"
      Columns(23).NumberFormat=   "Standard"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "BANK ACCOUNT"
      Columns(24).DataField=   "bank_account"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "COMPANY CODE"
      Columns(25).DataField=   "company_code"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   26
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=26"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2037"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1958"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=4207"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=4128"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=3281"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=3201"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=1931"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=1852"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=3069"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2990"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(59)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(64)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(65)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(67)=   "Column(12)._ColStyle=514"
      Splits(0)._ColumnProps(68)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(69)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(13)._ColStyle=514"
      Splits(0)._ColumnProps(73)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(74)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(75)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(77)=   "Column(14)._ColStyle=514"
      Splits(0)._ColumnProps(78)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(79)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(80)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(82)=   "Column(15)._ColStyle=514"
      Splits(0)._ColumnProps(83)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(84)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(85)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(16)._WidthInPix=3201"
      Splits(0)._ColumnProps(87)=   "Column(16)._ColStyle=514"
      Splits(0)._ColumnProps(88)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(89)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(17)._ColStyle=514"
      Splits(0)._ColumnProps(93)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(94)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(95)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(97)=   "Column(18)._ColStyle=513"
      Splits(0)._ColumnProps(98)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(99)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(100)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(102)=   "Column(19)._ColStyle=514"
      Splits(0)._ColumnProps(103)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(104)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(105)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(107)=   "Column(20)._ColStyle=514"
      Splits(0)._ColumnProps(108)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(109)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(110)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(21)._ColStyle=514"
      Splits(0)._ColumnProps(113)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(114)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(115)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(117)=   "Column(22)._ColStyle=514"
      Splits(0)._ColumnProps(118)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(119)=   "Column(23).Width=3625"
      Splits(0)._ColumnProps(120)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(23)._WidthInPix=3545"
      Splits(0)._ColumnProps(122)=   "Column(23)._ColStyle=514"
      Splits(0)._ColumnProps(123)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(124)=   "Column(24).Width=3466"
      Splits(0)._ColumnProps(125)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(24)._WidthInPix=3387"
      Splits(0)._ColumnProps(127)=   "Column(24)._ColStyle=513"
      Splits(0)._ColumnProps(128)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(129)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(25)._ColStyle=516"
      Splits(0)._ColumnProps(133)=   "Column(25).Order=26"
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
      Caption         =   "LIST OF SALARY"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=126,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=123,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=124,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=125,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=130,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=127,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=128,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=129,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=134,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=131,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=132,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=133,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=3"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=78,.parent=13,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=75,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=76,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=77,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=82,.parent=13,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=79,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=80,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=81,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=83,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=84,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=85,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=90,.parent=13,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=87,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=88,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=89,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=94,.parent=13,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=91,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=92,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=93,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=98,.parent=13,.alignment=2"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=95,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=96,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=97,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=102,.parent=13,.alignment=1"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=99,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=100,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=101,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=106,.parent=13,.alignment=1"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=103,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=104,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=105,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=110,.parent=13,.alignment=1"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=107,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=108,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=109,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=114,.parent=13,.alignment=1"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=111,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=112,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=113,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=118,.parent=13,.alignment=1"
      _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=115,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=116,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=117,.parent=17"
      _StyleDefs(130) =   "Splits(0).Columns(24).Style:id=122,.parent=13,.alignment=2"
      _StyleDefs(131) =   "Splits(0).Columns(24).HeadingStyle:id=119,.parent=14"
      _StyleDefs(132) =   "Splits(0).Columns(24).FooterStyle:id=120,.parent=15"
      _StyleDefs(133) =   "Splits(0).Columns(24).EditorStyle:id=121,.parent=17"
      _StyleDefs(134) =   "Splits(0).Columns(25).Style:id=138,.parent=13"
      _StyleDefs(135) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=14"
      _StyleDefs(136) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=15"
      _StyleDefs(137) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=17"
      _StyleDefs(138) =   "Named:id=33:Normal"
      _StyleDefs(139) =   ":id=33,.parent=0"
      _StyleDefs(140) =   "Named:id=34:Heading"
      _StyleDefs(141) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(142) =   ":id=34,.wraptext=-1"
      _StyleDefs(143) =   "Named:id=35:Footing"
      _StyleDefs(144) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(145) =   "Named:id=36:Selected"
      _StyleDefs(146) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(147) =   "Named:id=37:Caption"
      _StyleDefs(148) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(149) =   "Named:id=38:HighlightRow"
      _StyleDefs(150) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(151) =   "Named:id=39:EvenRow"
      _StyleDefs(152) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(153) =   "Named:id=40:OddRow"
      _StyleDefs(154) =   ":id=40,.parent=33"
      _StyleDefs(155) =   "Named:id=41:RecordSelector"
      _StyleDefs(156) =   ":id=41,.parent=34"
      _StyleDefs(157) =   "Named:id=42:FilterBar"
      _StyleDefs(158) =   ":id=42,.parent=33"
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10080
      TabIndex        =   1
      Top             =   5610
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
      MICON           =   "frm_list_salary.frx":058A
      PICN            =   "frm_list_salary.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmd_print_slip 
      Height          =   705
      Left            =   3690
      TabIndex        =   2
      Top             =   5610
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Slip"
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
      MICON           =   "frm_list_salary.frx":1638
      PICN            =   "frm_list_salary.frx":1654
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmdPrint 
      Height          =   705
      Left            =   4680
      TabIndex        =   3
      Top             =   5610
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Summary"
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
      MICON           =   "frm_list_salary.frx":26E6
      PICN            =   "frm_list_salary.frx":2702
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton cmd_print_bank 
      Height          =   705
      Left            =   5670
      TabIndex        =   4
      Top             =   5610
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Bank"
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
      MICON           =   "frm_list_salary.frx":3794
      PICN            =   "frm_list_salary.frx":37B0
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
Attribute VB_Name = "frm_list_salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsListSalary As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public tgl1 As String
Public tgl2 As String
Public str_company_code As String

Private Sub cmdPrint_Click()
    Call rpt_periode(0)
End Sub

Private Sub cmd_print_slip_Click()
    Call rpt_periode(1)
End Sub

Private Sub cmd_print_bank_Click()
    Call rpt_bank
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If rsListSalary.State Then rsListSalary.Close
    SQL = "CALL spr_list_salary('" & Format(frm_trans_salary_process.TDBGrid1.Columns("periode_from_").Value, "yyyy-MM-dd") & "'," & _
            "'" & Format(frm_trans_salary_process.TDBGrid1.Columns("periode_to_").Value, "yyyy-MM-dd") & "'," & _
            "'" & frm_trans_salary_process.TDBGrid1.Columns("month_").Value & "'," & _
            "'" & frm_trans_salary_process.TDBGrid1.Columns("company_code").Value & "')"
    rsListSalary.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsListSalary
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsListSalary.Filter = adFilterNone
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

Private Sub TDBGrid1_DblClick()
    frm_trans_salary_dtl.Show (1)
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsListSalary.Filter = getFilter()
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

Private Sub rpt_periode(ByVal j As Integer)
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim a As New frm_rpt_modal
Dim d1, d2 As String
Dim strsql As String
    
    If j = 0 Then
        str_file = "\report\rpt_summary_salary.rpt"
    ElseIf j = 1 Then
        str_file = "\report\rpt_slip_salary.rpt"
    End If
    
    d1 = Format(tgl1, "yyyy-MM-dd")
    d2 = Format(tgl2, "yyyy-MM-dd")
    
    str_sql = "call spr_salary_hardys_sum('" & d1 & "','" & d2 & "'," & _
              "0,'" & str_company_code & "','',0,'','" & LOGIN_LEVEL & "')"
    str_param_periode = "PERIODE : (" & Format(tgl2, "yyyy-MM") & ")"
    
    a.Caption = "SALARY REPORT"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
    Call a.Show(1)
End Sub

Private Sub rpt_bank()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim a As New frm_rpt_modal
Dim d1, d2 As String
    
    str_file = "\report\rpt_bank.rpt"
    
    d1 = Format(tgl1, "yyyy-MM-dd")
    d2 = Format(tgl2, "yyyy-MM-dd")
    
    str_sql = "call spr_salary_hardys_bank('" & d1 & "','" & d2 & "'," & _
                "0,'" & str_company_code & "','',0,'','MDR',0,'" & LOGIN_LEVEL & "')"
    
    a.Caption = "SALARY BANK REPORT"
    Call a.rpt_view(str_sql, str_file, str_param_periode)
    Call a.Show(1)
End Sub

