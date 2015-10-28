VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_enroll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER ENROLL (BASE IP)"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   14715
   Icon            =   "frm_mst_enroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   930
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   8250
      Width           =   14175
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   2040
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   600
      End
      Begin prj_tpc.vbButton cmd_set 
         Height          =   705
         Left            =   7050
         TabIndex        =   7
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Set"
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
         MICON           =   "frm_mst_enroll.frx":058A
         PICN            =   "frm_mst_enroll.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_unset 
         Height          =   705
         Left            =   8070
         TabIndex        =   8
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Unset"
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
         MICON           =   "frm_mst_enroll.frx":1638
         PICN            =   "frm_mst_enroll.frx":1654
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_download 
         Height          =   705
         Left            =   9090
         TabIndex        =   9
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Download"
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
         MICON           =   "frm_mst_enroll.frx":26E6
         PICN            =   "frm_mst_enroll.frx":2702
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
         Left            =   10950
         TabIndex        =   10
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
         MICON           =   "frm_mst_enroll.frx":3794
         PICN            =   "frm_mst_enroll.frx":37B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_exit 
         Height          =   705
         Left            =   12690
         TabIndex        =   11
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
         MICON           =   "frm_mst_enroll.frx":4842
         PICN            =   "frm_mst_enroll.frx":485E
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
      Height          =   6735
      Left            =   8880
      TabIndex        =   0
      Top             =   1410
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IP ADDRESS"
      Columns(0).DataField=   "ip_address"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EN. ID"
      Columns(1).DataField=   "enrollnumber"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NICK NAME"
      Columns(2).DataField=   "employee_nick_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "PASSWORD"
      Columns(3).DataField=   "password"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=64"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1138"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1058"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8708"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=64"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2805"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2725"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(3)._MinWidth=75697772"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST OF ENROLL"
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_mst_enroll.frx":58F0
      TabIndex        =   1
      Top             =   930
      Width           =   1695
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   345
      Left            =   14280
      OleObjectBlob   =   "frm_mst_enroll.frx":7856
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   405
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   6735
      Left            =   210
      TabIndex        =   12
      Top             =   1410
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "COMP. CODE"
      Columns(0).DataField=   "company_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "COMP. NAME"
      Columns(1).DataField=   "company_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DEPT. CODE"
      Columns(2).DataField=   "department_code"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DEPT. NAME"
      Columns(3).DataField=   "department_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DIV. CODE"
      Columns(4).DataField=   "division_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "SECTION"
      Columns(5).DataField=   "division_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "EMP. CODE"
      Columns(6).DataField=   "nik"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "EMPLOYEE CODE"
      Columns(7).DataField=   "employee_code"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "EMP. NAME"
      Columns(8).DataField=   "employee_name"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "NICK NAME"
      Columns(9).DataField=   "employee_nick_name"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   4
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "ACTIVE"
      Columns(10).DataField=   "flag_active"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "EMAIL"
      Columns(11).DataField=   "email"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "BIRTH DATE"
      Columns(12).DataField=   "date_birth"
      Columns(12).NumberFormat=   "dd-MM-yyyy"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "PLACE OF BIRTH"
      Columns(13).DataField=   "place_birth"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   16
      Columns(14)._MaxComboItems=   5
      Columns(14).ValueItems(0)._DefaultItem=   0
      Columns(14).ValueItems(0).Value=   "0"
      Columns(14).ValueItems(0).Value.vt=   8
      Columns(14).ValueItems(0).DisplayValue=   "Female"
      Columns(14).ValueItems(0).DisplayValue.vt=   8
      Columns(14).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(14).ValueItems(1)._DefaultItem=   0
      Columns(14).ValueItems(1).Value=   "1"
      Columns(14).ValueItems(1).Value.vt=   8
      Columns(14).ValueItems(1).DisplayValue=   "Male"
      Columns(14).ValueItems(1).DisplayValue.vt=   8
      Columns(14).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(14).ValueItems.Count=   2
      Columns(14).Caption=   "SEX"
      Columns(14).DataField=   "sex"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   16
      Columns(15)._MaxComboItems=   5
      Columns(15).ValueItems(0)._DefaultItem=   0
      Columns(15).ValueItems(0).Value=   "0"
      Columns(15).ValueItems(0).Value.vt=   8
      Columns(15).ValueItems(0).DisplayValue=   "Single"
      Columns(15).ValueItems(0).DisplayValue.vt=   8
      Columns(15).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(15).ValueItems(1)._DefaultItem=   0
      Columns(15).ValueItems(1).Value=   "1"
      Columns(15).ValueItems(1).Value.vt=   8
      Columns(15).ValueItems(1).DisplayValue=   "Married"
      Columns(15).ValueItems(1).DisplayValue.vt=   8
      Columns(15).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(15).ValueItems(2)._DefaultItem=   0
      Columns(15).ValueItems(2).Value=   "2"
      Columns(15).ValueItems(2).Value.vt=   8
      Columns(15).ValueItems(2).DisplayValue=   "Widow"
      Columns(15).ValueItems(2).DisplayValue.vt=   8
      Columns(15).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(15).ValueItems(3)._DefaultItem=   0
      Columns(15).ValueItems(3).Value=   "3"
      Columns(15).ValueItems(3).Value.vt=   8
      Columns(15).ValueItems(3).DisplayValue=   "Widower"
      Columns(15).ValueItems(3).DisplayValue.vt=   8
      Columns(15).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(15).ValueItems.Count=   4
      Columns(15).Caption=   "STATUS"
      Columns(15).DataField=   "marital_status"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "ADDRESS"
      Columns(16).DataField=   "emp_address"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "PHONE NUMBER"
      Columns(17).DataField=   "phone_number"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "BANK ACCOUNT"
      Columns(18).DataField=   "bank_account"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "START WORKING"
      Columns(19).DataField=   "start_working"
      Columns(19).NumberFormat=   "dd-MM-yyyy"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "TITLE CODE"
      Columns(20).DataField=   "title_code"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "TITLE NAME"
      Columns(21).DataField=   "title_name"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "END WORKING"
      Columns(22).DataField=   "end_working"
      Columns(22).NumberFormat=   "dd-MM-yyyy"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "REASON"
      Columns(23).DataField=   "reason"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Picture"
      Columns(24).DataField=   "picture"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   25
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   4004.788
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=25"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=3916"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3836"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2064"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1984"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=3545"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=3466"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(36)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(38)=   "Column(5).Width=2514"
      Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2434"
      Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=8708"
      Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(43)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(47)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(49)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=1508"
      Splits(0)._ColumnProps(52)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(54)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(55)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(57)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(58)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(8)._WidthInPix=1508"
      Splits(0)._ColumnProps(60)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(61)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(62)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(63)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(64)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(65)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(69)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(70)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(71)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(72)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(73)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(74)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(76)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(77)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(78)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(79)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(81)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(84)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(85)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(86)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(87)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(88)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(89)=   "Column(12).Width=2064"
      Splits(0)._ColumnProps(90)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(12)._WidthInPix=1984"
      Splits(0)._ColumnProps(92)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(93)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(94)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(95)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(96)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(97)=   "Column(13).Width=3016"
      Splits(0)._ColumnProps(98)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(13)._WidthInPix=2937"
      Splits(0)._ColumnProps(100)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(101)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(102)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(103)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(104)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(105)=   "Column(14).Width=2037"
      Splits(0)._ColumnProps(106)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(14)._WidthInPix=1958"
      Splits(0)._ColumnProps(108)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(109)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(110)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(111)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(112)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(113)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(117)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(118)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(119)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(120)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(121)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(122)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(124)=   "Column(16).AllowSizing=0"
      Splits(0)._ColumnProps(125)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(126)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(127)=   "Column(16).AllowFocus=0"
      Splits(0)._ColumnProps(128)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(129)=   "Column(16)._MinWidth=10"
      Splits(0)._ColumnProps(130)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(131)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(132)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(133)=   "Column(17).AllowSizing=0"
      Splits(0)._ColumnProps(134)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(135)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(136)=   "Column(17).AllowFocus=0"
      Splits(0)._ColumnProps(137)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(138)=   "Column(17)._MinWidth=54215968"
      Splits(0)._ColumnProps(139)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(140)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(141)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(142)=   "Column(18).AllowSizing=0"
      Splits(0)._ColumnProps(143)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(144)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(145)=   "Column(18).AllowFocus=0"
      Splits(0)._ColumnProps(146)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(147)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(148)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(149)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(150)=   "Column(19).AllowSizing=0"
      Splits(0)._ColumnProps(151)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(152)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(153)=   "Column(19).AllowFocus=0"
      Splits(0)._ColumnProps(154)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(155)=   "Column(19)._MinWidth=60129312"
      Splits(0)._ColumnProps(156)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(157)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(158)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(159)=   "Column(20).AllowSizing=0"
      Splits(0)._ColumnProps(160)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(161)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(162)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(163)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(164)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(165)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(166)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(167)=   "Column(21).AllowSizing=0"
      Splits(0)._ColumnProps(168)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(169)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(170)=   "Column(21).AllowFocus=0"
      Splits(0)._ColumnProps(171)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(172)=   "Column(21)._MinWidth=79702332"
      Splits(0)._ColumnProps(173)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(174)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(176)=   "Column(22).AllowSizing=0"
      Splits(0)._ColumnProps(177)=   "Column(22)._ColStyle=516"
      Splits(0)._ColumnProps(178)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(179)=   "Column(22).AllowFocus=0"
      Splits(0)._ColumnProps(180)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(181)=   "Column(22)._MinWidth=79914544"
      Splits(0)._ColumnProps(182)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(183)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(184)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(185)=   "Column(23).AllowSizing=0"
      Splits(0)._ColumnProps(186)=   "Column(23)._ColStyle=516"
      Splits(0)._ColumnProps(187)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(188)=   "Column(23).AllowFocus=0"
      Splits(0)._ColumnProps(189)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(190)=   "Column(23)._MinWidth=79789632"
      Splits(0)._ColumnProps(191)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(192)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(193)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(194)=   "Column(24)._ColStyle=516"
      Splits(0)._ColumnProps(195)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(196)=   "Column(24).Order=25"
      Splits(1)._UserFlags=   0
      Splits(1).Size  =   2
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).ScrollBars=   3
      Splits(1).DividerColor=   13160660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=25"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=1826"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1746"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=1720"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=1640"
      Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(25)=   "Column(3).Width=1720"
      Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=1640"
      Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=1508"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=1429"
      Splits(1)._ColumnProps(36)=   "Column(4).AllowSizing=0"
      Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=516"
      Splits(1)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(1)._ColumnProps(39)=   "Column(4).AllowFocus=0"
      Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(41)=   "Column(4)._MinWidth=80002672"
      Splits(1)._ColumnProps(42)=   "Column(5).Width=1508"
      Splits(1)._ColumnProps(43)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(44)=   "Column(5)._WidthInPix=1429"
      Splits(1)._ColumnProps(45)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(46)=   "Column(5)._ColStyle=516"
      Splits(1)._ColumnProps(47)=   "Column(5).Visible=0"
      Splits(1)._ColumnProps(48)=   "Column(5).AllowFocus=0"
      Splits(1)._ColumnProps(49)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(50)=   "Column(5)._MinWidth=80001968"
      Splits(1)._ColumnProps(51)=   "Column(6).Width=2170"
      Splits(1)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(6)._WidthInPix=2090"
      Splits(1)._ColumnProps(54)=   "Column(6)._ColStyle=516"
      Splits(1)._ColumnProps(55)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(56)=   "Column(6)._MinWidth=80000960"
      Splits(1)._ColumnProps(57)=   "Column(7).Width=1879"
      Splits(1)._ColumnProps(58)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(7)._WidthInPix=1799"
      Splits(1)._ColumnProps(60)=   "Column(7)._ColStyle=8708"
      Splits(1)._ColumnProps(61)=   "Column(7).Visible=0"
      Splits(1)._ColumnProps(62)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(63)=   "Column(7)._MinWidth=80000960"
      Splits(1)._ColumnProps(64)=   "Column(8).Width=3122"
      Splits(1)._ColumnProps(65)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(66)=   "Column(8)._WidthInPix=3043"
      Splits(1)._ColumnProps(67)=   "Column(8)._ColStyle=8708"
      Splits(1)._ColumnProps(68)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(69)=   "Column(8)._MinWidth=79999936"
      Splits(1)._ColumnProps(70)=   "Column(9).Width=2170"
      Splits(1)._ColumnProps(71)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(72)=   "Column(9)._WidthInPix=2090"
      Splits(1)._ColumnProps(73)=   "Column(9)._ColStyle=8708"
      Splits(1)._ColumnProps(74)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(75)=   "Column(9)._MinWidth=80007280"
      Splits(1)._ColumnProps(76)=   "Column(10).Width=1191"
      Splits(1)._ColumnProps(77)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(78)=   "Column(10)._WidthInPix=1111"
      Splits(1)._ColumnProps(79)=   "Column(10)._ColStyle=513"
      Splits(1)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(81)=   "Column(10)._MinWidth=80007280"
      Splits(1)._ColumnProps(82)=   "Column(11).Width=4233"
      Splits(1)._ColumnProps(83)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(84)=   "Column(11)._WidthInPix=4154"
      Splits(1)._ColumnProps(85)=   "Column(11)._ColStyle=516"
      Splits(1)._ColumnProps(86)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(87)=   "Column(11)._MinWidth=80007280"
      Splits(1)._ColumnProps(88)=   "Column(12).Width=2064"
      Splits(1)._ColumnProps(89)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(90)=   "Column(12)._WidthInPix=1984"
      Splits(1)._ColumnProps(91)=   "Column(12)._ColStyle=8705"
      Splits(1)._ColumnProps(92)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(93)=   "Column(12)._MinWidth=80007280"
      Splits(1)._ColumnProps(94)=   "Column(13).Width=3016"
      Splits(1)._ColumnProps(95)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(96)=   "Column(13)._WidthInPix=2937"
      Splits(1)._ColumnProps(97)=   "Column(13)._ColStyle=8708"
      Splits(1)._ColumnProps(98)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(99)=   "Column(13)._MinWidth=80010048"
      Splits(1)._ColumnProps(100)=   "Column(14).Width=2037"
      Splits(1)._ColumnProps(101)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(102)=   "Column(14)._WidthInPix=1958"
      Splits(1)._ColumnProps(103)=   "Column(14)._ColStyle=8705"
      Splits(1)._ColumnProps(104)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(105)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(106)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(107)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(108)=   "Column(15)._ColStyle=8705"
      Splits(1)._ColumnProps(109)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(110)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(111)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(112)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(113)=   "Column(16)._ColStyle=8708"
      Splits(1)._ColumnProps(114)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(115)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(116)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(117)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(118)=   "Column(17)._ColStyle=8708"
      Splits(1)._ColumnProps(119)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(120)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(121)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(122)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(123)=   "Column(18)._ColStyle=8708"
      Splits(1)._ColumnProps(124)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(125)=   "Column(19).Width=2725"
      Splits(1)._ColumnProps(126)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(127)=   "Column(19)._WidthInPix=2646"
      Splits(1)._ColumnProps(128)=   "Column(19)._ColStyle=8705"
      Splits(1)._ColumnProps(129)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(130)=   "Column(20).Width=2725"
      Splits(1)._ColumnProps(131)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(132)=   "Column(20)._WidthInPix=2646"
      Splits(1)._ColumnProps(133)=   "Column(20)._ColStyle=8708"
      Splits(1)._ColumnProps(134)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(135)=   "Column(21).Width=2725"
      Splits(1)._ColumnProps(136)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(137)=   "Column(21)._WidthInPix=2646"
      Splits(1)._ColumnProps(138)=   "Column(21)._ColStyle=8708"
      Splits(1)._ColumnProps(139)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(140)=   "Column(22).Width=2725"
      Splits(1)._ColumnProps(141)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(142)=   "Column(22)._WidthInPix=2646"
      Splits(1)._ColumnProps(143)=   "Column(22)._ColStyle=8705"
      Splits(1)._ColumnProps(144)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(145)=   "Column(23).Width=2725"
      Splits(1)._ColumnProps(146)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(147)=   "Column(23)._WidthInPix=2646"
      Splits(1)._ColumnProps(148)=   "Column(23)._ColStyle=8708"
      Splits(1)._ColumnProps(149)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(150)=   "Column(23)._MinWidth=80015760"
      Splits(1)._ColumnProps(151)=   "Column(24).Width=2725"
      Splits(1)._ColumnProps(152)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(153)=   "Column(24)._WidthInPix=2646"
      Splits(1)._ColumnProps(154)=   "Column(24)._ColStyle=516"
      Splits(1)._ColumnProps(155)=   "Column(24).Visible=0"
      Splits(1)._ColumnProps(156)=   "Column(24).Order=25"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST OF EMPLOYEE"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=222,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=219,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=220,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=221,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=98,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=234,.parent=13,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=231,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=232,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=233,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=242,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=239,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=240,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=241,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=50,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=47,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=48,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=49,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=54,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=62,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=59,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=60,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=61,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=66,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=63,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=64,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=65,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=102,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=110,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=107,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=108,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=109,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=46,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=43,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=44,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=45,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=58,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=55,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=56,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=57,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=94,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=91,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=92,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=93,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=106,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=103,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=104,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=105,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=118,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=115,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=116,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=117,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=122,.parent=13"
      _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=119,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=120,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=121,.parent=17"
      _StyleDefs(130) =   "Splits(0).Columns(24).Style:id=114,.parent=13"
      _StyleDefs(131) =   "Splits(0).Columns(24).HeadingStyle:id=111,.parent=14"
      _StyleDefs(132) =   "Splits(0).Columns(24).FooterStyle:id=112,.parent=15"
      _StyleDefs(133) =   "Splits(0).Columns(24).EditorStyle:id=113,.parent=17"
      _StyleDefs(134) =   "Splits(1).Style:id=123,.parent=1"
      _StyleDefs(135) =   "Splits(1).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(136) =   ":id=132,.fgcolor=&H80000009&"
      _StyleDefs(137) =   "Splits(1).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(138) =   ":id=124,.fgcolor=&H80000002&"
      _StyleDefs(139) =   "Splits(1).FooterStyle:id=125,.parent=3"
      _StyleDefs(140) =   "Splits(1).InactiveStyle:id=126,.parent=5"
      _StyleDefs(141) =   "Splits(1).SelectedStyle:id=128,.parent=6"
      _StyleDefs(142) =   "Splits(1).EditorStyle:id=127,.parent=7"
      _StyleDefs(143) =   "Splits(1).HighlightRowStyle:id=129,.parent=8"
      _StyleDefs(144) =   "Splits(1).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(145) =   "Splits(1).OddRowStyle:id=131,.parent=10"
      _StyleDefs(146) =   "Splits(1).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(147) =   "Splits(1).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(148) =   "Splits(1).Columns(0).Style:id=138,.parent=123"
      _StyleDefs(149) =   "Splits(1).Columns(0).HeadingStyle:id=135,.parent=124"
      _StyleDefs(150) =   "Splits(1).Columns(0).FooterStyle:id=136,.parent=125"
      _StyleDefs(151) =   "Splits(1).Columns(0).EditorStyle:id=137,.parent=127"
      _StyleDefs(152) =   "Splits(1).Columns(1).Style:id=142,.parent=123"
      _StyleDefs(153) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=124"
      _StyleDefs(154) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=125"
      _StyleDefs(155) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=127"
      _StyleDefs(156) =   "Splits(1).Columns(2).Style:id=146,.parent=123"
      _StyleDefs(157) =   "Splits(1).Columns(2).HeadingStyle:id=143,.parent=124"
      _StyleDefs(158) =   "Splits(1).Columns(2).FooterStyle:id=144,.parent=125"
      _StyleDefs(159) =   "Splits(1).Columns(2).EditorStyle:id=145,.parent=127"
      _StyleDefs(160) =   "Splits(1).Columns(3).Style:id=150,.parent=123"
      _StyleDefs(161) =   "Splits(1).Columns(3).HeadingStyle:id=147,.parent=124"
      _StyleDefs(162) =   "Splits(1).Columns(3).FooterStyle:id=148,.parent=125"
      _StyleDefs(163) =   "Splits(1).Columns(3).EditorStyle:id=149,.parent=127"
      _StyleDefs(164) =   "Splits(1).Columns(4).Style:id=154,.parent=123"
      _StyleDefs(165) =   "Splits(1).Columns(4).HeadingStyle:id=151,.parent=124"
      _StyleDefs(166) =   "Splits(1).Columns(4).FooterStyle:id=152,.parent=125"
      _StyleDefs(167) =   "Splits(1).Columns(4).EditorStyle:id=153,.parent=127"
      _StyleDefs(168) =   "Splits(1).Columns(5).Style:id=158,.parent=123"
      _StyleDefs(169) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=124"
      _StyleDefs(170) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=125"
      _StyleDefs(171) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=127"
      _StyleDefs(172) =   "Splits(1).Columns(6).Style:id=226,.parent=123"
      _StyleDefs(173) =   "Splits(1).Columns(6).HeadingStyle:id=223,.parent=124"
      _StyleDefs(174) =   "Splits(1).Columns(6).FooterStyle:id=224,.parent=125"
      _StyleDefs(175) =   "Splits(1).Columns(6).EditorStyle:id=225,.parent=127"
      _StyleDefs(176) =   "Splits(1).Columns(7).Style:id=162,.parent=123,.locked=-1"
      _StyleDefs(177) =   "Splits(1).Columns(7).HeadingStyle:id=159,.parent=124"
      _StyleDefs(178) =   "Splits(1).Columns(7).FooterStyle:id=160,.parent=125"
      _StyleDefs(179) =   "Splits(1).Columns(7).EditorStyle:id=161,.parent=127"
      _StyleDefs(180) =   "Splits(1).Columns(8).Style:id=166,.parent=123,.locked=-1"
      _StyleDefs(181) =   "Splits(1).Columns(8).HeadingStyle:id=163,.parent=124"
      _StyleDefs(182) =   "Splits(1).Columns(8).FooterStyle:id=164,.parent=125"
      _StyleDefs(183) =   "Splits(1).Columns(8).EditorStyle:id=165,.parent=127"
      _StyleDefs(184) =   "Splits(1).Columns(9).Style:id=230,.parent=123,.locked=-1"
      _StyleDefs(185) =   "Splits(1).Columns(9).HeadingStyle:id=227,.parent=124"
      _StyleDefs(186) =   "Splits(1).Columns(9).FooterStyle:id=228,.parent=125"
      _StyleDefs(187) =   "Splits(1).Columns(9).EditorStyle:id=229,.parent=127"
      _StyleDefs(188) =   "Splits(1).Columns(10).Style:id=238,.parent=123,.alignment=2"
      _StyleDefs(189) =   "Splits(1).Columns(10).HeadingStyle:id=235,.parent=124"
      _StyleDefs(190) =   "Splits(1).Columns(10).FooterStyle:id=236,.parent=125"
      _StyleDefs(191) =   "Splits(1).Columns(10).EditorStyle:id=237,.parent=127"
      _StyleDefs(192) =   "Splits(1).Columns(11).Style:id=246,.parent=123"
      _StyleDefs(193) =   "Splits(1).Columns(11).HeadingStyle:id=243,.parent=124"
      _StyleDefs(194) =   "Splits(1).Columns(11).FooterStyle:id=244,.parent=125"
      _StyleDefs(195) =   "Splits(1).Columns(11).EditorStyle:id=245,.parent=127"
      _StyleDefs(196) =   "Splits(1).Columns(12).Style:id=170,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(197) =   "Splits(1).Columns(12).HeadingStyle:id=167,.parent=124"
      _StyleDefs(198) =   "Splits(1).Columns(12).FooterStyle:id=168,.parent=125"
      _StyleDefs(199) =   "Splits(1).Columns(12).EditorStyle:id=169,.parent=127"
      _StyleDefs(200) =   "Splits(1).Columns(13).Style:id=174,.parent=123,.locked=-1"
      _StyleDefs(201) =   "Splits(1).Columns(13).HeadingStyle:id=171,.parent=124"
      _StyleDefs(202) =   "Splits(1).Columns(13).FooterStyle:id=172,.parent=125"
      _StyleDefs(203) =   "Splits(1).Columns(13).EditorStyle:id=173,.parent=127"
      _StyleDefs(204) =   "Splits(1).Columns(14).Style:id=178,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(205) =   "Splits(1).Columns(14).HeadingStyle:id=175,.parent=124"
      _StyleDefs(206) =   "Splits(1).Columns(14).FooterStyle:id=176,.parent=125"
      _StyleDefs(207) =   "Splits(1).Columns(14).EditorStyle:id=177,.parent=127"
      _StyleDefs(208) =   "Splits(1).Columns(15).Style:id=182,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(209) =   "Splits(1).Columns(15).HeadingStyle:id=179,.parent=124"
      _StyleDefs(210) =   "Splits(1).Columns(15).FooterStyle:id=180,.parent=125"
      _StyleDefs(211) =   "Splits(1).Columns(15).EditorStyle:id=181,.parent=127"
      _StyleDefs(212) =   "Splits(1).Columns(16).Style:id=186,.parent=123,.locked=-1"
      _StyleDefs(213) =   "Splits(1).Columns(16).HeadingStyle:id=183,.parent=124"
      _StyleDefs(214) =   "Splits(1).Columns(16).FooterStyle:id=184,.parent=125"
      _StyleDefs(215) =   "Splits(1).Columns(16).EditorStyle:id=185,.parent=127"
      _StyleDefs(216) =   "Splits(1).Columns(17).Style:id=190,.parent=123,.locked=-1"
      _StyleDefs(217) =   "Splits(1).Columns(17).HeadingStyle:id=187,.parent=124"
      _StyleDefs(218) =   "Splits(1).Columns(17).FooterStyle:id=188,.parent=125"
      _StyleDefs(219) =   "Splits(1).Columns(17).EditorStyle:id=189,.parent=127"
      _StyleDefs(220) =   "Splits(1).Columns(18).Style:id=194,.parent=123,.locked=-1"
      _StyleDefs(221) =   "Splits(1).Columns(18).HeadingStyle:id=191,.parent=124"
      _StyleDefs(222) =   "Splits(1).Columns(18).FooterStyle:id=192,.parent=125"
      _StyleDefs(223) =   "Splits(1).Columns(18).EditorStyle:id=193,.parent=127"
      _StyleDefs(224) =   "Splits(1).Columns(19).Style:id=198,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(225) =   "Splits(1).Columns(19).HeadingStyle:id=195,.parent=124"
      _StyleDefs(226) =   "Splits(1).Columns(19).FooterStyle:id=196,.parent=125"
      _StyleDefs(227) =   "Splits(1).Columns(19).EditorStyle:id=197,.parent=127"
      _StyleDefs(228) =   "Splits(1).Columns(20).Style:id=202,.parent=123,.locked=-1"
      _StyleDefs(229) =   "Splits(1).Columns(20).HeadingStyle:id=199,.parent=124"
      _StyleDefs(230) =   "Splits(1).Columns(20).FooterStyle:id=200,.parent=125"
      _StyleDefs(231) =   "Splits(1).Columns(20).EditorStyle:id=201,.parent=127"
      _StyleDefs(232) =   "Splits(1).Columns(21).Style:id=206,.parent=123,.locked=-1"
      _StyleDefs(233) =   "Splits(1).Columns(21).HeadingStyle:id=203,.parent=124"
      _StyleDefs(234) =   "Splits(1).Columns(21).FooterStyle:id=204,.parent=125"
      _StyleDefs(235) =   "Splits(1).Columns(21).EditorStyle:id=205,.parent=127"
      _StyleDefs(236) =   "Splits(1).Columns(22).Style:id=214,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(237) =   "Splits(1).Columns(22).HeadingStyle:id=211,.parent=124"
      _StyleDefs(238) =   "Splits(1).Columns(22).FooterStyle:id=212,.parent=125"
      _StyleDefs(239) =   "Splits(1).Columns(22).EditorStyle:id=213,.parent=127"
      _StyleDefs(240) =   "Splits(1).Columns(23).Style:id=218,.parent=123,.locked=-1"
      _StyleDefs(241) =   "Splits(1).Columns(23).HeadingStyle:id=215,.parent=124"
      _StyleDefs(242) =   "Splits(1).Columns(23).FooterStyle:id=216,.parent=125"
      _StyleDefs(243) =   "Splits(1).Columns(23).EditorStyle:id=217,.parent=127"
      _StyleDefs(244) =   "Splits(1).Columns(24).Style:id=210,.parent=123"
      _StyleDefs(245) =   "Splits(1).Columns(24).HeadingStyle:id=207,.parent=124"
      _StyleDefs(246) =   "Splits(1).Columns(24).FooterStyle:id=208,.parent=125"
      _StyleDefs(247) =   "Splits(1).Columns(24).EditorStyle:id=209,.parent=127"
      _StyleDefs(248) =   "Named:id=33:Normal"
      _StyleDefs(249) =   ":id=33,.parent=0"
      _StyleDefs(250) =   "Named:id=34:Heading"
      _StyleDefs(251) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(252) =   ":id=34,.wraptext=-1"
      _StyleDefs(253) =   "Named:id=35:Footing"
      _StyleDefs(254) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(255) =   "Named:id=36:Selected"
      _StyleDefs(256) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(257) =   "Named:id=37:Caption"
      _StyleDefs(258) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(259) =   "Named:id=38:HighlightRow"
      _StyleDefs(260) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(261) =   "Named:id=39:EvenRow"
      _StyleDefs(262) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(263) =   "Named:id=40:OddRow"
      _StyleDefs(264) =   ":id=40,.parent=33"
      _StyleDefs(265) =   "Named:id=41:RecordSelector"
      _StyleDefs(266) =   ":id=41,.parent=34"
      _StyleDefs(267) =   "Named:id=42:FilterBar"
      _StyleDefs(268) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER ENROLL"
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
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   3555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   930
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_enroll.frx":787A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14730
   End
End
Attribute VB_Name = "frm_mst_enroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCompany As New ADODB.Recordset
Dim rsEmployee As New ADODB.Recordset
Dim rsEnroll As New ADODB.Recordset

Dim cn_fp As Boolean
Dim glngEnrollData(DataSize) As Long
Dim glngEnrollPData As Integer
Dim gbytEnrollData(DataSize * 5) As Byte
Dim vEMachineNumber As Integer
Public start_rec, end_rec, running_rec, db_rec As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim oClause1 As String

Private Function set_bookmark_ado(ByVal lng_no As Long) As Boolean
    rsEnroll.MoveFirst
    
    rsEnroll.Find ("enrollnumber=" & lng_no)   ', 0, adSearchForward, 1)
    If Not (rsEnroll.EOF = True Or rsEnroll.BOF = True) Then
        set_bookmark_ado = True
    Else
        set_bookmark_ado = False
    End If
End Function

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_download_Click()
    frm_lookup_mst_device.public_int_mode = 0
    frm_lookup_mst_device.Show 1
End Sub

Public Sub cmd_refresh_Click()
    timer1.Enabled = True
End Sub

'Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Adodc1.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
'End Sub

Private Function delete_enroll_typeA(ByVal str_e As String) As Boolean
Dim eMachineNumber As Integer
Dim bln_e As Boolean

    eMachineNumber = 1
    If str_e = "" Then Exit Function
    bln_e = CZKEM1.DeleteEnrollData(CInt(eMachineNumber), CLng(str_e), CInt(eMachineNumber), 12)
    If bln_e Then
        delete_enroll_typeA = True
    Else
        delete_enroll_typeA = False
    End If
End Function

Private Function delete_enroll_typeB(ByVal str_e As String) As Boolean
Dim eMachineNumber As Integer
Dim bln_e As Boolean

    eMachineNumber = 1
    If str_e = "" Then Exit Function
    bln_e = CZKEM1.SSR_DeleteEnrollData(CInt(eMachineNumber), CLng(str_e), 12)
    If bln_e Then
        delete_enroll_typeB = True
    Else
        delete_enroll_typeB = False
    End If
End Function

Private Function check_multi_ip() As Boolean
Dim i As Integer
Dim item
Dim str_buf1, str_buf2 As String

    check_multi_ip = False
    
    'If Not TDBGrid1.ApproxCount > 0 And Not TDBGrid2.ApproxCount > 0 Then
    '    MsgBox "No data selected...", vbInformation, headerMSG
    '    Exit Sub
    'End If
    Set SelBks = TDBGrid1.SelBookmarks
    
    str_buf1 = TDBGrid1.Columns("ip_address").Value
    With TDBGrid1
        For Each item In SelBks
            
            str_buf2 = TDBGrid1.Columns("ip_address").CellText(item)
            If Not str_buf1 = str_buf2 Then
                check_multi_ip = True
                Exit Function
            End If
        Next
    End With
End Function

Private Sub cmd_set_Click()
Dim rs1 As New ADODB.Recordset
Dim i As Integer
Dim item
Dim ret As Boolean

On Error GoTo Err

    If Not TDBGrid1.ApproxCount > 0 And Not TDBGrid2.ApproxCount > 0 Then
        MsgBox "No data selected...", vbInformation, headerMSG
        Exit Sub
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    
    If check_multi_ip = True Then
        MsgBox "There are more than one device detected!", vbInformation, headerMSG
        Exit Sub
    End If
    FG_IP_ADDRESS = rsEnroll.Fields("ip_address").Value
    FG_PORT_NUMBER = rsEnroll.Fields("port_number").Value

    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from m_device where ip_address = '" & FG_IP_ADDRESS & "'", CnG, adOpenStatic, adLockReadOnly
    FG_DEVICE_TYPE = rs1.Fields("flag_type").Value

    If Not connect Then
        MsgBox "Error connecting to source device...", vbCritical, headerMSG
        Exit Sub
    End If
    
    CnG.BeginTrans
    With TDBGrid2
        For Each item In SelBks
            
            If FG_DEVICE_TYPE = 0 Then
                ret = set_enroll_name_typeA(TDBGrid1.Columns("enrollnumber").CellText(item), _
                        .Columns("employee_nick_name").Value, _
                        TDBGrid1.Columns("password").Value)
            Else
                ret = set_enroll_name_typeB(TDBGrid1.Columns("enrollnumber").CellText(item), _
                        .Columns("employee_nick_name").Value, _
                        TDBGrid1.Columns("password").Value)
            End If
            
            If ret Then
                CnG.Execute "call spi_m_enroll_link ('" & FG_IP_ADDRESS & "', " _
                    & TDBGrid1.Columns("enrollnumber").CellText(item) & ",'" _
                    & .Columns("employee_code").Value & "','" _
                    & Replace(.Columns("employee_name").Value, "'", "\'") & "','" _
                    & Replace(.Columns("employee_nick_name").Value, "'", "\'") & "','" _
                    & .Columns("company_code").Value & "','" _
                    & Replace(.Columns("company_name").Value, "'", "''") & "')"
            
            End If
        Next
    End With
    CnG.CommitTrans
    Call load_data_enroll
    Call disconnect
    
    Exit Sub

Err:
MsgBox "Set Enroll Failed...", vbInformation, headerMSG
CnG.RollbackTrans
Call disconnect
End Sub

Private Function set_enroll_name_typeA _
(ByVal en As Integer, ByVal str As String, ByVal Pwd As String) As Boolean
Dim iEnrollNumber
Dim vMachineNumber As Integer
Dim bEnrollData(1024) As Byte
Dim str_pass

    str_pass = Trim(Pwd)
    vMachineNumber = 1
    iEnrollNumber = CLng(en)
    If CZKEM1.SetUserInfo(vMachineNumber, iEnrollNumber, CStr(str), str_pass, 0, True) Then
        set_enroll_name_typeA = True
    Else
        set_enroll_name_typeA = False
    End If
End Function

Private Function set_enroll_name_typeB _
(ByVal en As Long, ByVal str As String, ByVal Pwd As String) As Boolean

'-------------
Dim dwEnrollNumber 'As String 'As Long
Dim name 'As String
Dim passWord 'As String
Dim privileg As Long
Dim Enabled As Boolean
'dwEnrollNumber = 1
name = "Henry"
passWord = "12"
privileg = 3
Enabled = True

Dim iEnrollNumber As Long
Dim vMachineNumber 'As Long
Dim bEnrollData(1024) As Byte
Dim str_pass As String

name = str
passWord = Trim(Pwd)
str_pass = Trim(Pwd)
vMachineNumber = 1
iEnrollNumber = CLng(en)
dwEnrollNumber = Trim(CStr(en))

If CZKEM1.SSR_SetUserInfo(vMachineNumber, dwEnrollNumber, name, passWord, 1, True) Then
    set_enroll_name_typeB = True
Else
    set_enroll_name_typeB = False
End If
End Function

Private Sub cmd_unset_Click()
Dim rs1 As New ADODB.Recordset
Dim i As Integer
Dim item
Dim ret As Boolean

On Error GoTo Err

    If Not TDBGrid1.ApproxCount > 0 And Not TDBGrid2.ApproxCount > 0 Then
        Exit Sub
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    
    If check_multi_ip = True Then
        MsgBox "There are more than one device detected!", vbInformation, headerMSG
        Exit Sub
    End If
    FG_IP_ADDRESS = rsEnroll.Fields("ip_address").Value
    FG_PORT_NUMBER = rsEnroll.Fields("port_number").Value
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select * from m_device where ip_address = '" & FG_IP_ADDRESS & "'", CnG, adOpenStatic, adLockReadOnly
    FG_DEVICE_TYPE = rs1.Fields("flag_type").Value

    If Not connect Then
        MsgBox "Error connecting to source device...", vbCritical, headerMSG
        Exit Sub
    End If
    
    CnG.BeginTrans
    With TDBGrid2
        For Each item In SelBks
            
            If FG_DEVICE_TYPE = 0 Then
                ret = set_enroll_name_typeA(TDBGrid1.Columns("enrollnumber").CellText(item), "", "01111")
            Else
                ret = set_enroll_name_typeB(TDBGrid1.Columns("enrollnumber").CellText(item), "", "01111")
            End If
            
            If ret Then _
            CnG.Execute "delete from m_enroll_link where ip_address = '" & rsEnroll.Fields("ip_address").Value _
            & "' and enrollnumber = " & TDBGrid1.Columns("enrollnumber").CellText(item)
        Next
    End With
    CnG.CommitTrans
    Call load_data_enroll
    Call disconnect
    
    Exit Sub

Err:
MsgBox "Unset Enroll Failed...", vbInformation, headerMSG
CnG.RollbackTrans
Call disconnect
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim item
Dim ret As Integer

On Error GoTo Err

    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    
    i = MsgBox("Are you sure to delete " & SelBks.Count & " Enroll Data?", vbOKCancel, headerMSG)
    If Not i = vbOK Then Exit Sub
    
    FG_IP_ADDRESS = TDBGrid1.Columns("ip_address").Value
    FG_PORT_NUMBER = get_port(TDBGrid1.Columns("ip_address").Value)
    
    If Not connect Then
        MsgBox "Error connecting to source device...", vbCritical, headerMSG
        Exit Sub
    End If
        
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        If FG_DEVICE_TYPE = 0 Then
            ret = delete_enroll_typeA(TDBGrid1.Columns("enrollnumber").CellText(item))
        Else
            ret = delete_enroll_typeB(TDBGrid1.Columns("enrollnumber").CellText(item))
        End If
        
        If ret = True Then
            CnG.Execute "delete from m_enroll where enrollnumber = " _
                & TDBGrid1.Columns("enrollnumber").CellText(item) & " and ip_address ='" _
                & TDBGrid1.Columns("ip_address").CellText(item) & "'"
            
            CnG.Execute "delete from m_enroll_link where enrollnumber = " _
                & TDBGrid1.Columns("enrollnumber").CellText(item) & " and ip_address ='" _
                & TDBGrid1.Columns("ip_address").CellText(item) & "'"
        End If
    Next
    CnG.CommitTrans
    MsgBox i & " enroll's data Succesfully Deleted...", vbInformation, headerMSG
    Call load_data_enroll
    Call disconnect
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Function get_port(ByVal str_ip As String) As Integer
Dim rs1 As New ADODB.Recordset

    rs1.Open "select * from m_device where ip_address = '" & str_ip & "'", CnG, adOpenStatic, adLockReadOnly
    
    If rs1.RecordCount > 0 Then
        get_port = rs1!port_number
    Else
        get_port = 4370
    End If
End Function

Private Sub Form_Load()
    cn_fp = False
    start_rec = 0
    end_rec = 100
    running_rec = 0
    db_rec = 0
    oClause = ""
    
    vMachineNumber = 1
    vEMachineNumber = 1
    
    Call load_data_enroll
    Call load_data_company
    
    Call load_data_user_access(Me)
    cmd_download.Enabled = blnUser_Edit
    cmd_set.Enabled = blnUser_Edit
    cmd_unset.Enabled = blnUser_Edit
    cmdDelete.Enabled = blnUser_Delete
    
    timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_enroll = Nothing
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_employee
    End If
End Sub

Private Sub TDBGrid1_FilterChange()
Dim i As Integer

On Error GoTo Err

    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsEnroll.Filter = getFilter()
    TDBGrid1.Col = i
    TDBGrid1.EditActive = True
    
    TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
    If TDBGrid1.ApproxCount < 1 Then
        Call clear_filter1
        TDBGrid1.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter1
End Sub

Private Sub TDBGrid2_FilterChange()
Dim i As Integer

On Error GoTo Err

    Set Cols = TDBGrid2.Columns
    i = TDBGrid2.Col
    TDBGrid2.HoldFields
    
    rsEmployee.Filter = getFilter()
    TDBGrid2.Col = i
    TDBGrid2.EditActive = True
    
    TDBGrid2.SelStart = Len(TDBGrid2.Columns(i).FilterText)
    If TDBGrid2.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid2.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid2.Columns
        Col.FilterText = ""
    Next Col
    rsEmployee.Filter = adFilterNone
End Sub

Private Sub clear_filter1()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsEnroll.Filter = adFilterNone
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

Private Sub TDBGrid2_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid2.Columns(ColIndex).Caption = "BIRTH DATE" Or _
        TDBGrid2.Columns(ColIndex).Caption = "START WORKING" Or _
        TDBGrid2.Columns(ColIndex).Caption = "END WORKING" Then
            Value = Format(Value, "dd-mm-yyyy")
    End If
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_employee()
    If rsEmployee.State Then rsEmployee.Close
    SQL = "select a.*, b.division_name, c.department_name from m_employee a join m_division b on a.company_code = b.company_code and a.division_code = b.division_code " & _
            "join m_department c on a.department_code = c.department_code " & _
            "where a.company_code = '" _
            & TDBCombo_company.Columns("company_code").Value & "' " & oClause
    rsEmployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid2.DataSource = rsEmployee
End Sub

Public Sub load_data_enroll()
    If rsEnroll.State Then rsEnroll.Close
    SQL = "SELECT a.ip_address,f_get_port(a.ip_address) AS port_number," & _
            "a.EnrollNumber,a.EnrollName,b.employee_code,b.employee_name," & _
            "b.employee_nick_name,b.company_code,b.company_name,a.Password " & _
          "FROM m_enroll a LEFT JOIN m_enroll_link b " & _
            "ON a.EnrollNumber = b.enrollnumber " & _
                "AND a.ip_address = b.ip_address " & oClause1
    rsEnroll.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsEnroll
End Sub

Private Function connect() As Boolean
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, True
        CZKEM1.disconnect
    End If
    
    cn_fp = CZKEM1.Connect_Net(FG_IP_ADDRESS, CLng(FG_PORT_NUMBER))
    
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, False
        connect = True
    Else
        connect = False
        Exit Function
    End If
End Function

Private Function disconnect() As Boolean
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, True
        CZKEM1.disconnect
    End If
End Function

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    
    If Not connect Then
        MsgBox "Error connecting to source device...", vbCritical, headerMSG
        Exit Sub
    End If
    
    frm_progess.Caption = "Downloading Enroll ID..."
    frm_progess.Show 1
    
    Call disconnect
End Sub

Public Function get_all_user_info_tipeA() As Long
Dim dwEnrollNmber As Long
Dim name As String
Dim passWord As String
Dim privilege As Long
Dim Enabled As Boolean
Dim dwMachineNumber As Long
Dim blni As Boolean
Dim i, j, k, tt As Long
Dim rs1 As New ADODB.Recordset
Dim str1 As String

'-------------------------
Dim id_qty As Long, vid As Long

On Error GoTo Err

    dwMachineNumber = 1
    i = 0
    tt = 0
    
    blni = CZKEM1.ReadAllUserID(1)
    If blni Then
        CnG.BeginTrans
         
        While CZKEM1.GetAllUserInfo(dwMachineNumber, dwEnrollNmber, name, passWord, privilege, Enabled)
            DoEvents
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select COUNT(*) as recs FROM m_enroll WHERE ip_address = '" & FG_IP_ADDRESS _
                & "' AND enrollnumber=" & Val(dwEnrollNmber), CnG, adOpenStatic, adLockReadOnly
            
            If rs1.Fields("recs").Value = 0 Then
                str1 = "INSERT INTO m_enroll(ip_address, enrollnumber, " _
                & "enrollname,emachinenumber,fingernumber,privilige,password) " _
                & "VALUES('" & FG_IP_ADDRESS & "'," & dwEnrollNmber & ",'', " _
                & dwMachineNumber & ", 1," & privilege & ",'" & passWord & "')"
                
                CnG.Execute str1
            End If
            
            i = i + 1
            
            If i = 10 Then
                k = k
            ElseIf i = 30 Then
                k = k
            ElseIf i = 50 Then
                k = k
            End If
            
        Wend
        CnG.CommitTrans
    End If
    
    get_all_user_info_tipeA = i
    Exit Function
    
Err:
    CnG.RollbackTrans
    MsgBox "Failed to read device...", vbCritical, headerMSG
    Call disconnect
    get_all_user_info_tipeA = 0
End Function

Public Function get_all_user_info_tipeB() As Long
Dim dwEnrollNmber As String
Dim name As String
Dim passWord As String
Dim privilege As Long
Dim Enabled As Boolean
Dim dwMachineNumber As Long
Dim blni As Boolean
Dim i, j, k, tt As Long
Dim rs1 As New ADODB.Recordset
Dim str1 As String

'-------------------------
Dim id_qty As Long, vid As Long

On Error GoTo Err

    dwMachineNumber = 1
    i = 0
    tt = 0
    
    blni = CZKEM1.ReadAllUserID(1)
    If blni Then
        CnG.BeginTrans
         
        While CZKEM1.SSR_GetAllUserInfo(dwMachineNumber, dwEnrollNmber, name, passWord, privilege, Enabled)
            DoEvents
            If rs1.State = 1 Then rs1.Close
            rs1.Open "select COUNT(*) as recs FROM m_enroll WHERE ip_address = '" & FG_IP_ADDRESS _
                & "' AND enrollnumber=" & Val(dwEnrollNmber), CnG, adOpenStatic, adLockReadOnly
            
            If rs1.Fields("recs").Value = 0 Then
                str1 = "INSERT INTO m_enroll(ip_address, enrollnumber, " _
                & "enrollname,emachinenumber,fingernumber,privilige,password) " _
                & "VALUES('" & FG_IP_ADDRESS & "'," & dwEnrollNmber & ",'', " _
                & dwMachineNumber & ", 1," & privilege & ",'" & passWord & "')"
                
                CnG.Execute str1
            End If
            
            i = i + 1
            
            If i = 10 Then
                k = k
            ElseIf i = 30 Then
                k = k
            ElseIf i = 50 Then
                k = k
            End If
            
        Wend
        CnG.CommitTrans
    End If
    
    get_all_user_info_tipeB = i
    Exit Function
    
Err:
    CnG.RollbackTrans
    MsgBox "Failed to read device...", vbCritical, headerMSG
    Call disconnect
    get_all_user_info_tipeB = 0
End Function


Private Sub TDBGrid2_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid2.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid2.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid2.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid2.Columns(ColIndex).DataField
    Call load_data_employee
'    TDBGrid_Employee.DataSource = rsEmployee
'    TDBGrid_Employee.Refresh
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid1.Columns(ColIndex).DataField Then
        oClause1 = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " DESC"
    Else
        oClause1 = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid1.Columns(ColIndex).DataField
    Call load_data_enroll
'    TDBGrid_Employee.DataSource = rsEmployee
'    TDBGrid_Employee.Refresh
End Sub
