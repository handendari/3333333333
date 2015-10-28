VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_thr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER THR"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11565
   Icon            =   "frmMstTHR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry_thr 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   2010
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CheckBox chk_percentage 
         Caption         =   "Percentage"
         Height          =   225
         Left            =   7890
         TabIndex        =   5
         Top             =   1410
         Width           =   1245
      End
      Begin VB.TextBox txt_thr_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   630
         Width           =   1695
      End
      Begin VB.TextBox txt_thr_under 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox txt_thr_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1740
         Width           =   3495
      End
      Begin VB.TextBox txt_thr_upper 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox txt_thr_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "FROM PARAMETER*"
         Height          =   195
         Left            =   4470
         TabIndex        =   17
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "NO.*"
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   4470
         TabIndex        =   15
         Top             =   1770
         Width           =   1020
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "TO PARAMETER*"
         Height          =   195
         Left            =   4470
         TabIndex        =   14
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "VALUE*"
         Height          =   195
         Left            =   4470
         TabIndex        =   13
         Top             =   1380
         Width           =   585
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   11295
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   540
         TabIndex        =   7
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
         MICON           =   "frmMstTHR.frx":058A
         PICN            =   "frmMstTHR.frx":05A6
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
         Left            =   1560
         TabIndex        =   8
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
         MICON           =   "frmMstTHR.frx":1638
         PICN            =   "frmMstTHR.frx":1654
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
         Left            =   2580
         TabIndex        =   9
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
         MICON           =   "frmMstTHR.frx":26E6
         PICN            =   "frmMstTHR.frx":2702
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
         Left            =   3600
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
         MICON           =   "frmMstTHR.frx":3794
         PICN            =   "frmMstTHR.frx":37B0
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
         Left            =   4620
         TabIndex        =   11
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
         MICON           =   "frmMstTHR.frx":4842
         PICN            =   "frmMstTHR.frx":485E
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
         Left            =   9870
         TabIndex        =   20
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
         MICON           =   "frmMstTHR.frx":58F0
         PICN            =   "frmMstTHR.frx":590C
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Thr 
      Height          =   3885
      Left            =   120
      TabIndex        =   18
      Top             =   900
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6853
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NUMBER"
      Columns(0).DataField=   "thr_number"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "FROM (MONTH)"
      Columns(1).DataField=   "thr_under"
      Columns(1).NumberFormat=   "Standard"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TO (MONTH)"
      Columns(2).DataField=   "thr_upper"
      Columns(2).NumberFormat=   "Standard"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "VALUE"
      Columns(3).DataField=   "thr_value"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PERCENTAGE"
      Columns(4).DataField=   "flag_percentage"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DESCRIPTION"
      Columns(5).DataField=   "description"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3043"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2963"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2858"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2778"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2355"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2011"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1931"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=6112"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=6033"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
      Caption         =   "LIST OF THR"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=3"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(58)  =   "Named:id=33:Normal"
      _StyleDefs(59)  =   ":id=33,.parent=0"
      _StyleDefs(60)  =   "Named:id=34:Heading"
      _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=34,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=35:Footing"
      _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=36:Selected"
      _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=37:Caption"
      _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(69)  =   "Named:id=38:HighlightRow"
      _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=39:EvenRow"
      _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=40:OddRow"
      _StyleDefs(74)  =   ":id=40,.parent=33"
      _StyleDefs(75)  =   "Named:id=41:RecordSelector"
      _StyleDefs(76)  =   ":id=41,.parent=34"
      _StyleDefs(77)  =   "Named:id=42:FilterBar"
      _StyleDefs(78)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER THR"
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
      Left            =   180
      TabIndex        =   19
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frmMstTHR.frx":699E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
End
Attribute VB_Name = "frm_mst_thr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsthr As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim strsql As String

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new = False
    
    str_sql = "select count(thr_number) as rec_count from m_thr where thr_number = '" & Trim(txt_thr_number) & "'"
    
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    
    txt_thr_number = ""
    If txt_thr_number.Enabled = True Then txt_thr_number.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not txt_thr_number.Text = rsthr.Fields("thr_number").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    'validasi thr number
    If Trim(txt_thr_number.Text) = "" Then
        MsgBox "THR Number is empty!", vbOKOnly + vbInformation, headerMSG
        txt_thr_number.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi thr under
    If Trim(txt_thr_under.Text) = "" Then
        MsgBox "Under Parameter is empty!", vbOKOnly + vbInformation, headerMSG
        txt_thr_under.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi thr upper
    If Trim(txt_thr_upper.Text) = "" Then
        MsgBox "Upper Parameter is empty!", vbOKOnly + vbInformation, headerMSG
        txt_thr_upper.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi thr grade
    If Trim(txt_thr_value.Text) = "" Then
        MsgBox "THR Value is empty!", vbOKOnly + vbInformation, headerMSG
        txt_thr_value.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

Private Sub load_data()
    If rsthr.State Then rsthr.Close
    SQL = "select * from m_thr " & oClause
    rsthr.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Thr.DataSource = rsthr
End Sub

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
On Error GoTo Err
Dim i As Integer
    CnG.BeginTrans
            
    If Not (TDBGrid_Thr.ApproxCount > 0 And TDBGrid_Thr.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Are you sure want to delete data '" _
        & TDBGrid_Thr.Columns("thr_value").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.Execute "delete from m_thr where thr_number = '" _
                    & Replace(TDBGrid_Thr.Columns("thr_number").Value, "'", "''") & "'"
    
    CnG.CommitTrans
        
    Call load_data
    int_mode = 0
    Call load_mode
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
'On Error GoTo Err
    vSetData = 1
    
    If Not (TDBGrid_Thr.ApproxCount > 0 And TDBGrid_Thr.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsthr
        txt_thr_number.Text = .Fields("thr_number").Value
        txt_thr_under.Text = FormatNumber(.Fields("thr_under").Value)
        txt_thr_upper.Text = FormatNumber(.Fields("thr_upper").Value)
        chk_percentage.Value = FormatNumber(.Fields("flag_percentage").Value)
        txt_thr_value.Text = .Fields("thr_value").Value
        txt_thr_description.Text = .Fields("description").Value
    End With

    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub new_data()
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()
On Error GoTo Err
    CnG.BeginTrans
    
    SQL = "INSERT INTO m_thr (thr_number,thr_under," & _
            "thr_upper,flag_percentage,thr_value,description,entry_date,entry_user) " & _
          "VALUES( " & _
            "'" & Trim(txt_thr_number.Text) & "','" & Val(DropAllComma(txt_thr_under)) & "'," & _
            "'" & Val(DropAllComma(txt_thr_upper)) & "','" & chk_percentage.Value & "','" & Val(DropAllComma(txt_thr_value)) & "'," & _
            "'" & Trim(txt_thr_description) & "',now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    CnG.BeginTrans
    
    SQL = "UPDATE m_thr SET thr_number = '" & Trim(txt_thr_number.Text) & "'," & _
            "thr_under = '" & Val(DropAllComma(txt_thr_under.Text)) & "'," & _
            "thr_upper = '" & Val(DropAllComma(txt_thr_upper.Text)) & "'," & _
            "flag_percentage = '" & chk_percentage.Value & "'," & _
            "thr_value = '" & Val(DropAllComma(txt_thr_value.Text)) & "'," & _
            "description = '" & Trim(txt_thr_description.Text) & "'," & _
            "edit_date = now(),edit_user = '" & LOGIN_CODE & "' " & _
          "WHERE thr_number = '" & Trim(txt_thr_number.Text) & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub simpan_data()
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
            If Not LCase(Ctr.name) = "txt_company" Then Ctr.Text = ""
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
    chk_percentage.Value = 0
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        fra_entry_thr.Visible = True
        txt_thr_number.Enabled = True
        TDBGrid_Thr.Enabled = False
        
        Call set_new_data

        If txt_thr_number.Enabled = True Then
            txt_thr_number.SetFocus
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry_thr.Visible = False
        TDBGrid_Thr.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        txt_thr_number.Enabled = False
        fra_entry_thr.Visible = True
        TDBGrid_Thr.Enabled = False
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
    Call load_data
    oClause = ""
    
    Call load_data_user_access(Me)
    int_mode = 0

    Call load_mode
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid_Thr.Columns
        Col.FilterText = ""
    Next Col
    rsthr.Filter = adFilterNone
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

    Set Cols = TDBGrid_Thr.Columns
    i = TDBGrid_Thr.Col
    TDBGrid_Thr.HoldFields
    
    rsthr.Filter = getFilter()
    TDBGrid_Thr.Col = i
    TDBGrid_Thr.EditActive = True
    
    TDBGrid_Thr.SelStart = Len(TDBGrid_Thr.Columns(i).FilterText)
    If TDBGrid_Thr.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid_Thr.Col = i
    End If

    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_thr = Nothing
End Sub

Private Sub cmdNew_Click()
    Call new_data
End Sub

Private Sub cmdSave_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Click()
    Call cancel_data
End Sub

Private Sub TDBGrid_thr_FilterChange()
    Call grid_filter
End Sub

Private Sub txt_thr_value_Validate(Cancel As Boolean)
    If Not Trim(txt_thr_value) = "" Then
        txt_thr_value = FormatNumber(DropAllComma(txt_thr_value))
    End If
End Sub

Private Sub TDBGrid_Thr_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Thr.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Thr.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Thr.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Thr.Columns(ColIndex).DataField
    Call load_data

End Sub
