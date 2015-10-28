VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_overtime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER OVERTIME"
   ClientHeight    =   6990
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
   Icon            =   "frm_mst_overtime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry 
      Height          =   2325
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   11295
      Begin VB.ComboBox txt_pengali 
         Height          =   315
         ItemData        =   "frm_mst_overtime.frx":058A
         Left            =   5310
         List            =   "frm_mst_overtime.frx":059D
         TabIndex        =   30
         Text            =   "1.5"
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5310
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1500
         Width           =   3495
      End
      Begin VB.TextBox txt_ot_upper 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5310
         TabIndex        =   16
         Top             =   780
         Width           =   825
      End
      Begin VB.TextBox txt_ot_under 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5310
         TabIndex        =   15
         Top             =   420
         Width           =   825
      End
      Begin VB.ComboBox cbo_shiftable 
         Height          =   315
         ItemData        =   "frm_mst_overtime.frx":05B2
         Left            =   120
         List            =   "frm_mst_overtime.frx":05BC
         TabIndex        =   4
         Text            =   "..."
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   4080
         TabIndex        =   21
         Top             =   1500
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MULTIPLIER*"
         Height          =   195
         Left            =   4110
         TabIndex        =   19
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TO (HOUR)*"
         Height          =   195
         Left            =   4170
         TabIndex        =   18
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FROM (HOUR)*"
         Height          =   195
         Left            =   3960
         TabIndex        =   17
         Top             =   420
         Width           =   1125
      End
   End
   Begin VB.TextBox txt_ot_name_type 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3390
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   5
      Top             =   870
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   5430
      Width           =   11295
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   570
         TabIndex        =   22
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&New Dtl"
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
         MICON           =   "frm_mst_overtime.frx":05C9
         PICN            =   "frm_mst_overtime.frx":05E5
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
         Left            =   1590
         TabIndex        =   23
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Save Dtl"
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
         MICON           =   "frm_mst_overtime.frx":1677
         PICN            =   "frm_mst_overtime.frx":1693
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
         Left            =   2610
         TabIndex        =   24
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Edit Dtl"
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
         MICON           =   "frm_mst_overtime.frx":2725
         PICN            =   "frm_mst_overtime.frx":2741
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
         Left            =   3630
         TabIndex        =   25
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Delete Dtl"
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
         MICON           =   "frm_mst_overtime.frx":37D3
         PICN            =   "frm_mst_overtime.frx":37EF
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
         Left            =   4650
         TabIndex        =   26
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cancel Dtl"
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
         MICON           =   "frm_mst_overtime.frx":4881
         PICN            =   "frm_mst_overtime.frx":489D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton CmdNew_Master 
         Height          =   705
         Left            =   7890
         TabIndex        =   27
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
         MICON           =   "frm_mst_overtime.frx":592F
         PICN            =   "frm_mst_overtime.frx":594B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdDelete_All 
         Height          =   705
         Left            =   8910
         TabIndex        =   28
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
         MICON           =   "frm_mst_overtime.frx":69DD
         PICN            =   "frm_mst_overtime.frx":69F9
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
         Left            =   10170
         TabIndex        =   29
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
         MICON           =   "frm_mst_overtime.frx":7A8B
         PICN            =   "frm_mst_overtime.frx":7AA7
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_ot 
      Height          =   375
      Left            =   1590
      OleObjectBlob   =   "frm_mst_overtime.frx":8B39
      TabIndex        =   6
      Top             =   870
      Width           =   1695
   End
   Begin VB.Frame fra_entry1 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   3510
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txt_ot_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4890
         MaxLength       =   10
         TabIndex        =   13
         Top             =   570
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_name_ot 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4890
         MaxLength       =   50
         TabIndex        =   14
         Top             =   930
         Width           =   3495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OT NAME*"
         Height          =   195
         Left            =   4035
         TabIndex        =   11
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OT CODE*"
         Height          =   195
         Left            =   4020
         TabIndex        =   10
         Top             =   600
         Width           =   765
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4005
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7064
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO."
      Columns(0).DataField=   "ot_number"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "FROM (HOUR)"
      Columns(1).DataField=   "from_value"
      Columns(1).NumberFormat=   "General Number"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TO (HOUR)"
      Columns(2).DataField=   "to_value"
      Columns(2).NumberFormat=   "General Number"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "MULTIPLIER"
      Columns(3).DataField=   "pengali"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DESCRIPTION"
      Columns(4).DataField=   "description"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2831"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2752"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2831"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2752"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=1768976485"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2487"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2408"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(3)._MinWidth=1768976485"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=9657"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=9578"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(4)._MinWidth=1768976485"
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
      Caption         =   "LIST OVERTIME"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
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
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER OVERTIME"
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
      TabIndex        =   12
      Top             =   180
      Width           =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "TYPE OVERTIME"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   930
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_overtime.frx":AADE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11970
   End
End
Attribute VB_Name = "frm_mst_overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsot As New ADODB.Recordset
Dim rsOT_Detail As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Dim vNo As Integer

'Private Function check_validate_exist_new() As Boolean
'Dim rs As New ADODB.Recordset
'Dim str_sql As String
'    check_validate_exist_new = False
'
'    str_sql = "select count(ot_number) as rec_count from m_ot_detail " _
'            & "where ot_code = '" & TDBCombo_ot.Text & "'"
'    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
'
'    If rs.Fields("rec_count").Value > 0 Then
'        check_validate_exist_new = True
'        Exit Function
'    End If
'End Function
'
'Private Sub check_invalid()
'    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
'
'End Sub
'
'Private Function check_validate_exist_edit() As Boolean
'    check_validate_exist_edit = False
'
'    If Not TDBCombo_company.Text = TDBCombo_company.Text And _
'    check_validate_exist_new Then
'        check_validate_exist_edit = True
'        Exit Function
'    End If
'End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True

    'validasi dari
    If Trim(txt_ot_under) = "" Then
        MsgBox "Kolom Dari Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_ot_under.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi sampai
    If Trim(txt_ot_upper) = "" Then
        MsgBox "Kolom Sampai Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_ot_upper.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi Pengali
    If Trim(txt_pengali) = "" Then
        MsgBox "Kolom Pengali Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_pengali.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

'Private Sub load_data()
'timer1.Enabled = True
'End Sub

Private Sub cmd_refresh_Click()
    Call load_data_ot
End Sub

Private Sub CmdCancel_Click()
    int_mode = 0
    Call load_mode
    CmdNew_Master.Caption = "&Tambah"
    cmdCancel.Caption = "&Batal Dtl"
End Sub

Private Sub cmdDelete_All_Click()
Dim i As Integer

On Error GoTo Err
    i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & txt_ot_name_type.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_ot_detail where " _
                & "ot_code = '" & TDBCombo_ot.Text & "'"
        CnG.Execute "delete from m_ot where ot_code = " _
                & "'" & TDBCombo_ot.Text & "'"
        
        CnG.CommitTrans
        
        Call load_mode
        Call load_data_ot
        int_mode = 0
        Call load_mode
        
        TDBCombo_ot.Text = ""
        txt_ot_name_type.Text = ""
        Set TDBGrid1.DataSource = Nothing
        
        Exit Sub
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG

End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
On Error GoTo Err
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
        & TDBGrid1.Columns("ot_number").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    CnG.Execute "delete from m_ot_detail where ot_code = '" & _
        TDBCombo_ot.Text & "' " & _
        "AND ot_number = '" & TDBGrid1.Columns("ot_number").Value & "'"
    CnG.CommitTrans
    
    Call load_data
    int_mode = 0
    Call load_mode
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
On Error GoTo Err
    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsOT_Detail
        txt_ot_under = .Fields("from_value").Value
        txt_ot_upper = .Fields("to_value").Value
        txt_pengali = .Fields("pengali").Value
        txt_description = .Fields("description").Value
        vNo = .Fields("ot_number").Value
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

Private Sub CmdNew_Master_Click()
Dim SQL As String
    
    If CmdNew_Master.Caption = "&New" Then
        CmdNew_Master.Caption = "&Save"
        cmdCancel.Caption = "&Cancel"
        Call set_buttons_enable(False, False, False, False, True, False, False)
        fra_entry1.Visible = True
        fra_entry.Visible = False
        
        txt_ot_code.Text = ""
        txt_name_ot.Text = ""
        txt_ot_code.SetFocus
        
        cmdDelete_All.Enabled = False
        CmdNew_Master.Enabled = True
    Else
        SQL = "INSERT INTO m_ot(ot_code,ot_name) " _
                & "VALUES ('" & txt_ot_code & "','" & txt_name_ot & "')"
        CnG.Execute SQL
        
        Call set_buttons_enable(True, False, True, True, False, True, True)
        CmdNew_Master.Caption = "&Add"
        cmdCancel.Caption = "&Cancel Dtl"
            
        fra_entry1.Visible = False
        
        cmdDelete_All.Enabled = True
        
        Call load_data_ot
        
        TDBCombo_ot.Text = txt_ot_code
        txt_ot_name_type.Text = txt_name_ot
        
        TDBGrid1.DataSource = Nothing
        
        txt_ot_code.Text = ""
        txt_name_ot.Text = ""
    End If
End Sub

Private Sub insert_new_data()

On Error GoTo Err
    CnG.BeginTrans
    
    SQL = "SELECT MAX(ot_number) ot_number FROM m_ot_detail WHERE ot_code = '" & TDBCombo_ot.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vNo = IIf(IsNull(rs!ot_number), 0, rs!ot_number)
    Else
        vNo = 0
    End If
    rs.Close
    
    vNo = vNo + 1
    
    SQL = "INSERT INTO m_ot_detail(ot_code,ot_number,from_value,to_value,pengali,description) " _
            & "VALUES " _
            & "('" & TDBCombo_ot.Text & "','" & vNo & "','" & Val(txt_ot_under.Text) & "'," _
            & "'" & Val(txt_ot_upper.Text) & "','" & Val(txt_pengali.Text) & "'," _
            & "'" & Trim(txt_description.Text) & "')"
    CnG.Execute SQL

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    CnG.BeginTrans
    
    SQL = "UPDATE m_ot_detail SET from_value = '" & Val(txt_ot_under.Text) & "'," _
            & "to_value = '" & Val(txt_ot_upper.Text) & "'," _
            & "pengali = '" & Val(txt_pengali.Text) & "'," _
            & "description = '" & Trim(txt_description.Text) & "' " _
            & "WHERE ot_number = '" & vNo & "' AND ot_code = '" & TDBCombo_ot.Text & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub CmdSave_Click()
    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
'        If check_validate_exist_new Then
'            Call check_invalid: Exit Sub
'        End If
        Call insert_new_data
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
'        If check_validate_exist_edit Then
'            Call check_invalid: Exit Sub
'        End If
        Call edit_old_data
    End If
    
    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal C As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
    cmdNew.Enabled = a And blnUser_Add
    cmdSave.Enabled = b
    cmdEdit.Enabled = C And blnUser_Edit
    cmdDelete.Enabled = d And blnUser_Delete
    cmdCancel.Enabled = e
    
    CmdNew_Master.Enabled = a And blnUser_Add
    cmdDelete_All.Enabled = d And blnUser_Delete
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_ot_name_type" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_ot" Then Ctr.Text = ""
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
'    TDBCombo_company.Text = ""
'    txt_company_name.Text = ""
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
                
        If Trim(TDBCombo_ot) = "" Then
            MsgBox "Tipe OT Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_ot.SetFocus
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        fra_entry.Visible = True
        cmdCancel.Caption = "&Cancel Dtl"
        TDBGrid1.Enabled = False
        Call set_new_data
        
        txt_ot_under.SetFocus
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry.Visible = False
        fra_entry1.Visible = False
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
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
    Call load_data_ot
    oClause = ""
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsOT_Detail.Filter = adFilterNone
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
    Set frm_mst_jamsostek = Nothing
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsOT_Detail.Filter = getFilter()
    TDBGrid1.Col = i
    TDBGrid1.EditActive = True
    
    TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
    If TDBGrid1.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid1.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
    If rsOT_Detail.State Then rsOT_Detail.Close
    SQL = "select * " _
            & "from m_ot_detail " _
            & "where ot_code = '" & TDBCombo_ot.Text & "' " & oClause
    rsOT_Detail.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsOT_Detail
End Sub

Private Sub Timer1_Timer()
    Call load_data_ot
    timer1.Enabled = False
End Sub

Private Sub load_data_ot()
    If rsot.State Then rsot.Close
    SQL = "select * from m_ot order by ot_code"
    rsot.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_ot.RowSource = rsot
End Sub

Private Sub TDBCombo_ot_ItemChange()
    If TDBCombo_ot.ApproxCount > 0 Then
        TDBCombo_ot.Text = TDBCombo_ot.Columns("ot_code").Value
        txt_ot_name_type = TDBCombo_ot.Columns("ot_name").Value
        
        Call load_data
    End If
End Sub

Private Sub txt_ot_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_name_ot_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
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

