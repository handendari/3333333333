VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_komponen_sal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER SALARY COMPONENT"
   ClientHeight    =   6690
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
   Icon            =   "frm_mst_komponen_sal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_entry 
      Height          =   2655
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   11295
      Begin VB.TextBox txt_formula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1410
         Width           =   3495
      End
      Begin VB.ComboBox cbo_type 
         Height          =   315
         ItemData        =   "frm_mst_komponen_sal.frx":058A
         Left            =   4800
         List            =   "frm_mst_komponen_sal.frx":0594
         TabIndex        =   3
         Text            =   "INCOME"
         Top             =   1050
         Width           =   1725
      End
      Begin VB.CheckBox chk_formula 
         Caption         =   "NO"
         Height          =   225
         Left            =   4800
         TabIndex        =   7
         Top             =   2130
         Width           =   1305
      End
      Begin VB.ComboBox cbo_tax_type 
         Height          =   315
         ItemData        =   "frm_mst_komponen_sal.frx":05A9
         Left            =   4800
         List            =   "frm_mst_komponen_sal.frx":05B3
         TabIndex        =   5
         Text            =   "ANNUALIZED"
         Top             =   1770
         Width           =   1725
      End
      Begin VB.TextBox txt_komp_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   690
         Width           =   3495
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
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_komp_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FORMULA*"
         Height          =   195
         Left            =   3180
         TabIndex        =   23
         Top             =   1410
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TYPE*"
         Height          =   195
         Left            =   3180
         TabIndex        =   21
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PART OF FORMULA"
         Height          =   195
         Left            =   3180
         TabIndex        =   20
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TAX TYPE*"
         Height          =   195
         Left            =   3180
         TabIndex        =   12
         Top             =   1770
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CODE*"
         Height          =   195
         Left            =   3180
         TabIndex        =   11
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "COMPONENT NAME*"
         Height          =   195
         Left            =   3180
         TabIndex        =   10
         Top             =   690
         Width           =   1500
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   11295
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   690
         TabIndex        =   13
         Top             =   300
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
         MICON           =   "frm_mst_komponen_sal.frx":05D3
         PICN            =   "frm_mst_komponen_sal.frx":05EF
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
         Left            =   1710
         TabIndex        =   14
         Top             =   300
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
         MICON           =   "frm_mst_komponen_sal.frx":1681
         PICN            =   "frm_mst_komponen_sal.frx":169D
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
         Left            =   2730
         TabIndex        =   15
         Top             =   300
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
         MICON           =   "frm_mst_komponen_sal.frx":272F
         PICN            =   "frm_mst_komponen_sal.frx":274B
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
         Left            =   3750
         TabIndex        =   16
         Top             =   300
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
         MICON           =   "frm_mst_komponen_sal.frx":37DD
         PICN            =   "frm_mst_komponen_sal.frx":37F9
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
         Left            =   4770
         TabIndex        =   17
         Top             =   300
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
         MICON           =   "frm_mst_komponen_sal.frx":488B
         PICN            =   "frm_mst_komponen_sal.frx":48A7
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
         Left            =   9810
         TabIndex        =   18
         Top             =   330
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
         MICON           =   "frm_mst_komponen_sal.frx":5939
         PICN            =   "frm_mst_komponen_sal.frx":5955
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_Select 
         Height          =   705
         Left            =   7770
         TabIndex        =   22
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Select"
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
         MICON           =   "frm_mst_komponen_sal.frx":69E7
         PICN            =   "frm_mst_komponen_sal.frx":6A03
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
      Height          =   4245
      Left            =   240
      TabIndex        =   0
      Top             =   810
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7488
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CODE"
      Columns(0).DataField=   "komp_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "COMPONENT NAME"
      Columns(1).DataField=   "komp_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TYPE"
      Columns(2).DataField=   "tipe"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "FORMULA"
      Columns(3).DataField=   "formula"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TAX TYPE"
      Columns(4).DataField=   "tipe_pjk"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "PART OF FORMULA"
      Columns(5).DataField=   "part_formula"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2170"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4630"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4551"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2487"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2408"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=4180"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=4101"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2540"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2461"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2805"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2725"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
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
      Caption         =   "LIST OF SALARY COMPONENT"
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
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER SALARY COMPONENT"
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
      TabIndex        =   19
      Top             =   150
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   -60
      Picture         =   "frm_mst_komponen_sal.frx":7A95
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frm_mst_komponen_sal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsKomp As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks
Public public_int_mode As Integer

Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False
    
    If rsKomp.State Then rsKomp.Close
    str_sql = "select count(komp_code) as rec_count from m_sal_komp where komp_code = '" _
    & Replace$(Trim$(txt_komp_code.Text), Chr$(39), Chr$(96)) & "'"
    rsKomp.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rsKomp.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        rsKomp.Close
        Exit Function
    End If
    
    rsKomp.Close
End Function

Private Function check_validate_exist_new_dtl(ByVal str_komp_code As String, ByVal str_group_code As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new_dtl = False

    str_sql = "select count(*) as rec_count from m_sal_komp_group_dtl where group_code='" & str_group_code _
                & "' and komp_code = '" & str_komp_code & "'"
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new_dtl = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    txt_komp_code = ""
    If txt_komp_code.Enabled = True Then txt_komp_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not txt_komp_code = rsKomp.Fields("komp_code").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    'validasi component code
    If Trim(txt_komp_code) = "" Then
        MsgBox "Component Code is empty!", vbOKOnly + vbInformation, headerMSG
        txt_komp_code.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi component name
    If Trim(txt_komp_name) = "" Then
        MsgBox "Component Name is empty!", vbOKOnly + vbInformation, headerMSG
        txt_komp_name.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

Private Sub load_data()
    Timer1.Enabled = True
End Sub

Private Sub chk_formula_Click()
    If chk_formula.Value = 0 Then
        chk_formula.Caption = "NO"
    Else
        chk_formula.Caption = "YES"
    End If
End Sub

Private Sub cmd_select_Click()
Dim i As Integer
Dim item

On Error GoTo Err

    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid1.SelBookmarks
    i = 0
    
    CnG.BeginTrans
    If public_int_mode = 0 Then
        For Each item In SelBks
            i = i + 1
            
            If check_validate_exist_new_dtl(TDBGrid1.Columns("komp_code").CellText(item), _
            frm_mst_komponen_sal_group.TDBGrid1.Columns("group_code").Value) = False Then
                
                CnG.Execute "insert into m_sal_komp_group_dtl values " & _
                    "('" & frm_mst_komponen_sal_group.TDBGrid1.Columns("group_code").Value & "', " & _
                    "'" & TDBGrid1.Columns("komp_code").CellText(item) & "', " & _
                    "now(),'" & LOGIN_NAME & "')"
                                                                                        
            End If
                            
        Next
    End If
    CnG.CommitTrans
    Call frm_mst_komponen_sal_group.load_data_komp_group_dtl
    MsgBox i & " component's data are successfully added", vbInformation, headerMSG
    
    Call CmdExit_Click
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Are you sure want to delete data '" _
        & TDBGrid1.Columns("komp_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    CnG.Execute "delete from m_sal_komp where komp_code = '" _
        & TDBGrid1.Columns("komp_code").Value & "'"
    CnG.CommitTrans
    
    Call load_data_komp
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
On Error GoTo Err
    With rsKomp
        txt_komp_code = .Fields("komp_code").Value
        txt_komp_name = .Fields("komp_name").Value
        cbo_type.ListIndex = .Fields("type").Value
        txt_formula.Text = .Fields("formula").Value
        cbo_tax_type.ListIndex = .Fields("tax_type").Value
        chk_formula.Value = .Fields("part_formula").Value
    End With
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
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
    
    cbo_type.ListIndex = 0
    cbo_tax_type.ListIndex = 0
    chk_formula.Value = 0
End Sub

Private Sub insert_new_data()
On Error GoTo Err
    CnG.BeginTrans
    
    SQL = "INSERT INTO m_sal_komp(komp_code,komp_name,type,formula,tax_type,part_formula,entry_date,entry_user) " & _
            "VALUES " & _
            "('" & Trim(txt_komp_code.Text) & "','" & Trim(txt_komp_name.Text) & "'," & _
            "'" & cbo_type.ListIndex & "','" & txt_formula.Text & "','" & cbo_tax_type.ListIndex & "'," & _
            "'" & chk_formula.Value & "',now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err
    CnG.BeginTrans
    
    SQL = "UPDATE m_sal_komp SET komp_code = '" & Trim(txt_komp_code.Text) & "'," & _
            "komp_name = '" & Trim(txt_komp_name.Text) & "'," & _
            "type = '" & cbo_type.ListIndex & "'," & _
            "formula = '" & txt_formula.Text & "'," & _
            "tax_type = '" & cbo_tax_type.ListIndex & "'," & _
            "part_formula = '" & chk_formula.Value & "'," & _
            "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
            "WHERE komp_code = '" & Trim(txt_komp_code.Text) & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
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
    
    Call load_data_komp
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
    
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        fra_entry.Visible = True
        txt_komp_code.Enabled = True
        TDBGrid1.Enabled = False
        Call set_new_data
        
        If txt_komp_code.Enabled = True Then
            txt_komp_code.SetFocus
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry.Visible = False
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        txt_komp_code.Enabled = False
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
    cbo_tax_type.ListIndex = 0
    cbo_type.ListIndex = 0
    cmd_select.Visible = False
    
    Call load_data
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsKomp.Filter = adFilterNone
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
    Set frm_mst_title = Nothing
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err

    Dim i As Integer
    
    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsKomp.Filter = getFilter()
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

Private Sub load_data_komp()
    If rsKomp.State Then rsKomp.Close
    SQL = "select a.*,case when tax_type = 0 then 'ANNUALIZED' else 'NOT ANNUALIZED' end tipe_pjk," & _
            "case when type = 0 then 'INCOME' else 'EXPENSE' end tipe " & _
          "from m_sal_komp a"
    rsKomp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsKomp
End Sub

Private Sub Timer1_Timer()
    Call load_data_komp
    Timer1.Enabled = False
End Sub

Private Sub txt_komp_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_komp_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_formula_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
