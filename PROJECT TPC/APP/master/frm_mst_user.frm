VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_user 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER USER"
   ClientHeight    =   10755
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   11760
   Icon            =   "frm_mst_user.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10350
      TabIndex        =   0
      Top             =   9960
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
      MICON           =   "frm_mst_user.frx":058A
      PICN            =   "frm_mst_user.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraPasteDtl 
      Caption         =   "Paste Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   3120
      TabIndex        =   53
      Top             =   4620
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox cboFileName 
         Height          =   315
         Left            =   1830
         TabIndex        =   9
         Top             =   510
         Width           =   2955
      End
      Begin prj_tpc.vbButton cmd_paste_cancel 
         Height          =   705
         Left            =   3810
         TabIndex        =   54
         Top             =   1260
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
         MICON           =   "frm_mst_user.frx":1638
         PICN            =   "frm_mst_user.frx":1654
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_paste_ok 
         Height          =   705
         Left            =   2760
         TabIndex        =   55
         Top             =   1260
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&OK"
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
         MICON           =   "frm_mst_user.frx":26E6
         PICN            =   "frm_mst_user.frx":2702
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TEMPLATE NAME*"
         Height          =   195
         Left            =   330
         TabIndex        =   10
         Top             =   570
         Width           =   1425
      End
   End
   Begin VB.Frame fraCopyDtl 
      Caption         =   "Copy Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   3120
      TabIndex        =   2
      Top             =   4620
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   49
         Top             =   570
         Width           =   2895
      End
      Begin prj_tpc.vbButton cmd_copy_cancel 
         Height          =   705
         Left            =   3870
         TabIndex        =   52
         Top             =   1260
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
         MICON           =   "frm_mst_user.frx":3794
         PICN            =   "frm_mst_user.frx":37B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_copy_ok 
         Height          =   705
         Left            =   2820
         TabIndex        =   50
         Top             =   1260
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&OK"
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
         MICON           =   "frm_mst_user.frx":4842
         PICN            =   "frm_mst_user.frx":485E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "TEMPLATE NAME*"
         Height          =   195
         Left            =   360
         TabIndex        =   51
         Top             =   600
         Width           =   1425
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9195
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   16219
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SUPER USER"
      TabPicture(0)   =   "frm_mst_user.frx":58F0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSTabUser"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "USER"
      TabPicture(1)   =   "frm_mst_user.frx":590C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraAdmin"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraAdmin 
         Height          =   8655
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   10905
         Begin prj_tpc.LynxGrid LynxGrid2 
            Height          =   3255
            Left            =   2940
            TabIndex        =   65
            Top             =   3060
            Visible         =   0   'False
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   5741
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
         Begin VB.TextBox txt_section 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   1110
            Width           =   4815
         End
         Begin VB.Frame fra_entry_user 
            Caption         =   "Entry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1725
            Left            =   180
            TabIndex        =   21
            Top             =   1500
            Visible         =   0   'False
            Width           =   10515
            Begin VB.TextBox txtkd_employee 
               Height          =   315
               Left            =   5700
               TabIndex        =   63
               Top             =   810
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txt_nmemployee 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000B&
               Height          =   315
               Left            =   4530
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   62
               Top             =   1170
               Width           =   3495
            End
            Begin VB.TextBox txt_nik 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2760
               MaxLength       =   10
               TabIndex        =   61
               Top             =   1170
               Width           =   1335
            End
            Begin VB.CheckBox chk_user_for 
               Caption         =   "ALL COMPANY"
               Height          =   225
               Left            =   8220
               TabIndex        =   26
               Top             =   480
               Width           =   1875
            End
            Begin VB.TextBox txt_KodeUser 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               Height          =   300
               Left            =   330
               MaxLength       =   40
               TabIndex        =   25
               Top             =   930
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txt_user_code 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   300
               Left            =   330
               MaxLength       =   40
               TabIndex        =   23
               Top             =   540
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txt_NamaUser 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   300
               Left            =   2760
               MaxLength       =   40
               TabIndex        =   22
               Top             =   450
               Width           =   2895
            End
            Begin VB.TextBox txt_PasswordUser 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   2760
               PasswordChar    =   "*"
               TabIndex        =   24
               Top             =   810
               Width           =   2895
            End
            Begin prj_tpc.vbButton cmd_periode_browse_employee 
               Height          =   315
               Left            =   4110
               TabIndex        =   64
               Top             =   1170
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
               MICON           =   "frm_mst_user.frx":5928
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "USER FOR"
               Height          =   195
               Left            =   6720
               TabIndex        =   30
               Top             =   480
               Width           =   825
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "EMPLOYEE*"
               Height          =   195
               Left            =   1560
               TabIndex        =   29
               Top             =   1200
               Width           =   930
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "USERNAME*"
               Height          =   195
               Left            =   1560
               TabIndex        =   28
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "PASSWORD*"
               Height          =   195
               Left            =   1560
               TabIndex        =   27
               Top             =   840
               Width           =   1005
            End
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   300
            Left            =   120
            Top             =   8070
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   30
            Left            =   -750
            TabIndex        =   7
            Top             =   7800
            Width           =   12585
         End
         Begin VB.TextBox txt_division 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   690
            Width           =   4815
         End
         Begin VB.TextBox txt_company_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   270
            Width           =   4815
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Level 
            Height          =   1965
            Left            =   180
            TabIndex        =   8
            Top             =   3360
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   3466
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "LEVEL CODE"
            Columns(0).DataField=   "access_level_code"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "LEVEL NAME"
            Columns(1).DataField=   "level_name"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "ACCESS"
            Columns(2).DataField=   "allow_access"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).Size  =   2
            Splits(0).Size.vt=   2
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   3
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2619"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=6482"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6403"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1667"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1588"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=3342336"
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
            Caption         =   "LIST OF ACCESS LEVEL"
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
            _StyleDefs(8)   =   ":id=4,.fgcolor=&H80000009&,.bold=0,.fontsize=900,.italic=0,.underline=0"
            _StyleDefs(9)   =   ":id=4,.strikethrough=0,.charset=0"
            _StyleDefs(10)  =   ":id=4,.fontname=Microsoft Sans Serif"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
         Begin prj_tpc.vbButton cmdNewUser 
            Height          =   705
            Left            =   180
            TabIndex        =   11
            Top             =   7860
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
            MICON           =   "frm_mst_user.frx":5944
            PICN            =   "frm_mst_user.frx":5960
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdSimpanUser 
            Height          =   705
            Left            =   1200
            TabIndex        =   12
            Top             =   7860
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
            MICON           =   "frm_mst_user.frx":69F2
            PICN            =   "frm_mst_user.frx":6A0E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdEditUser 
            Height          =   705
            Left            =   2220
            TabIndex        =   13
            Top             =   7860
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
            MICON           =   "frm_mst_user.frx":7AA0
            PICN            =   "frm_mst_user.frx":7ABC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdDeleteUser 
            Height          =   705
            Left            =   3240
            TabIndex        =   14
            Top             =   7860
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
            MICON           =   "frm_mst_user.frx":8B4E
            PICN            =   "frm_mst_user.frx":8B6A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmd_save_dtl 
            Height          =   705
            Left            =   6030
            TabIndex        =   15
            Top             =   7860
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
            MICON           =   "frm_mst_user.frx":9BFC
            PICN            =   "frm_mst_user.frx":9C18
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmd_delete_dtl 
            Height          =   705
            Left            =   7050
            TabIndex        =   16
            Top             =   7860
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
            MICON           =   "frm_mst_user.frx":ACAA
            PICN            =   "frm_mst_user.frx":ACC6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_form 
            Height          =   2265
            Left            =   180
            TabIndex        =   17
            Top             =   5400
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   3995
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "MENU"
            Columns(0).DataField=   "menu_name"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "FORM TITLE"
            Columns(1).DataField=   "form_title"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "READ"
            Columns(2).DataField=   "allow_read"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   4
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "ADD"
            Columns(3).DataField=   "allow_add"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   4
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "EDIT"
            Columns(4).DataField=   "allow_edit"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "DELETE"
            Columns(5).DataField=   "allow_delete"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   4
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "POST"
            Columns(6).DataField=   "allow_post"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   4
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "PRINT"
            Columns(7).DataField=   "allow_print"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).Size  =   2
            Splits(0).Size.vt=   2
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   3
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2170"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=5133"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5054"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1667"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1588"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=3342336"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1640"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1561"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=1693"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1614"
            Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(27)=   "Column(4)._MinWidth=182755968"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=1693"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1614"
            Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(33)=   "Column(5)._MinWidth=182748752"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=1667"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1588"
            Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(39)=   "Column(6)._MinWidth=182794208"
            Splits(0)._ColumnProps(40)=   "Column(7).Width=1746"
            Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=1667"
            Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(45)=   "Column(7)._MinWidth=182756384"
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
            Caption         =   "LIST OF PRIVILEGES"
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
            _StyleDefs(8)   =   ":id=4,.fgcolor=&H80000009&,.bold=0,.fontsize=900,.italic=0,.underline=0"
            _StyleDefs(9)   =   ":id=4,.strikethrough=0,.charset=0"
            _StyleDefs(10)  =   ":id=4,.fontname=Microsoft Sans Serif"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
         Begin prj_tpc.vbButton cmdGenerate 
            Height          =   705
            Left            =   5010
            TabIndex        =   18
            Top             =   7860
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Load Dtl"
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
            MICON           =   "frm_mst_user.frx":BD58
            PICN            =   "frm_mst_user.frx":BD74
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdCopyDtl 
            Height          =   705
            Left            =   8850
            TabIndex        =   19
            Top             =   7860
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Copy Dtl"
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
            MICON           =   "frm_mst_user.frx":CE06
            PICN            =   "frm_mst_user.frx":CE22
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_tpc.vbButton cmdPasteDtl 
            Height          =   705
            Left            =   9870
            TabIndex        =   20
            Top             =   7860
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Paste Dtl"
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
            MICON           =   "frm_mst_user.frx":DEB4
            PICN            =   "frm_mst_user.frx":DED0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   90
            Top             =   7800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Save File To"
            Filter          =   "*.ssd"
            InitDir         =   "C:\"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_User 
            Height          =   1635
            Left            =   180
            TabIndex        =   31
            Top             =   1560
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   2884
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "USER CODE"
            Columns(0).DataField=   "user_code"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "USER NAME"
            Columns(1).DataField=   "user_name"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "PASSWORD"
            Columns(2).DataField=   "user_pass"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "ID"
            Columns(3).DataField=   "employee_code"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "EMPLOYEE CODE"
            Columns(4).DataField=   "nik"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "EMPLOYEE NAME"
            Columns(5).DataField=   "employee_name"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "FLAG USER"
            Columns(6).DataField=   "flag_user"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "LEVEL CODE"
            Columns(7).DataField=   "level_code"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3440"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3360"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=4630"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4551"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=3651"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3572"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=900"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=820"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=7594"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=7514"
            Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2752"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2672"
            Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(44)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
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
            Caption         =   "LIST OF USER"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_company 
            Height          =   375
            Left            =   1050
            OleObjectBlob   =   "frm_mst_user.frx":EF62
            TabIndex        =   56
            Top             =   270
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_division 
            Height          =   375
            Left            =   1050
            OleObjectBlob   =   "frm_mst_user.frx":10F20
            TabIndex        =   57
            Top             =   690
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_section 
            Height          =   375
            Left            =   1050
            OleObjectBlob   =   "frm_mst_user.frx":12EDF
            TabIndex        =   59
            Top             =   1110
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SECTION"
            Height          =   195
            Left            =   270
            TabIndex        =   60
            Top             =   1170
            Width           =   705
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DIVISION"
            Height          =   195
            Left            =   270
            TabIndex        =   33
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   300
            Width           =   795
         End
      End
      Begin TabDlg.SSTab SSTabUser 
         Height          =   6705
         Left            =   -74820
         TabIndex        =   34
         Top             =   660
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   11827
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "SUPER ADMIN"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TDBGrid_SuperUser"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fra_entry_super_user"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fraTopCmd"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin VB.Frame fraTopCmd 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   660
            TabIndex        =   43
            Top             =   4950
            Width           =   7935
            Begin prj_tpc.vbButton cmdNewTop 
               Height          =   705
               Left            =   120
               TabIndex        =   44
               Top             =   30
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
               MICON           =   "frm_mst_user.frx":14EA5
               PICN            =   "frm_mst_user.frx":14EC1
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdSaveTop 
               Height          =   705
               Left            =   1140
               TabIndex        =   45
               Top             =   30
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
               MICON           =   "frm_mst_user.frx":15F53
               PICN            =   "frm_mst_user.frx":15F6F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdEditTop 
               Height          =   705
               Left            =   2160
               TabIndex        =   46
               Top             =   30
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
               MICON           =   "frm_mst_user.frx":17001
               PICN            =   "frm_mst_user.frx":1701D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdDeleteTop 
               Height          =   705
               Left            =   3180
               TabIndex        =   47
               Top             =   30
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
               MICON           =   "frm_mst_user.frx":180AF
               PICN            =   "frm_mst_user.frx":180CB
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
         Begin VB.Frame fra_entry_super_user 
            Caption         =   "Entry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Left            =   780
            TabIndex        =   35
            Top             =   3630
            Visible         =   0   'False
            Width           =   9405
            Begin VB.TextBox txt_full_name_super_user 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   300
               Left            =   2280
               MaxLength       =   200
               TabIndex        =   38
               Top             =   510
               Width           =   3855
            End
            Begin VB.TextBox txt_nama_super_user 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   300
               Left            =   2280
               MaxLength       =   40
               TabIndex        =   37
               Top             =   180
               Width           =   3855
            End
            Begin VB.TextBox txt_pwd_super_user 
               Appearance      =   0  'Flat
               BackColor       =   &H80000014&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   2280
               MaxLength       =   100
               PasswordChar    =   "*"
               TabIndex        =   40
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txt_kode_super_user 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               Height          =   300
               Left            =   270
               MaxLength       =   5
               TabIndex        =   36
               Top             =   510
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label9 
               Caption         =   "FULLNAME"
               Height          =   315
               Left            =   1200
               TabIndex        =   42
               Top             =   540
               Width           =   915
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "PASSWORD"
               Height          =   195
               Left            =   1200
               TabIndex        =   41
               Top             =   870
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "USERNAME"
               Height          =   195
               Left            =   1200
               TabIndex        =   39
               Top             =   180
               Width           =   915
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_SuperUser 
            Height          =   4365
            Left            =   780
            TabIndex        =   48
            Top             =   540
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   7699
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "USER CODE"
            Columns(0).DataField=   "user_code"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "USER NAME"
            Columns(1).DataField=   "user_name"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "PASSWORD"
            Columns(2).DataField=   "user_pass"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "FULL NAME"
            Columns(3).DataField=   "full_name"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3969"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3889"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=4048"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3969"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=5186"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5106"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=6297"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=6218"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
            Caption         =   "LIST OF SUPER ADMIN"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
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
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER USER"
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
      Left            =   360
      TabIndex        =   1
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_user.frx":1915D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frm_mst_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSuperUser As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDivision As New ADODB.Recordset
Dim rsSection As New ADODB.Recordset
Dim rsLevel As New ADODB.Recordset
Dim rsUser As New ADODB.Recordset
Dim rsAccessLevel As New ADODB.Recordset
Dim rsForm As New ADODB.Recordset

Dim nmFile As String
Dim conLite As New cConnection
Dim rsLite As cRecordset
Dim rsAtt As New ADODB.Recordset
Dim rss As New cRecordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Sub EnableButtonEntryUser _
(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, ByVal d As Boolean)
    If SSTab1.Tab = 0 Then
        cmdNewTop.Enabled = a And blnUser_Add
        cmdEditTop.Enabled = b And blnUser_Edit
        cmdDeleteTop.Enabled = c And blnUser_Delete
        cmdSaveTop.Enabled = d
    Else
        cmdNewUser.Enabled = a And blnUser_Add
        cmdEditUser.Enabled = b And blnUser_Edit
        cmdDeleteUser.Enabled = c And blnUser_Delete
        cmdSimpanUser.Enabled = d
        
        cmdGenerate.Enabled = a And blnUser_Add
        cmd_save_dtl.Enabled = b And blnUser_Add
        cmd_delete_dtl.Enabled = c And blnUser_Delete
    End If
End Sub

Private Sub fill_grid_user()
Dim v_flag_user As String
    
    If SSTab1.Tab = 0 Then
        If rsSuperUser.State Then rsSuperUser.Close
        SQL = "select * from m_user where user_level = 100 order by user_code"
        rsSuperUser.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_SuperUser.DataSource = rsSuperUser
    Else
        If rsUser.State Then rsUser.Close
        
        If TDBCombo_company.Text <> "" And TDBCombo_division.Text = "" And TDBCombo_section.Text = "" Then
            SQL = "select a.*,b.employee_name,b.level_code,b.nik " & _
                "from m_user a left join m_employee b on a.employee_code = b.employee_code " & _
                "where case when a.flag_user = 0 then a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '' AND a.division_code = '' AND " & _
                        "a.user_level <> 100 " & _
                        "else a.user_level <> 100 end " & _
                "order by a.user_name"
        ElseIf TDBCombo_company.Text <> "" And TDBCombo_division.Text <> "" And TDBCombo_section.Text = "" Then
            SQL = "select a.*,b.employee_name,b.level_code,b.nik " & _
                "from m_user a left join m_employee b on a.employee_code = b.employee_code " & _
                "where case when a.flag_user = 0 then a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '" & TDBCombo_division.Text & "' AND a.division_code = '' AND " & _
                        "a.user_level <> 100 " & _
                        "else a.user_level <> 100 end " & _
                "order by a.user_name"
        ElseIf TDBCombo_company.Text <> "" And TDBCombo_division.Text <> "" And TDBCombo_section.Text <> "" Then
            SQL = "select a.*,b.employee_name,b.level_code,b.nik " & _
                "from m_user a left join m_employee b on a.employee_code = b.employee_code " & _
                "where case when a.flag_user = 0 then a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '" & TDBCombo_division.Text & "' AND a.division_code = '" & TDBCombo_section.Text & "' AND " & _
                        "a.user_level <> 100 " & _
                        "else a.user_level <> 100 end " & _
                "order by a.user_name"
        End If
        rsUser.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_User.DataSource = rsUser
    End If
End Sub

Private Sub ShowWindowEntryUser(ByVal i As Boolean)
    If i = True Then ' Max
        If SSTab1.Tab = 0 Then
            fra_entry_super_user.Visible = True
        Else
            fra_entry_user.Visible = True
        End If
    ElseIf i = False Then ' Min
        If SSTab1.Tab = 0 Then
            fra_entry_super_user.Visible = False
        Else
            fra_entry_user.Visible = False
        End If
    End If
End Sub

Private Sub cmd_copy_cancel_Click()
    fraCopyDtl.Visible = False
End Sub

Private Sub cmd_paste_cancel_Click()
    fraPasteDtl.Visible = False
End Sub

Private Sub cmd_copy_ok_Click()
Dim i As Integer
Dim namaFile As String

On Error GoTo Err
    nmFile = App.Path & "\template\" & txt_name & ".ssd"
    namaFile = Replace(App.Path, "\", "\\") & "\\template\\" & txt_name & ".ssd"
    
    If txt_name.Text = "" Then
        MsgBox "Template Name Must Be Filled!", vbExclamation, headerMSG
        Exit Sub
    End If
    
    If checkFile(nmFile) Then
        i = MsgBox("Template has been existed ! " & _
            "Overwrite File ?", vbOKCancel, headerMSG)
        If i <> vbOK Then
            Exit Sub
        End If
        conLite.OpenDB nmFile
        
        SQL = "DELETE FROM m_template WHERE file_name = '" & txt_name & "'"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_template (file_name,file_location) " & _
                "VALUES( '" & txt_name & "','" & namaFile & "')"
        CnG.Execute SQL
    Else
        SQL = "INSERT INTO m_template (file_name,file_location) " & _
                "VALUES( '" & txt_name & "','" & namaFile & "')"
        CnG.Execute SQL
    
        conLite.CreateNewDB nmFile
    End If

    Call exportDataUserDetail
    Call exportDataUserAccess
    
    MsgBox "Copy Template Successfully!", vbInformation, headerMSG
    fraCopyDtl.Visible = False
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmd_paste_ok_Click()
Dim rsFileLoc As New ADODB.Recordset

On Error GoTo Err
    If cboFileName.Text = "" Then
        MsgBox "Template not found !", vbExclamation, headerMSG
        Exit Sub
    End If

    SQL = "SELECT * FROM m_template WHERE file_name = '" & Trim(cboFileName.Text) & "'"
    rsFileLoc.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsFileLoc.RecordCount > 0 Then
        nmFile = rsFileLoc!file_location
    End If
    rsFileLoc.Close

    Call importDataUserDetail
    Call importDataUserAccess

    Call fill_grid_user_form

    MsgBox "Paste Template Successfully!", vbInformation, headerMSG
    fraPasteDtl.Visible = False
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmd_delete_dtl_Click()
    TDBGrid_Form.Delete
    'TDBGrid_Level.Delete
End Sub

Private Sub cmd_save_dtl_Click()
'    TDBGrid_form.Update
'    TDBGrid_Level.Update
    rsAccessLevel.Update
    rsForm.Update

    MsgBox "Save Successfully!", vbInformation, headerMSG
End Sub

Private Sub delete_data()
Dim i As Integer

On Error GoTo Err
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_SuperUser.ApproxCount > 0 And TDBGrid_SuperUser.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid_SuperUser.Columns("user_name").Value & "' ?", vbOKCancel, headerMSG)
        
        If i = vbOK Then
            
            CnG.Execute "delete from m_user where user_code = '" _
            & TDBGrid_SuperUser.Columns("user_code").Value & "'"
        End If
    Else
        If Not (TDBGrid_User.ApproxCount > 0 And TDBGrid_User.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Are you sure want to delete data '" _
            & TDBGrid_User.Columns("user_name").Value & "' ?", vbOKCancel, headerMSG)
        
        If i = vbOK Then
            CnG.Execute "delete from t_user where level_code = '" _
            & TDBGrid_User.Columns("user_code").Value & "'"
            
            CnG.Execute "delete from t_user_access_level where level_code = '" _
            & TDBGrid_User.Columns("user_code").Value & "'"
            
            CnG.Execute "delete from m_user where user_code = '" _
            & TDBGrid_User.Columns("user_code").Value & "'"
        End If
    End If
    
    Call fill_grid_user
    
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()

    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_SuperUser.ApproxCount > 0 And TDBGrid_SuperUser.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        If cmdEditTop.Caption = "&Edit" Then
            cmdEditTop.Caption = "&Cancel"
            Call EnableButtonEntryUser(False, True, False, True)
            Call ShowWindowEntryUser(True)
            
            fra_entry_super_user.Caption = "Edit"
            txt_nama_super_user = TDBGrid_SuperUser.Columns("user_name").Value
            txt_pwd_super_user = DecryptINI(Trim(TDBGrid_SuperUser.Columns("user_pass").Value), pEncryptionPassword)
            txt_full_name_super_user = TDBGrid_SuperUser.Columns("full_name").Value
            txt_nama_super_user.SelStart = 0
            txt_nama_super_user.SelLength = Len(Trim(txt_nama_super_user))
            txt_nama_super_user.SetFocus
        Else
            cmdEditTop.Caption = "&Edit"
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
        End If
    Else
        If Not (TDBGrid_User.ApproxCount > 0 And TDBGrid_User.Bookmark > 0) Then
            MsgBox "No Data selected!", vbInformation, headerMSG
            Exit Sub
        End If
        
        If cmdEditUser.Caption = "&Edit" Then
            cmdEditUser.Caption = "&Cancel"
            Call EnableButtonEntryUser(False, True, False, True)
            Call ShowWindowEntryUser(True)
            fra_entry_user.Caption = "Edit"
            
            txt_user_code = TDBGrid_User.Columns("user_code").Value
            txt_NamaUser = TDBGrid_User.Columns("user_name").Value
            txt_PasswordUser = DecryptINI(Trim(TDBGrid_User.Columns("user_pass").Value), pEncryptionPassword)
            txtkd_employee.Text = IIf(IsNull(TDBGrid_User.Columns("employee_code").Value), "", TDBGrid_User.Columns("employee_code").Value)
            txt_nik.Text = IIf(IsNull(TDBGrid_User.Columns("employee_code").Value), "", TDBGrid_User.Columns("employee_code").Value)
            txt_nmemployee.Text = IIf(IsNull(TDBGrid_User.Columns("employee_name").Value), "", TDBGrid_User.Columns("employee_name").Value)
            chk_user_for.Value = IIf(IsNull(TDBGrid_User.Columns("flag_user").Value), 0, TDBGrid_User.Columns("flag_user").Value)
            
            txt_NamaUser.SelStart = 0
            txt_NamaUser.SelLength = Len(Trim(txt_NamaUser))
            txt_NamaUser.SetFocus
'            Call EnabledOptionUser(False)
        Else
            cmdEditUser.Caption = "&Edit"
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
'            Call EnabledOptionUser(True)
        End If
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function CekValidateDataUser() As Boolean
    If SSTab1.Tab = 0 Then
        If Not Trim(txt_nama_super_user) = "" And Not Trim(txt_pwd_super_user) = "" Then
            CekValidateDataUser = True
        Else
            CekValidateDataUser = False
        End If
    Else
        If Not Trim(txt_NamaUser) = "" And Not Trim(txt_PasswordUser) = "" Then
            CekValidateDataUser = True
        Else
            CekValidateDataUser = False
        End If
    End If
End Function

Private Function CekDuplicateNameUser() As Boolean
Dim rs1 As New ADODB.Recordset
    SQL = "select count(user_name) as JumlahRec from m_user " & _
            "where company_code = '" & TDBCombo_company.Text & "' AND " & _
            "user_name = '" & Replace$(Trim$(txt_NamaUser), Chr(39), Chr(96)) & "'"
    rs1.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs1!JumlahRec > 0 Then
        CekDuplicateNameUser = True
    Else
        CekDuplicateNameUser = False
    End If
    rs1.Close
End Function

Private Function CekDuplicateKodeUser() As Boolean
Dim rs1 As New ADODB.Recordset
    SQL = "select count(user_code) as JumlahRec from m_user " & _
            "where user_code = '" & Replace$(Trim$(txt_user_code), Chr(39), Chr(96)) & _
            "'"
    rs1.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs1!JumlahRec > 0 Then
        CekDuplicateKodeUser = True
    Else
        CekDuplicateKodeUser = False
    End If
    rs1.Close
End Function

Private Sub cmdGenerate_Click()
Dim str_user_code As String, int_level As Integer

    If Not (TDBGrid_User.ApproxCount > 0 And TDBGrid_User.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If
        
    str_user_code = TDBGrid_User.Columns("user_code").Value
    int_level = TDBGrid_User.Columns("level_code").Value
    Call generate_form(str_user_code, int_level)
    Call fill_grid_user_form
End Sub

Private Sub generate_form(ByVal str_user_code As String, ByVal int_level As Integer)

'    SQL = "Insert into t_user (level_code, sub_menu_code, sub_menu_name, menu_code, menu_name, " & _
'                "form_name, form_title, allow_read, allow_add, allow_edit, allow_delete, allow_post, allow_print) " & _
'            "Select '" & str_user_code & "', sub_menu_code, sub_menu_name, menu_code, menu_name, form_name, form_title, " & _
'                "1,1,1,1,1,1 " & _
'            "from m_sub_menu where (sub_menu_code <> 'M02-01' AND sub_menu_code <> 'M02-02' AND sub_menu_code <> 'M01-01' AND sub_menu_code <> 'M02-021' AND sub_menu_code <> 'M02-03') " & _
'                "AND ifnull(form_name,'')<>'' and sub_menu_code not in " & _
'                "(select sub_menu_code from t_user where level_code = '" & str_user_code & "')"
'    CnG.Execute SQL
    
    SQL = "Insert into t_user (level_code, sub_menu_code, sub_menu_name, menu_code, menu_name, " & _
                "form_name, form_title, allow_read, allow_add, allow_edit, allow_delete, allow_post, allow_print) " & _
            "Select '" & str_user_code & "', sub_menu_code, sub_menu_name, menu_code, menu_name, form_name, form_title, " & _
                "1,1,1,1,1,1 " & _
            "from m_sub_menu where (sub_menu_code <> 'M03-01' AND sub_menu_code <> 'M03-03') " & _
                "AND ifnull(form_name,'')<>'' and sub_menu_code not in " & _
                "(select sub_menu_code from t_user where level_code = '" & str_user_code & "')"
    CnG.Execute SQL
    
    SQL = "Insert into t_user_access_level (level_code, access_level_code, level_name, allow_access) " & _
            "Select '" & str_user_code & "', level_code, level_name, 0 " & _
            "from m_level where level_code not in " & _
                "(select access_level_code from t_user_access_level where level_code = '" & str_user_code & "')"
    CnG.Execute SQL

End Sub

Private Sub fill_grid_user_form()
On Error Resume Next
    If rsForm.State Then rsForm.Close
    SQL = "select * from t_user where level_code = '" & _
            TDBGrid_User.Columns("user_code").Value & _
            "' order by menu_code asc, sub_menu_code asc"
    rsForm.Open SQL, CnG, adOpenDynamic, adLockOptimistic
    
    Set TDBGrid_Form.DataSource = rsForm
    
    If rsAccessLevel.State Then rsAccessLevel.Close
    SQL = "select * from t_user_access_level where level_code = '" & _
            TDBGrid_User.Columns("user_code").Value & _
            "' order by access_level_code asc"
    rsAccessLevel.Open SQL, CnG, adOpenDynamic, adLockOptimistic
    
    Set TDBGrid_Level.DataSource = rsAccessLevel

End Sub
Private Sub new_data()
    If SSTab1.Tab = 0 Then
        If cmdNewTop.Caption = "&New" Then
            cmdNewTop.Caption = "&Cancel"
            fra_entry_super_user.Visible = True
            fra_entry_super_user.Caption = "Entry"
            
            txt_nama_super_user.SetFocus
            txt_nama_super_user = ""
            txt_full_name_super_user.Text = ""
            txt_pwd_super_user = ""
            
            cmdEditTop.Enabled = False
            cmdSaveTop.Enabled = True
            cmdDeleteTop.Enabled = False
        Else
            cmdNewTop.Caption = "&New"
            fra_entry_super_user.Visible = False
            txt_nama_super_user = ""
            txt_full_name_super_user = ""
            txt_pwd_super_user = ""
            
            cmdEditTop.Enabled = True
            cmdSaveTop.Enabled = False
            cmdDeleteTop.Enabled = True
        End If
    Else
        If cmdNewUser.Caption = "&New" Then
            cmdNewUser.Caption = "&Cancel"
            Call EnableButtonEntryUser(True, False, False, True)
            Call ShowWindowEntryUser(True)
            fra_entry_user.Caption = "Entry"
            
            txt_user_code = ""
            txt_NamaUser = ""
            txt_PasswordUser = ""
            txtkd_employee.Text = ""
            txt_nmemployee.Text = ""
            txt_nik.Text = ""
            txt_NamaUser.SetFocus
            
            chk_user_for.Value = 0
            
            Set TDBGrid_Form.DataSource = Nothing
            Set TDBGrid_Level.DataSource = Nothing
            
'            Call EnabledOptionUser(False)
        Else
            cmdNewUser.Caption = "&New"
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
'            Call EnabledOptionUser(True)
            
            If TDBGrid_User.Bookmark > 0 Then
            'Set GridEXForm.ADORecordset = Nothing
                Set TDBGrid_Form.DataSource = Nothing
                Set TDBGrid_Level.DataSource = Nothing
                Exit Sub
            End If
            Call fill_grid_user_form
        End If
    End If
End Sub

Private Sub simpan_data()
Dim rstot As New ADODB.Recordset
Dim int_level As Integer
Dim vTot As Integer
Dim clsFunc As New clsFunction

'On Error GoTo Err
    If SSTab1.Tab = 0 Then
        If fra_entry_super_user.Caption = "Entry" Then
            If CekValidateDataUser = False Then
                MsgBox "Data is not valid", vbCritical, headerMSG
                Exit Sub
            End If
            
            SQL = "select max(user_code) tot from m_user"
            rstot.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                vTot = IIf(IsNull(rstot!tot), 0, rstot!tot)
            rstot.Close
            
            vTot = vTot + 1
            
            SQL = "insert into m_user (user_code, user_name, user_pass, user_level,full_name) " _
                & "values('" & vTot & "', '" & Replace$(Trim$(txt_nama_super_user), Chr(39), Chr(96)) _
                & "','" & EncryptINI(Replace(Trim(txt_pwd_super_user.Text), Chr(39), Chr(96)), pEncryptionPassword) & "', " _
                & "100,'" & Trim(txt_full_name_super_user.Text) & "')"
            CnG.Execute SQL
            
            Call fill_grid_user
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
            cmdNewTop.Caption = "&New"
'            Call EnabledOptionTop(True)
            
        ElseIf fra_entry_super_user.Caption = "Edit" Then
            If CekValidateDataUser = False Then
                MsgBox "Data is not valid", vbCritical, headerMSG
                Exit Sub
            End If
            
            SQL = "update m_user set user_name = '" & Replace$(Trim$(txt_nama_super_user), Chr(39), Chr(96)) & _
                    "', full_name = '" & Trim(txt_full_name_super_user.Text) & _
                    "', user_pass = '" & EncryptINI(Replace(Trim(txt_pwd_super_user.Text), Chr(39), Chr(96)), pEncryptionPassword) & _
                    "' where user_code = '" & TDBGrid_SuperUser.Columns("user_code").Value & "'"
            CnG.Execute SQL
            
            Call fill_grid_user
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
            cmdEditTop.Caption = "&Edit"
'            Call EnabledOptionTop(True)
        End If
    Else
        If clsFunc.isEmployeeExist(txtkd_employee.Text) = False Then
            MsgBox "Invalid Employee Code...!!", vbExclamation, "Error"
            Exit Sub
        End If
        
        If fra_entry_user.Caption = "Entry" Then
            If CekValidateDataUser = False Then
                MsgBox "Data is not valid", vbCritical, headerMSG
                Exit Sub
            End If
            If CekDuplicateNameUser = True Then
                MsgBox "This User Name is Already Exist...", vbCritical, headerMSG
                Exit Sub
            End If
            If CekDuplicateKodeUser = True Then
                MsgBox "This User Code is Already Exist...", vbCritical, headerMSG
                Exit Sub
            End If
            
            SQL = "select max(user_code) tot from m_user"
            rstot.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                vTot = IIf(IsNull(rstot!tot), 0, rstot!tot)
            rstot.Close
            
            CnG.BeginTrans
            
            vTot = vTot + 1
            SQL = "INSERT INTO m_user(user_code,user_name,user_pass,user_level,company_code," & _
                    "employee_code,flag_user,department_code,division_code) " & _
                  "VALUES( " & _
                    "'" & vTot & "','" & Trim(txt_NamaUser) & "'," & _
                    "'" & Replace(EncryptINI(Trim(txt_PasswordUser.Text), pEncryptionPassword), "'", "''") & "'," & _
                    "0,'" & TDBCombo_company.Columns("company_code").Value & "'," & _
                    "'" & txtkd_employee.Text & "','" & chk_user_for.Value & "'," & _
                    "'" & TDBCombo_division.Text & "','" & TDBCombo_section.Text & "')"
            CnG.Execute SQL
            
            clsFunc.InsertLog ("Insert hakakses level : " & txt_user_code.Text)
            
            CnG.CommitTrans
            
            Call fill_grid_user
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
            cmdNewUser.Caption = "&New"
'            Call EnabledOptionUser(True)
            
        ElseIf fra_entry_user.Caption = "Edit" Then
            If CekValidateDataUser = False Then
                MsgBox "Editing Data Not Valid", vbCritical, "Request validate data"
                Exit Sub
            End If
            
            If Not Trim(txt_NamaUser) = TDBGrid_User.Columns("user_name").Value _
            And CekDuplicateNameUser = True Then
                MsgBox "Data found!", vbCritical, headerMSG
                Exit Sub
            End If
            
            CnG.BeginTrans
            SQL = "UPDATE m_user Set " _
                & "user_code = '" & Trim(txt_user_code.Text) & "'," _
                & "user_name = '" & Trim(txt_NamaUser.Text) & "'," _
                & "user_pass = '" & Replace(EncryptINI(Trim(txt_PasswordUser.Text), pEncryptionPassword), "'", "''") & "'," _
                & "company_code = '" & TDBCombo_company.Columns("company_code").Value & "'," _
                & "employee_code = '" & txtkd_employee.Text & "'," _
                & "flag_user = '" & chk_user_for.Value & "'," _
                & "department_code = '" & TDBCombo_division.Text & "'," _
                & "division_code = '" & TDBCombo_section.Text & "' " _
                & "WHERE user_code = '" & TDBGrid_User.Columns("user_code").Value & "'"
            CnG.Execute SQL
            
            clsFunc.InsertLog ("Edit hakakses level : " & txt_user_code.Text)
            CnG.CommitTrans
            
            Call fill_grid_user
            Call EnableButtonEntryUser(True, True, True, False)
            Call ShowWindowEntryUser(False)
            cmdEditUser.Caption = "&Edit"
'            Call EnabledOptionUser(True)
        End If
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Load()
    Call fill_grid_user
    Call createGridKar
    Call load_data_user_access(Me)
    oClause = ""
    
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    
    If LOGIN_LEVEL = 100 Then
        SSTab1.Tab = 0
    Else
        SSTab1.Tab = 1
        Call SSTab1_Click(0)
        SSTabUser.Enabled = False
        Set TDBGrid_SuperUser.DataSource = Nothing
    End If
End Sub

Public Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select company_code, company_name from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_company.RowSource = rsCompany
End Sub

Public Sub load_data_department()
    If rsDivision.State Then rsDivision.Close
    SQL = "select department_code, department_name from m_department " & _
            "where company_code = '" & TDBCombo_company.Text & "' order by company_code"
    rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_division.RowSource = rsDivision
End Sub

Public Sub load_data_division()
    If rsSection.State Then rsSection.Close
    SQL = "select division_code, division_name from m_division " & _
            "where company_code = '" & TDBCombo_company.Text & "' " & _
            "and department_code = '" & TDBCombo_division.Text & "' order by company_code"
    rsSection.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_section.RowSource = rsSection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set conLite = Nothing
    Set rss = Nothing
    Set rsLite = Nothing
    Set frm_mst_user = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        If LOGIN_LEVEL = 100 Then Call fill_grid_user
        txt_company_name.Text = ""
    Else
        Call load_data_company
        timer1.Enabled = True
    End If
End Sub

'Private Sub EnabledOptionUser(ByVal i As Boolean)
''fraOption.Enabled = i
'GridEXUser.Enabled = i
'End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_department
        Call fill_grid_user
'        If TDBCombo_userlevel.Text <> "" Then
'            fill_grid_user
'        Else
'            Call load_data_userlevel
'        End If
    End If
End Sub

Private Sub TDBCombo_division_ItemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("department_code").Value
        txt_division = TDBCombo_division.Columns("department_name").Value
        
        Call load_data_division
        Call fill_grid_user
    End If
End Sub

Private Sub TDBCombo_section_ItemChange()
    If TDBCombo_section.ApproxCount > 0 Then
        TDBCombo_section.Text = TDBCombo_section.Columns("division_code").Value
        txt_section = TDBCombo_section.Columns("division_name").Value
        
        Call fill_grid_user
    End If
End Sub

Private Sub TDBGrid_User_DblClick()
    Call edit_data
End Sub

Private Sub TDBGrid_User_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not TDBGrid_User.Bookmark > 0 Then
        Set TDBGrid_Form.DataSource = Nothing
        Set TDBGrid_Level.DataSource = Nothing
        Exit Sub
    End If
    Call fill_grid_user_form
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
    If LOGIN_LEVEL = 100 Then
        TDBCombo_company.Locked = False
    Else
        TDBCombo_company.Locked = True
    End If
End Sub


Private Sub cmdNewTop_Click()
    Call new_data
End Sub

Private Sub cmdSaveTop_Click()
    Call simpan_data
End Sub

Private Sub cmdEditTop_Click()
    Call edit_data
End Sub

Private Sub cmdDeleteTop_Click()
    Call delete_data
End Sub


Private Sub cmdNewUser_Click()
    Call new_data
End Sub

Private Sub cmdSimpanUser_Click()
    Call simpan_data
End Sub

Private Sub cmdEditUser_Click()
    Call edit_data
End Sub

Private Sub cmdDeleteUser_Click()
    Call delete_data
End Sub


Private Sub cmdCopyDtl_Click()
    fraCopyDtl.Visible = True
    txt_name.Text = ""
    
'Dim i As Integer
'    CommonDialog1.ShowSave
'
'    If Right(CommonDialog1.FileName, 4) = ".ssd" Then
'        nmFile = CommonDialog1.FileName
'    Else
'        nmFile = CommonDialog1.FileName & ".ssd"
'    End If
'
'    If checkFile(nmFile) Then
'        i = MsgBox("File has been existed ! " & _
'            "Overwrite File ?", vbOKCancel, headerMSG)
'        If i <> vbOK Then
'            Exit Sub
'        End If
'        conLite.OpenDB nmFile
'    Else
'        conLite.CreateNewDB nmFile
'    End If
'
'    Call exportDataUserDetail
'    Call exportDataUserAccess
'
'    MsgBox "Copy Template Successfully!", vbInformation, headerMSG
End Sub

Private Sub cmdPasteDtl_Click()
    fraPasteDtl.Visible = True
    cboFileName.Clear
    
    SQL = "SELECT file_name from m_template order by file_name"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
    
        If Not rs.EOF Then
            rs.MoveFirst
            While Not rs.EOF
                cboFileName.AddItem rs!file_name
            rs.MoveNext
            Wend
        End If
    
    End If
    rs.Close
    
'Dim i As Integer
'    CommonDialog1.ShowOpen
'
'    If Not Right(CommonDialog1.FileName, 4) = ".ssd" Then
'        MsgBox "Invalid file name ! Filename format must be in <filename>.ssd !"
'        Exit Sub
'    End If
'
'    If Not checkFile(CommonDialog1.FileName) Then
'        MsgBox "File not found !"
'        Exit Sub
'    End If
'
'    Call importDataUserDetail
'    Call importDataUserAccess
'
'    Call fill_grid_user_form
'
'    MsgBox "Paste Template Successfully!", vbInformation, headerMSG
End Sub

Private Sub exportDataUserDetail()
Dim cmd As cCommand
'Dim rss As New cRecordset

On Error GoTo Err
    conLite.Execute "CREATE TABLE IF NOT EXISTS t_user (" & _
                            "level_code TEXT(30) NOT NULL,sub_menu_code TEXT(30) NOT NULL," & _
                            "sub_menu_name TEXT(50) DEFAULT NULL,menu_code TEXT(50) DEFAULT NULL," & _
                            "menu_name TEXT(50) DEFAULT NULL,form_name TEXT(50) DEFAULT NULL," & _
                            "form_title TEXT(50) DEFAULT NULL,allow_read INTEGER DEFAULT NULL," & _
                            "allow_add INTEGER DEFAULT NULL,allow_edit INTEGER DEFAULT NULL," & _
                            "allow_delete INTEGER DEFAULT NULL,allow_post INTEGER DEFAULT NULL," & _
                            "allow_print INTEGER DEFAULT NULL,PRIMARY KEY(level_code,sub_menu_code)); "
                        
    SQL = "INSERT INTO t_user " & _
                "(level_code,sub_menu_code,sub_menu_name,menu_code," & _
                "menu_name,form_name,form_title,allow_read," & _
                "allow_add,allow_edit,allow_delete,allow_post," & _
                "allow_print) " & _
                "Values " & _
                "(?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                   
    Set cmd = conLite.CreateCommand(SQL)
    
    SQL = "SELECT level_code,sub_menu_code,sub_menu_name,menu_code," & _
                "menu_name,form_name,form_title,allow_read," & _
                "allow_add,allow_edit,allow_delete,allow_post," & _
                "allow_print " & _
             "FROM t_user " & _
             "WHERE level_code = '" & TDBGrid_User.Columns("user_code").Value & "'"
             
    Set rsAtt = New ADODB.Recordset

    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    conLite.BeginTrans
    conLite.Execute ("DELETE FROM t_user")
    
    If Not rsAtt.EOF Then
        rsAtt.MoveFirst
        While Not rsAtt.EOF
            DoEvents
            cmd.SetText 1, rsAtt!level_code
            cmd.SetText 2, rsAtt!sub_menu_code
            cmd.SetText 3, IIf(IsNull(rsAtt!sub_menu_name), "", rsAtt!sub_menu_name)
            cmd.SetText 4, IIf(IsNull(rsAtt!menu_code), "", rsAtt!menu_code)
            cmd.SetText 5, IIf(IsNull(rsAtt!menu_name), "", rsAtt!menu_name)
            cmd.SetText 6, IIf(IsNull(rsAtt!form_name), "", rsAtt!form_name)
            cmd.SetText 7, IIf(IsNull(rsAtt!form_title), "", rsAtt!form_title)
            cmd.SetInt32 8, IIf(IsNull(rsAtt!allow_read), 0, rsAtt!allow_read)
            cmd.SetInt32 9, IIf(IsNull(rsAtt!allow_add), 0, rsAtt!allow_add)
            cmd.SetInt32 10, IIf(IsNull(rsAtt!allow_edit), 0, rsAtt!allow_edit)
            cmd.SetInt32 11, IIf(IsNull(rsAtt!allow_delete), 0, rsAtt!allow_delete)
            cmd.SetInt32 12, IIf(IsNull(rsAtt!allow_post), 0, rsAtt!allow_post)
            cmd.SetInt32 13, IIf(IsNull(rsAtt!allow_print), 0, rsAtt!allow_print)
            cmd.Execute
        rsAtt.MoveNext
        Wend
    End If
    rsAtt.Close
    conLite.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub exportDataUserAccess()
Dim cmd As cCommand
'Dim rss As New cRecordset

On Error GoTo Err
    conLite.Execute "CREATE TABLE IF NOT EXISTS t_user_access_level (" & _
                            "level_code TEXT(30) NOT NULL,access_level_code INTEGER NOT NULL," & _
                            "level_name TEXT(50) DEFAULT NULL,allow_access INTEGER DEFAULT NULL," & _
                            "PRIMARY KEY(level_code,access_level_code)); "
                        
    SQL = "INSERT INTO t_user_access_level " & _
                "(level_code,access_level_code,level_name,allow_access) " & _
                "Values " & _
                "(?,?,?,?)"
                                   
    Set cmd = conLite.CreateCommand(SQL)
    
    SQL = "SELECT level_code,access_level_code,level_name,allow_access " & _
             "FROM t_user_access_level " & _
             "WHERE level_code = '" & TDBGrid_User.Columns("user_code").Value & "'"
             
    Set rsAtt = New ADODB.Recordset

    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    conLite.BeginTrans
    conLite.Execute ("DELETE FROM t_user_access_level")
    
    If Not rsAtt.EOF Then
        rsAtt.MoveFirst
        While Not rsAtt.EOF
            
            DoEvents
            cmd.SetText 1, rsAtt!level_code
            cmd.SetInt32 2, IIf(IsNull(rsAtt!access_level_code), 0, rsAtt!access_level_code)
            cmd.SetText 3, IIf(IsNull(rsAtt!level_name), "", rsAtt!level_name)
            cmd.SetInt32 4, IIf(IsNull(rsAtt!allow_access), 0, rsAtt!allow_access)
            cmd.Execute
            
        rsAtt.MoveNext
        Wend
    End If
    rsAtt.Close
    conLite.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Function check_exist_new_t_user(lvlCode As String, sub_menu_code As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_exist_new_t_user = False
    
    str_sql = "select count(level_code) as rec_count from t_user where level_code = '" & _
                Replace$(Trim$(lvlCode), Chr$(39), Chr$(96)) & "' " & _
                "AND sub_menu_code = '" & sub_menu_code & "'"
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_exist_new_t_user = True
        Exit Function
    End If
End Function

Private Function check_exist_new_t_user_access(lvlCode As String, access_level_code As String) As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_exist_new_t_user_access = False
    
    str_sql = "select count(level_code) as rec_count from t_user_access_level where level_code = '" & _
                Replace$(Trim$(lvlCode), Chr$(39), Chr$(96)) & "' " & _
                "AND access_level_code = '" & access_level_code & "'"
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_exist_new_t_user_access = True
        Exit Function
    End If
End Function

Private Sub importDataUserDetail()
Dim cmd As cCommand
Dim aa, bb, i As Integer
'Dim rss As New cRecordset

On Error GoTo Err
    conLite.OpenDB nmFile
    
    If Not SelectTable("t_user") Then
        MsgBox "Cannot find data on selected file.. Please check your file again !"
        Exit Sub
    Else
        SQL = "SELECT * FROM t_user"
        rss.OpenRecordset SQL, conLite, True
        Set rsLite = conLite.OpenRecordset(SQL, True)
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
        rs.Open "select * from t_user", CnG, adOpenKeyset, adLockOptimistic
            
        For aa = 0 To rss.RecordCount - 1
        
            DoEvents
            If Not check_exist_new_t_user(TDBGrid_User.Columns("user_code").Value, rss!sub_menu_code) Then
                i = i + 1
                        
                With rs
                    .AddNew
                    '-----------------------
                    .Fields("level_code").Value = TDBGrid_User.Columns("user_code").Value
                    .Fields("sub_menu_code").Value = rss!sub_menu_code
                    .Fields("sub_menu_name").Value = rss!sub_menu_name
                    .Fields("menu_code").Value = rss!menu_code
                    .Fields("menu_name").Value = rss!menu_name
                    .Fields("form_name").Value = rss!form_name
                    .Fields("form_title").Value = rss!form_title
                    .Fields("allow_read").Value = rss!allow_read
                    .Fields("allow_add").Value = rss!allow_add
                    .Fields("allow_edit").Value = rss!allow_edit
                    .Fields("allow_delete").Value = rss!allow_delete
                    .Fields("allow_post").Value = rss!allow_post
                    .Fields("allow_print").Value = rss!allow_print
                    '-----------------------
                    .Update
                End With
            Else
                SQL = "UPDATE t_user SET level_code = '" & TDBGrid_User.Columns("user_code").Value & "'," _
                    & "sub_menu_code = '" & rss!sub_menu_code & "'," _
                    & "sub_menu_name = '" & Replace(rss!sub_menu_name, "'", "''") & "'," _
                    & "menu_code = '" & rss!menu_code & "'," _
                    & "menu_name = '" & Replace(rss!menu_name, "'", "''") & "'," _
                    & "form_name = '" & rss!form_name & "'," _
                    & "form_title = '" & rss!form_title & "'," _
                    & "allow_read = '" & rss!allow_read & "'," _
                    & "allow_add = '" & rss!allow_add & "'," _
                    & "allow_edit = '" & rss!allow_edit & "'," _
                    & "allow_delete = '" & rss!allow_delete & "'," _
                    & "allow_post = '" & rss!allow_post & "'," _
                    & "allow_print = '" & rss!allow_print & "' " _
                    & "WHERE level_code = '" & TDBGrid_User.Columns("user_code").Value & "' " _
                        & "AND sub_menu_code = '" & rss!sub_menu_code & "'"
                CnG.Execute SQL
            End If
            rss.MoveNext
        Next
        rs.Close
        CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub importDataUserAccess()
Dim cmd As cCommand
Dim aa, bb, i As Integer
'Dim rss As New cRecordset

On Error GoTo Err
    conLite.OpenDB nmFile
    
    If Not SelectTable("t_user_access_level") Then
        MsgBox "Cannot find data on selected file.. Please check your file again !"
        Exit Sub
    Else
        SQL = "SELECT * FROM t_user_access_level"
        rss.OpenRecordset SQL, conLite, True
        Set rsLite = conLite.OpenRecordset(SQL, True)
    End If
    
    If rss.RecordCount > 0 Then
    CnG.BeginTrans
    
    i = 0
    
    Set rs = New ADODB.Recordset
            
    If rs.State = 1 Then rs.Close
        rs.Open "select * from t_user_access_level", CnG, adOpenKeyset, adLockOptimistic
                
        For aa = 0 To rss.RecordCount - 1
        
            DoEvents
            If Not check_exist_new_t_user_access(TDBGrid_User.Columns("user_code").Value, rss!access_level_code) Then
                i = i + 1
                        
                With rs
                    .AddNew
                    '-----------------------
                    .Fields("level_code").Value = TDBGrid_User.Columns("user_code").Value
                    .Fields("access_level_code").Value = rss!access_level_code
                    .Fields("level_name").Value = rss!level_name
                    .Fields("allow_access").Value = rss!allow_access
                    '-----------------------
                    .Update
                End With
            Else
                SQL = "UPDATE t_user_access_level SET level_code = '" & TDBGrid_User.Columns("user_code").Value & "'," _
                    & "access_level_code = '" & rss!access_level_code & "'," _
                    & "level_name = '" & rss!level_name & "'," _
                    & "allow_access = '" & rss!allow_access & "' " _
                    & "WHERE level_code = '" & TDBGrid_User.Columns("user_code").Value & "' " _
                        & "AND access_level_code = '" & rss!access_level_code & "'"
                CnG.Execute SQL
            End If
            rss.MoveNext
        Next
        rs.Close
        CnG.CommitTrans
    Else
        aa = 0
    End If
    
    Set rss = Nothing
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Function checkFile(nmFile As String) As Boolean
    If Dir$(nmFile) <> "" Then
        checkFile = True
    Else
        checkFile = False
    End If
End Function

Public Function SelectTable(tableName As String) As Boolean
'Dim rss As New cRecordset
On Error GoTo Err
    SQL = "SELECT 1 FROM " & tableName

    rss.OpenRecordset SQL, conLite, True
    
    SelectTable = True
    Exit Function

Err:
SelectTable = False
End Function

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
        
        If LOGIN_LEVEL = 100 Then
            If TDBCombo_company.Text <> "" And TDBCombo_division.Text = "" And TDBCombo_section.Text = "" Then
                SQL = "select nik,employee_name,employee_code " & _
                    "from m_employee a " & _
                    "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "flag_active <> 0 " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_nik.Text & "%')"
            ElseIf TDBCombo_company.Text <> "" And TDBCombo_division.Text <> "" And TDBCombo_section.Text = "" Then
                SQL = "select nik,employee_name,employee_code " & _
                    "from m_employee a " & _
                    "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '" & TDBCombo_division.Text & "' AND flag_active <> 0 " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_nik.Text & "%')"
            ElseIf TDBCombo_company.Text <> "" And TDBCombo_division.Text <> "" And TDBCombo_section.Text <> "" Then
                SQL = "select nik,employee_name,employee_code " & _
                    "from m_employee a " & _
                    "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '" & TDBCombo_division.Text & "' AND a.division_code = '" & TDBCombo_section.Text & "' AND flag_active <> 0 " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_nik.Text & "%')"
            End If
        Else
            If TDBCombo_company.Text <> "" And TDBCombo_division.Text = "" And TDBCombo_section.Text = "" Then
                SQL = "select nik,employee_name,employee_code " & _
                    "from m_employee a " & _
                    "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "flag_active <> 0 " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
            ElseIf TDBCombo_company.Text <> "" And TDBCombo_division.Text <> "" And TDBCombo_section.Text = "" Then
                SQL = "select nik,employee_name,employee_code " & _
                    "from m_employee a " & _
                    "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '" & TDBCombo_division.Text & "' AND flag_active <> 0 " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
            ElseIf TDBCombo_company.Text <> "" And TDBCombo_division.Text <> "" And TDBCombo_section.Text <> "" Then
                SQL = "select nik,employee_name,employee_code " & _
                    "from m_employee a " & _
                    "where a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND " & _
                        "a.department_code = '" & TDBCombo_division.Text & "' AND a.division_code = '" & TDBCombo_section.Text & "' AND flag_active <> 0 " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                        "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
            End If
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
                txtkd_employee.Text = rs!employee_code
                txt_nmemployee.Text = rs!EMPLOYEE_NAME
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
            txt_nmemployee.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txtkd_employee.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
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
        txtkd_employee.Text = ""
        txt_nmemployee.Text = ""
    End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmd_periode_browse_employee_Click()
    isiGridKar (1)
End Sub


Private Sub txt_full_name_super_user_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
