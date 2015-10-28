VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frm_trans_shift 
   Caption         =   "SETTING SHIFT KARYAWAN"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_trans_shift.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6840
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown2 
      Height          =   1455
      Left            =   6600
      TabIndex        =   19
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "KODE SHIFT"
      Columns(0).DataField=   "kode_shift"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NAMA SHIFT"
      Columns(1).DataField=   "nama_shift"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   "kode_shift"
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   13160660
      ValueTranslate  =   0   'False
      ScrollTrack     =   -1  'True
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   255
      Left            =   7920
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      Calendar        =   "frm_trans_shift.frx":058A
      Caption         =   "frm_trans_shift.frx":068D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_trans_shift.frx":06F2
      Keys            =   "frm_trans_shift.frx":0710
      Spin            =   "frm_trans_shift.frx":076E
      AlignHorizontal =   2
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "11-07-2007"
      ValidateMode    =   0
      ValueVT         =   1549008903
      Value           =   39274
      CenturyMode     =   0
   End
   Begin MSAdodcLib.Adodc Adodc_karyawan 
      Height          =   375
      Left            =   1800
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc_karyawan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fra_setting 
      Caption         =   "DAFTAR SETTING SHIFT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   13935
      Begin VB.Frame Frame2 
         Caption         =   "DEFAULT SHIFT"
         Height          =   855
         Left            =   8640
         TabIndex        =   20
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txt_nama_shift 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   22
            Top             =   360
            Width           =   2415
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_shift 
            Height          =   495
            Left            =   360
            OleObjectBlob   =   "frm_trans_shift.frx":0796
            TabIndex        =   21
            Top             =   360
            Width           =   1575
         End
         Begin MSAdodcLib.Adodc Adodc_shift 
            Height          =   375
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc_shift"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "DEFAULT PERIODE"
         Height          =   855
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPicker_start 
            Height          =   315
            Left            =   480
            TabIndex        =   14
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   22806531
            CurrentDate     =   39270.5
         End
         Begin MSComCtl2.DTPicker DTPicker_end 
            Height          =   315
            Left            =   2280
            TabIndex        =   15
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   22806531
            CurrentDate     =   39270.5
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "S/D"
            Height          =   195
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Timer timer_combo 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   0
         Top             =   0
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_no_shift 
         Height          =   495
         Left            =   1440
         OleObjectBlob   =   "frm_trans_shift.frx":29B6
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc Adodc_no_shift 
         Height          =   375
         Left            =   1920
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txt_kode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   17
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NO URUT"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   675
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   9340
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO URUT"
      Columns(0).DataField=   "no_urut"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KODE KARYAWAN"
      Columns(1).DataField=   "kode_karyawan"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NAMA KARYAWAN"
      Columns(2).DataField=   "nama_karyawan"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "KODE SHIFT"
      Columns(3).DataField=   "kode_shift"
      Columns(3).DropDown=   "TDBDropDown2"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "NAMA SHIFT"
      Columns(4).DataField=   "nama_shift"
      Columns(4).NumberFormat=   "FormatText Event"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "BERLAKU"
      Columns(5).DataField=   "start"
      Columns(5).NumberFormat=   "FormatText Event"
      Columns(5).ExternalEditor=   "TDBDate1"
      Columns(5).ExternalEditor.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "FLAG"
      Columns(6).DataField=   "flag"
      Columns(6).NumberFormat=   "FormatText Event"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "BERAKHIR"
      Columns(7).DataField=   "end"
      Columns(7).NumberFormat=   "FormatText Event"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1693"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1614"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6138"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6059"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(17)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=3016"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2937"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(3).AutoDropDown=1"
      Splits(0)._ColumnProps(25)=   "Column(3).AutoCompletion=1"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=4683"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4604"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2937"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2858"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=2884"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2805"
      Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(6)._MinWidth=10"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2937"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2858"
      Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(48)=   "Column(7)._MinWidth=54215968"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "DAFTAR KARYAWAN"
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
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000A&"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=98,.parent=13,.bgcolor=&H80000013&"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=102,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=110,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=107,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=108,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=109,.parent=17"
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
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   7440
      Width           =   13935
      Begin VB.Timer timer_grid_detail 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7440
         Picture         =   "frm_trans_shift.frx":4F79
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_trans_shift.frx":5503
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_trans_shift.frx":5A8D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11520
         Picture         =   "frm_trans_shift.frx":6017
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_trans_shift.frx":65A1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Report"
         Height          =   645
         Left            =   8520
         Picture         =   "frm_trans_shift.frx":6B2B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_trans_shift.frx":70B5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_trans_shift.frx":763F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_trans_shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim int_no_urut As Integer
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim str_kode_rekening As String
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns



Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
check_validate_exist_new = False

SQL = "select count(no_urut) as jml from tm_shift_karyawan where no_urut = " & Int(txt_kode)
rs.Open SQL, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("jml").Value > 0 Then
    MsgBox "No Urut Sudah Ada !", vbCritical, headerMSG
    txt_kode = ""
    If txt_kode.Enabled = True Then txt_kode.SetFocus
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi no_urut
If Trim(txt_kode) = "" Then
    MsgBox "No Urut belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_kode.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi start
If IsNull(DTPicker_start.Value) = True Then
    MsgBox "Tanggal Start belum diisi", vbOKOnly + vbInformation, headerMSG
    DTPicker_start.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi end
If IsNull(DTPicker_end.Value) = True Then
    MsgBox "Tanggal End belum diisi", vbOKOnly + vbInformation, headerMSG
    DTPicker_end.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi shift
If cek_validate_tdbcombo(TDBCombo_shift) = False Then
    MsgBox "Shift belum dipilih", vbOKOnly + vbInformation, headerMSG
    TDBCombo_shift.SetFocus
    check_validate_new = False
    Exit Function
End If
End Function

Private Sub load_data_grid()
timer_grid_detail.Enabled = True
End Sub

Private Sub load_data_combo()
timer_combo.Enabled = True
End Sub

Private Sub load_data_dropdown()
Adodc2.RecordSource = "select * from m_shift order by kode_shift"
Adodc2.Refresh

TDBDropDown2.DataSource = Adodc2
End Sub


Private Sub cmd_refresh_Click()
Dim i As Integer, str_sql As String

If cek_validate_tdbcombo(TDBCombo_no_shift) = False Then
    MsgBox "Data No Urut belum dipilih", vbInformation, headerMSG
    TDBCombo_no_shift.SetFocus
    Exit Sub
End If

i = MsgBox("Anda yakin ingin mendaftarkan semua karyawan di No. Urut " _
& TDBCombo_no_shift.Columns("no_urut_").Value & "' ?", vbOKCancel, headerMSG)

If Not i = vbOK Then Exit Sub

str_sql = "select kode_karyawan, nama_karyawan from m_karyawan where kode_karyawan not in " _
& "(select kode_karyawan from td_shift_karyawan where no_urut = " _
& Int(TDBCombo_no_shift.Columns("no_urut_").Value) & ")"

Adodc3.RecordSource = str_sql
Adodc3.Refresh

If Not Adodc3.Recordset.RecordCount > 0 Then
    MsgBox "Semua karyawan sudah didaftarkan", vbInformation, headerMSG
    Exit Sub
End If

While Not Adodc3.Recordset.EOF
    With Adodc1.Recordset
        .AddNew
        .Fields("no_urut").Value = TDBCombo_no_shift.Columns("no_urut_").Value
        .Fields("kode_karyawan").Value = Adodc3.Recordset.Fields("kode_karyawan").Value
        .Fields("nama_karyawan").Value = Adodc3.Recordset.Fields("nama_karyawan").Value
        .Fields("kode_shift").Value = TDBCombo_shift.Columns("kode_shift").Value
        .Fields("nama_shift").Value = TDBCombo_shift.Columns("nama_shift").Value
        .Fields("start").Value = DTPicker_start.Value
        .Fields("flag").Value = 1
        .Fields("end").Value = DTPicker_end.Value
        .Fields("keterangan").Value = "-"
        .Update
    End With
    Adodc3.Recordset.MoveNext
Wend

MsgBox Adodc3.Recordset.RecordCount & " data sudah didaftarkan", vbInformation, headerMSG

Call load_data_combo
int_mode = 0
Call load_mode
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If TDBGrid1.ApproxCount > 0 Then
    i = MsgBox("Data karyawan sudah ada" & vbCr _
    & "Anda yakin ingin menghapus data No Urut berikut semua karyawan yang sudah didaftar ?" _
    , vbYesNo + vbQuestion, headerMSG)

    If Not i = vbYes Then Exit Sub
End If

'CnG.BeginTrans
CnG.Execute "delete from td_shift_karyawan where no_urut = " _
        & Adodc_no_shift.Recordset.Fields("no_urut_").Value
CnG.Execute "delete from tm_shift_karyawan where no_urut = " _
        & Adodc_no_shift.Recordset.Fields("no_urut_").Value
'CnG.CommitTrans

Call load_data_combo
int_mode = 0
Call load_mode
End Sub



Private Sub cmdEdit_Click()
If cek_validate_tdbcombo(TDBCombo_no_shift) = False Then
    MsgBox "Data No Urut belum dipilih", vbInformation, headerMSG
    TDBCombo_no_shift.SetFocus
    Exit Sub
End If

If TDBGrid1.ApproxCount > 0 Then
    MsgBox "Data Karyawan sudah ada", vbCritical, headerMSG
    Exit Sub
End If

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from tm_shift_karyawan where no_urut = " _
& Adodc_no_shift.Recordset.Fields("no_urut").Value, CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from tm_shift_karyawan where no_urut = -777", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdPosting_Click()
Dim rs As New ADODB.Recordset
Dim i As Integer

rs.Open "select count(kode_supplier) as jml_rec from m_supplier " _
& "where isnull(flag_posting,0)=0", CnG, adOpenStatic, adLockReadOnly

If rs.Fields("jml_rec").Value < 1 Then
    MsgBox "Semua Data Supplier sudah terposting", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Anda yakin ingin melakukan posting " & rs.Fields("jml_rec").Value _
& " data supplier ?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub

If rs.State = 1 Then rs.Close
rs.Open "select * from m_supplier " _
& "where isnull(flag_posting,0)=0 order by kode_supplier", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans
With rs
    While Not .EOF
        .Fields("flag_posting").Value = 1
        .Update
        
        .MoveNext
    Wend
End With
CnG.CommitTrans

MsgBox rs.RecordCount & " data Supplier sudah terposting", vbInformation, headerMSG
Call load_data_grid
int_mode = 0
Call load_mode
End Sub

Private Sub CmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
'CnG.BeginTrans

With rsBound
    .AddNew
    
    .Fields("no_urut").Value = Int(txt_kode)     ' key
    .Fields("start").Value = DTPicker_start.Value
    .Fields("end").Value = DTPicker_end.Value
    .Fields("kode_shift").Value = Adodc_shift.Recordset.Fields("kode_shift").Value
    .Fields("nama_shift").Value = Adodc_shift.Recordset.Fields("nama_shift").Value
    
    .Update
End With

'CnG.CommitTrans
End Sub

Private Sub edit_old_data()
'CnG.BeginTrans

With rsBound
    '.AddNew
    '.Fields("no_urut").Value = Int(txt_kode)     ' key
    .Fields("start").Value = DTPicker_start.Value
    .Fields("end").Value = DTPicker_end.Value
    .Fields("kode_shift").Value = Adodc_shift.Recordset.Fields("kode_shift").Value
    .Fields("nama_shift").Value = Adodc_shift.Recordset.Fields("nama_shift").Value
    
    .Update
End With

'CnG.CommitTrans
End Sub

Private Sub CmdSave_Click()
If int_mode = 1 Then
    If Not check_validate_new Then Exit Sub
    If check_validate_exist_new Then Exit Sub
    Call insert_new_data
ElseIf int_mode = 2 Then
    If Not check_validate_new Then Exit Sub
    Call edit_old_data
End If

Call load_data_combo
int_mode = 0
Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
CmdNew.Enabled = a And blnUser_Tambah
CmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Ubah
cmdDelete.Enabled = d And blnUser_Hapus
CmdCancel.Enabled = e

CmdPrint.Enabled = f
cmd_refresh.Enabled = g 'And blnUser_posting
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        Ctr.Text = ""
    ElseIf TypeOf Ctr Is TDBCombo Then
        Ctr.Text = ""
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
    ElseIf TypeOf Ctr Is TDBGrid Then
        Ctr.DataSource = Nothing
    End If
Next
End Sub

Private Sub set_enabled_control(ByVal i As Boolean)
Dim Ctr As Control
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

Public Sub set_edit_data()
txt_kode = TDBCombo_no_shift.Columns("no_urut_").Value
End Sub

Private Sub set_new_data()
DTPicker_start.Value = Now
DTPicker_end.Value = Now
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    txt_kode.Visible = True
    TDBCombo_no_shift.Visible = False
    TDBGrid1.Enabled = False
    DTPicker_start.Enabled = True
    DTPicker_end.Enabled = True
    TDBCombo_shift.Enabled = True
    
    If txt_kode.Enabled = True Then
        txt_kode.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    txt_kode.Visible = False
    TDBCombo_no_shift.Visible = True
    TDBCombo_no_shift.Enabled = True
    DTPicker_start.Enabled = False
    DTPicker_end.Enabled = False
    TDBCombo_shift.Enabled = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    TDBCombo_no_shift.Enabled = False
    TDBGrid1.Enabled = False
    DTPicker_start.Enabled = True
    DTPicker_end.Enabled = True
    TDBCombo_shift.Enabled = True
    Call set_edit_data
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

Private Sub Command1_Click()
MsgBox Chr$(40)
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc_no_shift.ConnectionString = strConn
Adodc_karyawan.ConnectionString = strConn
Adodc2.ConnectionString = strConn
Adodc_shift.ConnectionString = strConn
Adodc3.ConnectionString = strConn

Call load_data_shift
Call load_data_karyawan
Call load_data_combo
Call load_data_dropdown
int_mode = 0
Call LoadData_UserAccess(Me.Caption)
Call load_mode
End Sub

Private Function get_setting_rekening(ByVal int_no As Integer) As Boolean
Dim rs As New ADODB.Recordset

rs.Open "exec sp_get_setting_rekening " & int_no, CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    str_kode_rekening = rs.Fields("kode_rekening").Value
    get_setting_rekening = True
Else
    str_kode_rekening = ""
    get_setting_rekening = False
End If
End Function

Private Sub txtFax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtTelp1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtTelp2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Function get_inc_kode() As String
Dim rs As New ADODB.Recordset
Dim str_inc_kode As String

rs.Open "select max(kode_supplier) as curr_kode from m_supplier", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    If IsNull(rs.Fields("curr_kode").Value) = True Then
        str_inc_kode = "00001"
    Else
        str_inc_kode = rs.Fields("curr_kode").Value
        str_inc_kode = Right("0000" & Trim(Str(CLng(str_inc_kode) + 1)), 5)
    End If
End If

get_inc_kode = str_inc_kode
End Function


Private Sub txtTelp3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub clear_filter()
For Each Col In TDBGrid1.Columns
    Col.FilterText = ""
Next Col
Adodc1.Recordset.Filter = adFilterNone
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

Private Sub TDBCombo_no_shift_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBCombo_no_shift.Columns(ColIndex).Caption = "BERLAKU" Or _
TDBCombo_no_shift.Columns(ColIndex).Caption = "BERAKHIR" Then
    Value = Format(Value, "dd-mm-yyyy")
End If
End Sub

Private Sub view_data()
With Adodc_no_shift.Recordset

    DTPicker_start.Value = .Fields("start").Value
    DTPicker_end.Value = .Fields("end").Value
    int_no_urut = .Fields("no_urut").Value
    Call set_data_shift(.Fields("kode_shift").Value)
    Call load_data_grid
    
End With
End Sub

Private Sub set_data_shift(ByVal str_kode As String)
Adodc_shift.Recordset.MoveFirst
Adodc_shift.Recordset.Find ("kode_shift='" & str_kode & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_shift.Recordset.EOF = True Or Adodc_shift.Recordset.BOF = True) Then
    TDBCombo_shift.Bookmark = Adodc_shift.Recordset.AbsolutePosition
    Call TDBCombo_shift_ItemChange
Else
    TDBCombo_shift.Text = ""
End If
End Sub

Private Sub TDBCombo_no_shift_ItemChange()
If TDBCombo_no_shift.ApproxCount > 0 And Not Trim(TDBCombo_no_shift.Text) = "" Then
    Adodc_no_shift.Recordset.Bookmark = TDBCombo_no_shift.Bookmark
    If int_mode = 0 Then
        Call view_data
    End If
Else
    Call clear_view_data
End If
End Sub

Private Sub TDBCombo_shift_ItemChange()
If Not (TDBCombo_shift.ApproxCount > 0 And TDBCombo_shift.Bookmark > 0) Then Exit Sub

Adodc_shift.Recordset.Bookmark = TDBCombo_shift.Bookmark
TDBCombo_shift.Text = TDBCombo_shift.Columns("kode_shift").Value
txt_nama_shift = TDBCombo_shift.Columns("nama_shift").Value
End Sub

Private Sub TDBDropDown2_DropDownClose()
If Not (TDBDropDown2.ApproxCount > 0 And TDBDropDown2.Bookmark > 0) Then Exit Sub

TDBGrid1.Columns("nama_shift").Value = TDBDropDown2.Columns("nama_shift").Value
End Sub

Private Sub TDBGrid1_ButtonClick(ByVal ColIndex As Integer)
If cek_validate_tdbcombo(TDBCombo_no_shift) = False Then Exit Sub

If TDBGrid1.Columns(ColIndex).Caption = "KODE KARYAWAN" Or _
TDBGrid1.Columns(ColIndex).Caption = "NAMA KARYAWAN" Then
    frm_lookup_mst_karyawan.public_int_mode = 0
    frm_lookup_mst_karyawan.Show 1
End If
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo ErrHandler

Dim i As Integer

Set Cols = TDBGrid1.Columns
i = TDBGrid1.Col
TDBGrid1.HoldFields

Adodc1.Recordset.Filter = getFilter()
TDBGrid1.Col = i
TDBGrid1.EditActive = True

TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
If TDBGrid1.ApproxCount < 1 Then
    Call clear_filter
    TDBGrid1.Col = i
End If
Exit Sub

ErrHandler:
'MsgBox Err.Source & ":" & vbCrLf & Err.Description
MsgBox "Tidak ada data pada kolom ini " & vbCr _
& "atau data filter yang anda inputkan tidak valid", vbCritical, "Filtering not allowed"
Call clear_filter
End Sub

Private Sub TDBGrid1_FormatText _
(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "BERLAKU" _
Or TDBGrid1.Columns(ColIndex).Caption = "BERAKHIR" Then
    Value = Format(Value, "dd-mm-yyyy")
End If
End Sub

Private Sub timer_combo_Timer()
Adodc_no_shift.RecordSource = _
"select *, trim(str(no_urut)) as no_urut_ from tm_shift_karyawan order by no_urut"
Adodc_no_shift.Refresh

TDBCombo_no_shift.RowSource = Adodc_no_shift
timer_combo.Enabled = False
End Sub

Private Sub timer_grid_detail_Timer()
Adodc1.RecordSource = "select * from td_shift_karyawan where no_urut=" _
& int_no_urut & " order by kode_shift, kode_karyawan"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
timer_grid_detail.Enabled = False
End Sub

Public Sub load_data_karyawan()
Adodc_karyawan.RecordSource = "select * from m_karyawan order by kode_karyawan"
Adodc_karyawan.Refresh
End Sub

Public Sub load_data_shift()
Adodc_shift.RecordSource = "select * from m_shift order by kode_shift"
Adodc_shift.Refresh
TDBCombo_shift.RowSource = Adodc_shift
End Sub
