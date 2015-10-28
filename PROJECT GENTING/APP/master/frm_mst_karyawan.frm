VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_karyawan 
   Caption         =   "MASTER KARYAWAN"
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
   Icon            =   "frm_mst_karyawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   1200
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
   Begin VB.Frame fra_entry 
      Height          =   2775
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   13935
      Begin VB.ComboBox cbo_jns_kelamin 
         Height          =   315
         ItemData        =   "frm_mst_karyawan.frx":058A
         Left            =   2160
         List            =   "frm_mst_karyawan.frx":0594
         TabIndex        =   27
         Text            =   "cbo_jns_kelamin"
         Top             =   2040
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker_tgl_lahir 
         Height          =   315
         Left            =   2160
         TabIndex        =   23
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   49938435
         CurrentDate     =   39270
      End
      Begin VB.TextBox txt_nama 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txt_alamat 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   7680
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txt_no_fix_telp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         MaxLength       =   30
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txt_no_hp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7680
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txt_tmp_lahir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
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
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_kode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_fp 
         Height          =   495
         Left            =   2160
         OleObjectBlob   =   "frm_mst_karyawan.frx":05A6
         TabIndex        =   24
         Top             =   360
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc_fp 
         Height          =   375
         Left            =   1440
         Top             =   360
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
      Begin VB.Line Line1 
         X1              =   720
         X2              =   10800
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "JNS KELAMIN"
         Height          =   195
         Left            =   720
         TabIndex        =   26
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ID FP"
         Height          =   195
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TGL/TMP LAHIR"
         Height          =   195
         Left            =   720
         TabIndex        =   22
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAMA"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   195
         Left            =   6720
         TabIndex        =   19
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Telp"
         Height          =   195
         Left            =   6720
         TabIndex        =   18
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. Hp"
         Height          =   195
         Left            =   6720
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   12303
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID FP"
      Columns(0).DataField=   "id_fp"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NIP"
      Columns(1).DataField=   "kode_karyawan"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NAMA"
      Columns(2).DataField=   "nama_karyawan"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TGL LAHIR"
      Columns(3).DataField=   "tanggal_lahir_"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TMP LAHIR"
      Columns(4).DataField=   "tempat_lahir"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "JNS KELAMIN"
      Columns(5).DataField=   "jenis_kelamin_"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "NO TELP"
      Columns(6).DataField=   "no_fix_telp"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "NO HP"
      Columns(7).DataField=   "no_hp"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "ALAMAT"
      Columns(8).DataField=   "alamat"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "tanggal_lahir"
      Columns(9).DataField=   "tanggal_lahir"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "jenis_kelamin"
      Columns(10).DataField=   "jenis_kelamin"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1535"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2408"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2328"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4710"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4630"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2064"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1984"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=3016"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2937"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2037"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(7)._MinWidth=10"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(8)._MinWidth=54215968"
      Splits(0)._ColumnProps(48)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(49)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(51)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(52)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(58)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(59)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(60)=   "Column(10)._MinWidth=60129312"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=98,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
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
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=110,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=35:Footing"
      _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=36:Selected"
      _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=37:Caption"
      _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(89)  =   "Named:id=38:HighlightRow"
      _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   360
      TabIndex        =   8
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
         Left            =   8280
         Picture         =   "frm_mst_karyawan.frx":267F
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_mst_karyawan.frx":2C09
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_mst_karyawan.frx":3193
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11520
         Picture         =   "frm_mst_karyawan.frx":371D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_mst_karyawan.frx":3CA7
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Report"
         Height          =   645
         Left            =   7200
         Picture         =   "frm_mst_karyawan.frx":4231
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_mst_karyawan.frx":47BB
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_mst_karyawan.frx":4D45
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_mst_karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FlagNew As Boolean
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim str_kode_rekening As String
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns


Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
check_validate_exist_new = False

SQL = "select count(kode_karyawan) as jml from m_karyawan where kode_karyawan = '" _
& Replace$(Trim$(txt_kode), Chr$(39), Chr$(96)) & "'"
rs.Open SQL, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("jml").Value > 0 Then
    MsgBox "Kode Karyawan Sudah Ada !", vbCritical, headerMSG
    txt_kode = ""
    If txt_kode.Enabled = True Then txt_kode.SetFocus
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi combo fp
If cek_validate_tdbcombo(TDBCombo_fp) = False Then
    MsgBox "ID belum dipilih", vbOKOnly + vbInformation, headerMSG
    TDBCombo_fp.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi Kode/NIP
If Trim(txt_kode) = "" Then
    MsgBox "NIP belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_kode.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi Nama
If Trim(txt_nama) = "" Then
    MsgBox "Nama belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_nama.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi tgl lahir
If IsNull(DTPicker_tgl_lahir.Value) = True Then
    MsgBox "Tanggal lahir belum diisi", vbOKOnly + vbInformation, headerMSG
    DTPicker_tgl_lahir.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi Tempat lahir
If Trim(txt_tmp_lahir) = "" Then
    MsgBox "Tempat lahir belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_tmp_lahir.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi Jenis Kelamin
If cbo_jns_kelamin.ListIndex < 0 Then
    MsgBox "Jenis kelamin belum dipilih", vbOKOnly + vbInformation, headerMSG
    cbo_jns_kelamin.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi No Fix Telp
If Trim(txt_no_fix_telp) = "" Then
    MsgBox "No. Telp belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_no_fix_telp.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi No HP
If Trim(txt_no_hp) = "" Then
    MsgBox "No. HP belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_no_hp.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi Alamat
If Trim(txt_alamat) = "" Then
    MsgBox "Alamat belum diisi", vbOKOnly + vbInformation, headerMSG
    txt_alamat.SetFocus
    check_validate_new = False
    Exit Function
End If

End Function

Private Sub load_data_grid()
timer_grid_detail.Enabled = True
End Sub

Private Sub fill_combo_fp()
Adodc_fp.RecordSource = "select trim(str(enrollnumber)) as enrollnumber from m_enroll order by enrollnumber"
Adodc_fp.Refresh

TDBCombo_fp.RowSource = Adodc_fp
End Sub

Private Sub cmd_refresh_Click()
Call load_data_grid
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "Data Karyawan belum dipilih", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Anda yakin ingin menghapus data karyawan '" _
    & TDBGrid1.Columns("kode_karyawan").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

'CnG.BeginTrans
CnG.Execute "delete from m_karyawan where kode_karyawan = '" & TDBGrid1.Columns("kode_karyawan").Value & "'"
'CnG.CommitTrans

Call load_data_grid
int_mode = 0
Call load_mode
End Sub

Private Sub set_data_fp(ByVal str_kode As String)
Adodc_fp.Recordset.MoveFirst
Adodc_fp.Recordset.Find ("enrollnumber='" & str_kode & "'")   ', 0, adSearchForward, 1)
If Not (Adodc_fp.Recordset.EOF = True Or Adodc_fp.Recordset.BOF = True) Then
    TDBCombo_fp.Bookmark = Adodc_fp.Recordset.AbsolutePosition
    Call TDBCombo_fp_ItemChange
Else
    TDBCombo_fp.Text = ""
End If
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    txt_kode = .Fields("kode_karyawan").Value
    Call set_data_fp(.Fields("id_fp").Value)
    txt_nama = .Fields("nama_karyawan").Value
    DTPicker_tgl_lahir.Value = .Fields("tanggal_lahir").Value
    txt_tmp_lahir = .Fields("tempat_lahir").Value
    cbo_jns_kelamin.ListIndex = .Fields("jenis_kelamin").Value
    txt_no_fix_telp = "" & .Fields("no_fix_telp").Value
    txt_no_hp = "" & .Fields("no_hp").Value
    txt_alamat = "" & .Fields("alamat").Value
End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_karyawan where kode_karyawan = '" _
& Adodc1.Recordset.Fields("kode_karyawan").Value & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_karyawan where kode_karyawan = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

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
    
    .Fields("kode_karyawan").Value = Trim(txt_kode)     ' key
    .Fields("id_fp").Value = Adodc_fp.Recordset.Fields("enrollnumber").Value
    .Fields("nama_karyawan").Value = Trim(txt_nama)
    .Fields("tanggal_lahir").Value = DTPicker_tgl_lahir.Value
    .Fields("tempat_lahir").Value = Trim(txt_tmp_lahir)
    .Fields("jenis_kelamin").Value = cbo_jns_kelamin.ListIndex
    .Fields("no_fix_telp").Value = Trim(txt_no_fix_telp)
    .Fields("no_hp").Value = Trim(txt_no_hp)
    .Fields("alamat").Value = Trim(txt_alamat)
    
    .Update
End With

'CnG.CommitTrans
End Sub

Private Sub edit_old_data()
'CnG.BeginTrans

With rsBound
    '.AddNew
    '.Fields("kode_karyawan").Value = Trim(txt_kode)     ' key
    .Fields("id_fp").Value = Adodc_fp.Recordset.Fields("enrollnumber").Value
    .Fields("nama_karyawan").Value = Trim(txt_nama)
    .Fields("tanggal_lahir").Value = DTPicker_tgl_lahir.Value
    .Fields("tempat_lahir").Value = Trim(txt_tmp_lahir)
    .Fields("jenis_kelamin").Value = cbo_jns_kelamin.ListIndex
    .Fields("no_fix_telp").Value = Trim(txt_no_fix_telp)
    .Fields("no_hp").Value = Trim(txt_no_hp)
    .Fields("alamat").Value = Trim(txt_alamat)
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

Call load_data_grid
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

Private Sub set_new_data()
cbo_jns_kelamin.ListIndex = 1
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_kode.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_kode.Enabled = True Then
        txt_kode.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    txt_kode.Enabled = False
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

Private Sub Command1_Click()
MsgBox Chr$(40)
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc_fp.ConnectionString = strConn
'MsgBox str_kode_rekening

Call load_data_grid
Call fill_combo_fp
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

Private Sub TDBCombo_fp_ItemChange()
If TDBCombo_fp.ApproxCount > 0 And Not Trim(TDBCombo_fp.Text) = "" Then
    Adodc_fp.Recordset.Bookmark = TDBCombo_fp.Bookmark
    TDBCombo_fp.Text = TDBCombo_fp.Columns("enrollnumber").Value
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

Private Sub timer_grid_detail_Timer()
Adodc1.RecordSource = _
"select *, format(tanggal_lahir,'dd-MM-yyyy') as tanggal_lahir_, " _
& "iif(jenis_kelamin=0, 'Wanita', 'Pria') as jenis_kelamin_ from m_karyawan order by kode_karyawan"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1

timer_grid_detail.Enabled = False
End Sub

Private Sub txt_kode_Change()
'If Not (int_mode = 1 Or int_mode = 2) Then Exit Sub
'
'If Trim(txt_kode) = "" Then
'    txt_kode_sub_rekening = ""
'Else
'    txt_kode_sub_rekening = str_kode_rekening & "-" & Trim(txt_kode)
'End If
End Sub

Private Sub txt_nama_Change()
'If Not (int_mode = 1 Or int_mode = 2) Then Exit Sub
'
'If Trim(txt_nama) = "" Then
'    txt_nama_sub_rekening = ""
'Else
'    txt_nama_sub_rekening = Trim(txt_nama)
'End If
End Sub
