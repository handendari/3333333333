VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_menu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER MENU"
   ClientHeight    =   8820
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   13770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton opt_sub_menu 
      Caption         =   "SUB MENU"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton opt_menu 
      Caption         =   "MENU"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   -120
      Top             =   2160
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4080
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
   Begin VB.Frame fra_sub_menu 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   13575
      Begin VB.Frame fra_entry_sub_menu 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Width           =   13335
         Begin MSAdodcLib.Adodc Adodc_menu 
            Height          =   375
            Left            =   -360
            Top             =   330
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Caption         =   ""
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
         Begin VB.CommandButton cmd_browse_form 
            Caption         =   "..."
            Height          =   310
            Left            =   12240
            TabIndex        =   8
            ToolTipText     =   "Press to browse data..."
            Top             =   600
            Width           =   465
         End
         Begin VB.TextBox txt_level 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CheckBox chk_flag_detail 
            Height          =   255
            Left            =   2160
            TabIndex        =   7
            Top             =   2400
            Value           =   1  'Checked
            Width           =   375
         End
         Begin VB.TextBox txt_parent_menu_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox chk_form_modal 
            Enabled         =   0   'False
            Height          =   255
            Left            =   9120
            TabIndex        =   9
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox chk_flag_sizeable 
            Enabled         =   0   'False
            Height          =   255
            Left            =   9120
            TabIndex        =   10
            Top             =   1680
            Width           =   375
         End
         Begin VB.ComboBox cbo_user_level 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frm_mst_menu.frx":000C
            Left            =   9120
            List            =   "frm_mst_menu.frx":0016
            TabIndex        =   11
            Text            =   "Pilihan"
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txt_sub_menu_action 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   19
            Top             =   3120
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.TextBox txt_form_title 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9120
            MaxLength       =   50
            TabIndex        =   18
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txt_form_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   9120
            MaxLength       =   50
            TabIndex        =   17
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txt_sub_menu_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   2
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt_sub_menu_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   3
            Top             =   960
            Width           =   3615
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_menu 
            Height          =   375
            Left            =   2160
            OleObjectBlob   =   "frm_mst_menu.frx":0027
            TabIndex        =   4
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEVEL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   600
            TabIndex        =   40
            Top             =   2040
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DETAIL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   600
            TabIndex        =   39
            Top             =   2400
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE INDUK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   990
            TabIndex        =   38
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SIZEABLE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   10
            Left            =   7800
            TabIndex        =   37
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FORM MODAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   7800
            TabIndex        =   36
            Top             =   1320
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEVEL USER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   7800
            TabIndex        =   34
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ACTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   3360
            TabIndex        =   33
            Top             =   3120
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JUDUL FORM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   7950
            TabIndex        =   32
            Top             =   960
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA FORM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   7950
            TabIndex        =   31
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MENU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   600
            TabIndex        =   30
            Top             =   1320
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE SUB MENU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   660
            TabIndex        =   29
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA SUB MENU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   630
            TabIndex        =   28
            Top             =   960
            Width           =   1290
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   6855
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   12091
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE SM"
         Columns(0).DataField=   "sub_menu_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA SM"
         Columns(1).DataField=   "sub_menu_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DETAIL"
         Columns(2).DataField=   "flag_detail"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA MENU"
         Columns(3).DataField=   "menu_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NAMA FORM"
         Columns(4).DataField=   "form_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "INDUK"
         Columns(5).DataField=   "parent_menu_code"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "JUDUL FORM"
         Columns(6).DataField=   "form_title"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "MODAL"
         Columns(7).DataField=   "form_modal"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   4
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "SIZEABLE"
         Columns(8).DataField=   "flag_sizeable"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   16
         Columns(9)._MaxComboItems=   5
         Columns(9).ValueItems(0)._DefaultItem=   0
         Columns(9).ValueItems(0).Value=   "2"
         Columns(9).ValueItems(0).Value.vt=   8
         Columns(9).ValueItems(0).DisplayValue=   "Admin"
         Columns(9).ValueItems(0).DisplayValue.vt=   8
         Columns(9).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(9).ValueItems(1)._DefaultItem=   0
         Columns(9).ValueItems(1).Value=   "1"
         Columns(9).ValueItems(1).Value.vt=   8
         Columns(9).ValueItems(1).DisplayValue=   "User"
         Columns(9).ValueItems(1).DisplayValue.vt=   8
         Columns(9).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(9).ValueItems.Count=   2
         Columns(9).Caption=   "LEVEL USER"
         Columns(9).DataField=   "user_level"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=6647137"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3651"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3572"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1244"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1164"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2831"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2752"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=512"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=4260"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=4180"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=1931"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1852"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=512"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(5)._MinWidth=1919247457"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=3995"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=3916"
         Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=512"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(6)._MinWidth=1919247457"
         Splits(0)._ColumnProps(39)=   "Column(7).Width=1270"
         Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=1191"
         Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(44)=   "Column(8).Width=1508"
         Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=1429"
         Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(49)=   "Column(9).Width=1879"
         Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=1799"
         Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=513"
         Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
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
         Caption         =   "DAFTAR SUB MENU (SM)"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   13160660
         ScrollTrack     =   -1  'True
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
         _StyleDefs(21)  =   "Splits(0).Style:id=29,.parent=1"
         _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=48,.parent=4"
         _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=30,.parent=2,.alignment=2"
         _StyleDefs(24)  =   "Splits(0).FooterStyle:id=31,.parent=3"
         _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=32,.parent=5"
         _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
         _StyleDefs(27)  =   "Splits(0).EditorStyle:id=43,.parent=7"
         _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8"
         _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
         _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=47,.parent=10"
         _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=49,.parent=11"
         _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=50,.parent=12"
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=16,.parent=29,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=30"
         _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=31"
         _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=28,.parent=29,.alignment=0"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=30"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=31"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=74,.parent=29,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=30"
         _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=31"
         _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=43"
         _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=20,.parent=29,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=30"
         _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=31"
         _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=43"
         _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=24,.parent=29,.alignment=0"
         _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=21,.parent=30"
         _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=22,.parent=31"
         _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=23,.parent=43"
         _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=70,.parent=29,.alignment=0"
         _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=30"
         _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=31"
         _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=43"
         _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=54,.parent=29,.alignment=0"
         _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=30"
         _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=31"
         _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=43"
         _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=66,.parent=29,.alignment=2"
         _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=30"
         _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=31"
         _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=43"
         _StyleDefs(65)  =   "Splits(0).Columns(8).Style:id=58,.parent=29,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=30"
         _StyleDefs(67)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=31"
         _StyleDefs(68)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=43"
         _StyleDefs(69)  =   "Splits(0).Columns(9).Style:id=62,.parent=29,.alignment=2"
         _StyleDefs(70)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=30"
         _StyleDefs(71)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=31"
         _StyleDefs(72)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=43"
         _StyleDefs(73)  =   "Named:id=33:Normal"
         _StyleDefs(74)  =   ":id=33,.parent=0"
         _StyleDefs(75)  =   "Named:id=34:Heading"
         _StyleDefs(76)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   ":id=34,.wraptext=-1"
         _StyleDefs(78)  =   "Named:id=35:Footing"
         _StyleDefs(79)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   "Named:id=36:Selected"
         _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   "Named:id=37:Caption"
         _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(84)  =   "Named:id=38:HighlightRow"
         _StyleDefs(85)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(86)  =   "Named:id=39:EvenRow"
         _StyleDefs(87)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(88)  =   "Named:id=40:OddRow"
         _StyleDefs(89)  =   ":id=40,.parent=33"
         _StyleDefs(90)  =   "Named:id=41:RecordSelector"
         _StyleDefs(91)  =   ":id=41,.parent=34"
         _StyleDefs(92)  =   "Named:id=42:FilterBar"
         _StyleDefs(93)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame fra_menu 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   13575
      Begin VB.Frame fra_entry 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Width           =   13335
         Begin VB.TextBox txt_menu_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5280
            MaxLength       =   50
            TabIndex        =   14
            Top             =   1320
            Width           =   4095
         End
         Begin VB.TextBox txt_menu_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5280
            MaxLength       =   20
            TabIndex        =   13
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txt_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5280
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1680
            Width           =   4095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA MENU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   4050
            TabIndex        =   25
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE MENU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   4140
            TabIndex        =   24
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   3975
            TabIndex        =   23
            Top             =   1680
            Width           =   1050
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   12091
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE MENU"
         Columns(0).DataField=   "menu_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA MENU"
         Columns(1).DataField=   "menu_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KETERANGAN"
         Columns(2).DataField=   "description"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3995"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3916"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=6647137"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=8096"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8017"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=8255"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=8176"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=1"
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
         Caption         =   "DAFTAR MENU"
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
         _StyleDefs(21)  =   "Splits(0).Style:id=29,.parent=1"
         _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=48,.parent=4"
         _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=30,.parent=2"
         _StyleDefs(24)  =   "Splits(0).FooterStyle:id=31,.parent=3"
         _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=32,.parent=5"
         _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
         _StyleDefs(27)  =   "Splits(0).EditorStyle:id=43,.parent=7"
         _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8"
         _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
         _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=47,.parent=10"
         _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=49,.parent=11"
         _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=50,.parent=12"
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=16,.parent=29,.alignment=2"
         _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=30"
         _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=31"
         _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=20,.parent=29,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=30"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=31"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=24,.parent=29,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=30"
         _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=31"
         _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=43"
         _StyleDefs(45)  =   "Named:id=33:Normal"
         _StyleDefs(46)  =   ":id=33,.parent=0"
         _StyleDefs(47)  =   "Named:id=34:Heading"
         _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   ":id=34,.wraptext=-1"
         _StyleDefs(50)  =   "Named:id=35:Footing"
         _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   "Named:id=36:Selected"
         _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=39:EvenRow"
         _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(60)  =   "Named:id=40:OddRow"
         _StyleDefs(61)  =   ":id=40,.parent=33"
         _StyleDefs(62)  =   "Named:id=41:RecordSelector"
         _StyleDefs(63)  =   ":id=41,.parent=34"
         _StyleDefs(64)  =   "Named:id=42:FilterBar"
         _StyleDefs(65)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame fra_control2 
      Height          =   1095
      Left            =   240
      TabIndex        =   35
      Top             =   7560
      Width           =   13335
      Begin prj_panji.vbButton cmd_new2 
         Height          =   705
         Left            =   1020
         TabIndex        =   41
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Tambah"
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
         MICON           =   "frm_mst_menu.frx":1FD6
         PICN            =   "frm_mst_menu.frx":1FF2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_save2 
         Height          =   705
         Left            =   2040
         TabIndex        =   42
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Simpan"
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
         MICON           =   "frm_mst_menu.frx":3084
         PICN            =   "frm_mst_menu.frx":30A0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_edit2 
         Height          =   705
         Left            =   3060
         TabIndex        =   43
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Ubah"
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
         MICON           =   "frm_mst_menu.frx":4132
         PICN            =   "frm_mst_menu.frx":414E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_delete2 
         Height          =   705
         Left            =   4080
         TabIndex        =   44
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Hapus"
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
         MICON           =   "frm_mst_menu.frx":51E0
         PICN            =   "frm_mst_menu.frx":51FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_cancel2 
         Height          =   705
         Left            =   5100
         TabIndex        =   45
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Batal"
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
         MICON           =   "frm_mst_menu.frx":628E
         PICN            =   "frm_mst_menu.frx":62AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_exit2 
         Height          =   705
         Left            =   10140
         TabIndex        =   46
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Keluar"
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
         MICON           =   "frm_mst_menu.frx":733C
         PICN            =   "frm_mst_menu.frx":7358
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
   Begin VB.Frame fra_control1 
      Height          =   1095
      Left            =   240
      TabIndex        =   20
      Top             =   7560
      Width           =   13335
      Begin prj_panji.vbButton cmd_new 
         Height          =   705
         Left            =   1320
         TabIndex        =   47
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Tambah"
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
         MICON           =   "frm_mst_menu.frx":83EA
         PICN            =   "frm_mst_menu.frx":8406
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_save 
         Height          =   705
         Left            =   2340
         TabIndex        =   48
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Simpan"
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
         MICON           =   "frm_mst_menu.frx":9498
         PICN            =   "frm_mst_menu.frx":94B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_edit 
         Height          =   705
         Left            =   3360
         TabIndex        =   49
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Ubah"
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
         MICON           =   "frm_mst_menu.frx":A546
         PICN            =   "frm_mst_menu.frx":A562
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_delete 
         Height          =   705
         Left            =   4380
         TabIndex        =   50
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Hapus"
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
         MICON           =   "frm_mst_menu.frx":B5F4
         PICN            =   "frm_mst_menu.frx":B610
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_cancel 
         Height          =   705
         Left            =   5400
         TabIndex        =   51
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Batal"
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
         MICON           =   "frm_mst_menu.frx":C6A2
         PICN            =   "frm_mst_menu.frx":C6BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmd_exit 
         Height          =   705
         Left            =   10440
         TabIndex        =   52
         Top             =   270
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Keluar"
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
         MICON           =   "frm_mst_menu.frx":D750
         PICN            =   "frm_mst_menu.frx":D76C
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
End
Attribute VB_Name = "frm_mst_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMenu As New ADODB.Recordset
Dim rsSubMenu As New ADODB.Recordset
Dim int_mode, int_mode2 As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub cmd_browse_form_Click()
frm_mst_form.public_int_mode = 0
frm_mst_form.cmd_Select.Visible = True
frm_mst_form.Show 1
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_exit2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call load_data_user_access(Me)
'Adodc1.ConnectionString = strConn
'Adodc2.ConnectionString = strConn
'Adodc_menu.ConnectionString = strConn
'rs.ActiveConnection = strConn

Call load_data_menu
Call load_data_sub_menu

int_mode = 0
Call load_mode
int_mode2 = 0
Call load_mode2

opt_menu = True
Call opt_menu_Click
'Call set_sizeable_form(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_menu = Nothing
End Sub

Private Sub opt_menu_Click()
    Call opt_Click
End Sub

Private Sub opt_sub_menu_Click()
Call opt_Click
End Sub

Private Sub opt_Click()
If cmd_cancel.Enabled = True Then
    MsgBox "Proses Input Belum Selesai...", vbInformation, headerMSG
    Exit Sub
End If

If opt_menu Then
    fra_menu.Visible = True
    fra_control1.Visible = True
    fra_control2.Visible = False
    fra_sub_menu.Visible = False
    
ElseIf opt_sub_menu Then
    fra_menu.Visible = False
    fra_sub_menu.Visible = True
    fra_control2.Visible = True
    fra_control1.Visible = False
End If
End Sub



Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "BIRTH DATE" Or _
TDBGrid1.Columns(ColIndex).Caption = "START WORKING" Or _
TDBGrid1.Columns(ColIndex).Caption = "END WORKING" Then
    Value = Format(Value, "dd-mm-yyyy")
End If
End Sub

Private Sub load_data_menu()
'Adodc1.RecordSource = "select * from m_menu order by menu_code"
'Adodc1.Refresh
'TDBGrid1.DataSource = Adodc1
'
'Adodc_menu.RecordSource = "select * from m_menu order by menu_code"
'Adodc_menu.Refresh
'TDBCombo_menu.RowSource = Adodc_menu
If rsMenu.State = 1 Then rsMenu.Close

SQL = "select * from m_menu order by menu_code"
rsMenu.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

TDBGrid1.DataSource = rsMenu
TDBCombo_menu.RowSource = rsMenu
'rs.Close

End Sub

Private Sub load_data_sub_menu()
'Adodc2.RecordSource = "select * from m_sub_menu order by sub_menu_code"
'Adodc2.Refresh
'TDBGrid2.DataSource = Adodc2
If rsSubMenu.State = 1 Then rsSubMenu.Close

SQL = "select * from m_sub_menu order by sub_menu_code"
rsSubMenu.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

TDBGrid2.DataSource = rsSubMenu
End Sub

'==============================

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(menu_code) as rec_count from m_menu where menu_code = '" & Trim(txt_menu_code) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data Sudah Ada...", vbCritical, headerMSG
txt_menu_code = ""
If txt_menu_code.Enabled = True Then txt_menu_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not txt_menu_code = rsMenu.Fields("menu_code").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_menu_code) = "" Then
    MsgBox "Kode Menu Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_menu_code.SetFocus
    check_validate_new = False
    Exit Function
End If

If Trim(txt_menu_name) = "" Then
    MsgBox "Nama Menu Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_menu_name.SetFocus
    check_validate_new = False
    Exit Function
End If

End Function

Private Sub cmd_cancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmd_delete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
    Exit Sub
End If

If check_exist_sub_menu = True Then
    MsgBox "Data Ditemukan Pada Sub Menu...", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
    & TDBGrid1.Columns("menu_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_sub_menu where menu_code = '" & TDBGrid1.Columns("menu_code").Value & "'"
CnG.Execute "delete from m_menu where menu_code = '" & TDBGrid1.Columns("menu_code").Value & "'"
CnG.CommitTrans

Call load_data_menu
int_mode = 0
Call load_mode
End Sub

Private Function check_exist_sub_menu() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_exist_sub_menu = False

str_sql = "select count(menu_code) as rec_count from m_sub_menu where menu_code = '" _
    & rsMenu.Fields("menu_code").Value & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic

If rs.Fields("rec_count").Value > 0 Then
    check_exist_sub_menu = True
    Exit Function
End If
End Function


Public Sub set_edit_data()
With rsMenu

    txt_menu_code = .Fields("menu_code").Value
    txt_menu_name = .Fields("menu_name").Value
    txt_description = .Fields("description").Value
    
End With
End Sub

Private Sub cmd_Edit_Click()
If rscari.State = 1 Then rscari.Close
rscari.Open "select * from m_menu where menu_code = '" _
& rsMenu.Fields("menu_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub cmd_new_Click()
If rsMenu.State = 1 Then rsMenu.Close
rsMenu.Open "select * from m_menu where menu_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub insert_new_data()
CnG.BeginTrans

With rsMenu
    .AddNew
    
    .Fields("menu_code").Value = Trim(txt_menu_code)
    '---------------------------------------------------------------------------
    .Fields("menu_name").Value = Trim(txt_menu_name)
    .Fields("description").Value = Trim(txt_description)
    
    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
'On Error GoTo err_capture
Dim strsql As String

CnG.BeginTrans
strsql = "UPDATE m_menu SET " _
        & "menu_code = '" & Trim(txt_menu_code.Text) & "'," _
        & "menu_name = '" & Trim(txt_menu_name.Text) & "'," _
        & "description = '" & Trim(txt_description) & "' " _
        & "WHERE menu_code = '" & Trim(txt_menu_code.Text) & "'"
CnG.Execute (strsql)

'With rsBound
'    '.AddNew
'
'    .Fields("menu_code").Value = Trim(txt_menu_code)
'    '---------------------------------------------------------------------------
'    .Fields("menu_name").Value = Trim(txt_menu_name)
'    .Fields("description").Value = Trim(txt_description)
'
'    CnG.Execute "update m_sub_menu set menu_code = '" & Trim(txt_menu_code) & "', menu_name = '" _
'    & Trim(txt_menu_name) & "' where menu_code = '" _
'    & Adodc1.Recordset.Fields("menu_code").Value & "'"
'
'    .Update
'End With
CnG.CommitTrans

Exit Sub
err_capture:
rsMenu.CancelBatch adAffectCurrent: rsMenu.Close: CnG.RollbackTrans
End Sub

Private Sub cmd_save_Click()
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

Call load_data_menu
int_mode = 0
Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
cmd_new.Enabled = a And blnUser_Add
cmd_save.Enabled = b
cmd_edit.Enabled = c And blnUser_Edit
cmd_delete.Enabled = d And blnUser_Delete
cmd_cancel.Enabled = e

'CmdPrint.Enabled = f
'cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data(ByRef fra1 As Frame)
Dim Ctr As CONTROL

For Each Ctr In Me
    If TypeOf Ctr Is Timer Then GoTo continew

    If Ctr.Container Is fra1 Then
    
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not Ctr.name = "txt_company_name" Then
                Ctr.Text = ""
            End If
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not Ctr.name = "TDBCombo_company" Then
                Ctr.Text = ""
            End If
        ElseIf TypeOf Ctr Is DTPicker Then
            Ctr.Value = Now
        End If

    End If
    
continew:
Next

End Sub

Private Sub Command1_Click()
On Error GoTo err_handle

Dim Ctr As CONTROL
For Each Ctr In Me
    If TypeOf Ctr Is Timer Then GoTo continew
        
        If Ctr.Container Is fra_entry Then
            MsgBox Ctr.name
        End If
        
continew:
    'MsgBox Ctr.Name
Next

Exit Sub

err_handle:
MsgBox Err.Description
MsgBox Ctr.name
End Sub

Private Function get_actual_number() As Integer
Dim rs As New ADODB.Recordset
Dim str1 As String, int1 As Integer

str1 = "select isnull(max(menu_code),0) as recs from m_menu"
rs.Open str1, CnG, adOpenStatic, adLockBatchOptimistic

If rs.RecordCount > 0 Then
    int1 = rs!recs + 1
Else
    int1 = 1
End If

get_actual_number = int1
End Function

Private Sub set_new_data()
'txt_menu_code = get_actual_number
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data(fra_entry)
    fra_entry.Visible = True
    txt_menu_code.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_menu_code.Enabled = True Then
        txt_menu_code.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data(fra_entry)
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
'    txt_menu_code.Enabled = False
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



'--- ---
Private Sub set_data_mode2()
If int_mode2 = 1 Then        'NEW
    Call clear_view_data(fra_entry_sub_menu)
    fra_entry_sub_menu.Visible = True
    txt_sub_menu_code.Enabled = True
    TDBGrid2.Enabled = False
    Call set_new_data2
    
    If txt_sub_menu_code.Enabled = True Then
        txt_sub_menu_code.SetFocus
    End If
    
ElseIf int_mode2 = 0 Then    'VIEW
    Call clear_view_data(fra_entry_sub_menu)
    fra_entry_sub_menu.Visible = False
    TDBGrid2.Enabled = True

ElseIf int_mode2 = 2 Then    'EDIT
    Call set_edit_data2
'    txt_sub_menu_code.Enabled = False
    fra_entry_sub_menu.Visible = True
    TDBGrid2.Enabled = False
End If
End Sub

Private Sub load_mode2()
If int_mode2 = 1 Then        ' new
    Call set_buttons_enable2(False, True, False, False, True, False, False)
ElseIf int_mode2 = 0 Then    ' view
    Call set_buttons_enable2(True, False, True, True, False, True, True)
ElseIf int_mode2 = 2 Then    ' edit/revise
    Call set_buttons_enable2(False, True, False, False, True, False, False)
End If

Call set_data_mode2
End Sub

Private Sub set_buttons_enable2(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
cmd_new2.Enabled = a And blnUser_Add
cmd_save2.Enabled = b
cmd_edit2.Enabled = c And blnUser_Edit
cmd_delete2.Enabled = d And blnUser_Delete
cmd_cancel2.Enabled = e

'CmdPrint.Enabled = f
'cmd_refresh.Enabled = g
End Sub

Public Sub set_edit_data2()
With rsSubMenu

    txt_sub_menu_code = .Fields("sub_menu_code").Value
    txt_sub_menu_name = .Fields("sub_menu_name").Value
    Call set_data_tdbcombo_recordset(rsMenu, TDBCombo_menu, "menu_code='" & .Fields("menu_code").Value & "'")
    Call TDBCombo_menu_ItemChange
    txt_parent_menu_code = Trim("" & .Fields("parent_menu_code").Value)
    chk_flag_detail.Value = Val("" & .Fields("flag_detail").Value)
    
    txt_level = .Fields("level").Value
    chk_flag_detail = .Fields("flag_detail").Value
    txt_form_name = .Fields("form_name").Value
    txt_form_title = .Fields("form_title").Value
    txt_sub_menu_action = IIf(IsNull(.Fields("sub_menu_action").Value), 0, .Fields("sub_menu_action").Value)
    chk_form_modal.Value = .Fields("form_modal").Value
    chk_flag_sizeable.Value = Val("" & .Fields("flag_sizeable").Value)
    
    cbo_user_level.ListIndex = IIf(.Fields("user_level").Value = 2, 0, 1)

End With
End Sub

Private Sub TDBCombo_menu_ItemChange()
If TDBCombo_menu.ApproxCount > 0 Then
    TDBCombo_menu.Text = TDBCombo_menu.Columns("menu_name").Value
End If
End Sub

Private Sub cmd_new2_Click()
'If rscari.State = 1 Then rscari.Close
'rscari.Open "select * from m_menu where menu_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

int_mode2 = 1
Call load_mode2
End Sub

Private Sub insert_new_data2()
Dim rs1 As New ADODB.Recordset

rs1.Open "select * from m_sub_menu where sub_menu_code='uOu'", CnG, adOpenKeyset, adLockOptimistic
CnG.BeginTrans

With rs1
    .AddNew
    
    .Fields("sub_menu_code").Value = Trim(txt_sub_menu_code)
    '-------------------------------
    .Fields("sub_menu_name").Value = Trim(txt_sub_menu_name)
    .Fields("menu_code").Value = TDBCombo_menu.Columns("menu_code").Value
    .Fields("menu_name").Value = TDBCombo_menu.Columns("menu_name").Value
    .Fields("parent_menu_code").Value = Trim(txt_parent_menu_code)
    
    .Fields("level").Value = Val(txt_level)
    .Fields("flag_detail").Value = IIf(chk_flag_detail.Value = vbChecked, 1, 0)
    .Fields("form_name").Value = Trim(txt_form_name)
    .Fields("form_title").Value = Trim(txt_form_title)
    .Fields("sub_menu_action").Value = Trim(txt_sub_menu_action)
    .Fields("form_modal").Value = IIf(chk_form_modal.Value = vbChecked, 1, 0)
    .Fields("flag_sizeable").Value = IIf(chk_flag_sizeable.Value = vbChecked, 1, 0)
    
    .Fields("user_level").Value = IIf(cbo_user_level.ListIndex = 0, 2, 1)
    
    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data2()
Dim rs1 As New ADODB.Recordset
'On Error GoTo err_capture
Dim strsql As String

'rs1.Open "select * from m_sub_menu where sub_menu_code='" _
'        & rsSubMenu.Fields("sub_menu_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic
CnG.BeginTrans

strsql = "UPDATE m_sub_menu SET " _
        & "sub_menu_code = '" & Trim(txt_sub_menu_code.Text) & "'," _
        & "sub_menu_name = '" & Replace(Trim(txt_sub_menu_name.Text), "'", "") & "'," _
        & "menu_code = '" & TDBCombo_menu.Columns("menu_code").Value & "'," _
        & "menu_name = '" & TDBCombo_menu.Columns("menu_name").Value & "'," _
        & "parent_menu_code = '" & Trim(txt_parent_menu_code.Text) & "'," _
        & "level = '" & Val(txt_level.Text) & "'," _
        & "flag_detail = '" & IIf(chk_flag_detail.Value = vbChecked, 1, 0) & "'," _
        & "form_name = '" & Trim(txt_form_name.Text) & "'," _
        & "form_title = '" & Replace(Trim(txt_form_title.Text), "'", "") & "'," _
        & "sub_menu_action = '" & Trim(txt_sub_menu_action.Text) & "'," _
        & "form_modal = '" & IIf(chk_form_modal.Value = vbChecked, 1, 0) & "'," _
        & "flag_sizeable = '" & IIf(chk_flag_sizeable.Value = vbChecked, 1, 0) & "'," _
        & "user_level = '" & IIf(cbo_user_level.ListIndex = 0, 2, 1) & "' " _
        & "WHERE sub_menu_code = '" & rsSubMenu.Fields("sub_menu_code").Value & "'"
CnG.Execute strsql

'With rs1
'    '.AddNew
'
'    .Fields("sub_menu_code").Value = Trim(txt_sub_menu_code)
'    '-------------------------------
'    .Fields("sub_menu_name").Value = Trim(txt_sub_menu_name)
'    .Fields("menu_code").Value = TDBCombo_menu.Columns("menu_code").Value
'    .Fields("menu_name").Value = TDBCombo_menu.Columns("menu_name").Value
'    .Fields("parent_menu_code").Value = Trim(txt_parent_menu_code)
'
'    .Fields("level").Value = Val(txt_level)
'    .Fields("flag_detail").Value = IIf(chk_flag_detail.Value = vbChecked, 1, 0)
'    .Fields("form_name").Value = Trim(txt_form_name)
'    .Fields("form_title").Value = Trim(txt_form_title)
'    .Fields("sub_menu_action").Value = Trim(txt_sub_menu_action)
'    .Fields("form_modal").Value = IIf(chk_form_modal.Value = vbChecked, 1, 0)
'    .Fields("flag_sizeable").Value = IIf(chk_flag_sizeable.Value = vbChecked, 1, 0)
'
'    .Fields("user_level").Value = IIf(cbo_user_level.ListIndex = 0, 2, 1)
'
'    .Update
'End With

CnG.CommitTrans

Exit Sub
err_capture:
rsMenu.CancelBatch adAffectCurrent: rsMenu.Close: CnG.RollbackTrans
End Sub

Private Sub cmd_save2_Click()
If int_mode2 = 1 Then
    If Not check_validate_new2 Then Exit Sub
    If check_validate_exist_new2 Then
        Call check_invalid2: Exit Sub
    End If
    Call insert_new_data2
ElseIf int_mode2 = 2 Then
    If Not check_validate_new2 Then Exit Sub
    If check_validate_exist_edit2 Then
        Call check_invalid2: Exit Sub
    End If
    Call edit_old_data2
End If

Call load_data_sub_menu
int_mode2 = 0
Call load_mode2
End Sub


Private Function check_validate_exist_new2() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new2 = False

str_sql = "select count(sub_menu_code) as rec_count from m_sub_menu where sub_menu_code = '" _
& Trim(txt_sub_menu_code) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new2 = True
    Exit Function
End If
End Function

Private Sub check_invalid2()
MsgBox "Data Sudah Ada...", vbCritical, headerMSG
txt_sub_menu_code = ""
If txt_sub_menu_code.Enabled = True Then txt_sub_menu_code.SetFocus
End Sub

Private Function check_validate_exist_edit2() As Boolean
check_validate_exist_edit2 = False

If Not txt_sub_menu_code = rsSubMenu.Fields("sub_menu_code").Value And _
check_validate_exist_new2 Then
    check_validate_exist_edit2 = True
    Exit Function
End If
End Function

Private Function check_validate_new2() As Boolean
check_validate_new2 = True

If Trim(txt_sub_menu_code) = "" Then
    MsgBox "Kode Sub Menu Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_sub_menu_code.SetFocus
    check_validate_new2 = False
    Exit Function
End If

If Trim(txt_sub_menu_name) = "" Then
    MsgBox "Nama Sub Menu Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_sub_menu_name.SetFocus
    check_validate_new2 = False
    Exit Function
End If

End Function

Private Sub cmd_cancel2_Click()
int_mode2 = 0
Call load_mode2
End Sub

Private Sub cmd_delete2_Click()
Dim i As Integer

If Not (TDBGrid2.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
    & rsSubMenu.Fields("sub_menu_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_sub_menu where sub_menu_code = '" & rsSubMenu.Fields("sub_menu_code").Value & "'"
CnG.CommitTrans

Call load_data_sub_menu
int_mode2 = 0
Call load_mode2
End Sub

Private Sub cmd_Edit2_Click()
If rscari.State = 1 Then rscari.Close
rscari.Open "select * from m_menu where menu_code = '" _
& rsMenu.Fields("menu_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode2 = 2
Call load_mode2
End Sub

Private Sub set_new_data2()
chk_flag_detail = vbChecked
txt_level = "2"

cbo_user_level.ListIndex = 0
chk_flag_sizeable = vbUnchecked
chk_form_modal = vbUnchecked
End Sub

Private Sub TDBGrid1_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid1, Adodc1)
End Sub

Private Sub TDBGrid2_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid2, Adodc2)
End Sub

Private Sub txt_form_name_KeyPress(KeyAscii As Integer)
If Not int_mode = 0 And KeyAscii = 13 Then
    Call cmd_browse_form_Click
End If
End Sub
