VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_jamsostek 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER JAMSOSTEK"
   ClientHeight    =   6690
   ClientLeft      =   -15
   ClientTop       =   240
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
   Icon            =   "frm_mst_jamsostek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_entry 
      Height          =   2325
      Left            =   240
      TabIndex        =   16
      Top             =   2730
      Width           =   11295
      Begin VB.TextBox txt_jp_employee 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   7050
         MaxLength       =   50
         TabIndex        =   5
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox txt_max_jp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   7050
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1380
         Width           =   1695
      End
      Begin VB.TextBox txt_jp_company 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   7050
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1020
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   4020
         OleObjectBlob   =   "frm_mst_jamsostek.frx":000C
         TabIndex        =   27
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5820
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Top             =   300
         Width           =   3495
      End
      Begin VB.TextBox txt_tk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1380
         Width           =   1695
      End
      Begin VB.TextBox txt_salary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cbo_shiftable 
         Height          =   315
         ItemData        =   "frm_mst_jamsostek.frx":1F72
         Left            =   120
         List            =   "frm_mst_jamsostek.frx":1F7C
         TabIndex        =   25
         Text            =   "..."
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt_prs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox txt_jkm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1020
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
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_jkk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   1
         Top             =   660
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_company 
         Height          =   375
         Left            =   5040
         Top             =   270
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MAX SALARY FOR JP"
         Height          =   435
         Left            =   5880
         TabIndex        =   43
         Top             =   1350
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "JP (CO) (%)"
         Height          =   195
         Left            =   5910
         TabIndex        =   42
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "JP (EMP) (%)"
         Height          =   195
         Left            =   5910
         TabIndex        =   41
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN"
         Height          =   195
         Left            =   2580
         TabIndex        =   28
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "JHT (EMP) (%)"
         Height          =   195
         Left            =   2580
         TabIndex        =   24
         Top             =   1380
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "JHT (CO) (%)"
         Height          =   195
         Left            =   2580
         TabIndex        =   23
         Top             =   1740
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "JKK (%)"
         Height          =   195
         Left            =   2580
         TabIndex        =   21
         Top             =   660
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JKM (%)"
         Height          =   195
         Left            =   2580
         TabIndex        =   20
         Top             =   1020
         Width           =   615
      End
   End
   Begin VB.Frame fra_entry1 
      Height          =   1815
      Left            =   240
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txt_jamsostek_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4890
         MaxLength       =   10
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_name_jamsostek 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4890
         MaxLength       =   50
         TabIndex        =   36
         Top             =   930
         Width           =   3495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "NAMA JAMSOSTEK"
         Height          =   195
         Left            =   3480
         TabIndex        =   40
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "KODE JAMSOSTEK"
         Height          =   195
         Left            =   3480
         TabIndex        =   39
         Top             =   600
         Width           =   1320
      End
   End
   Begin VB.TextBox txt_jamsostek_name_type 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3390
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   30
      Top             =   600
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   17
      Top             =   5160
      Width           =   11295
      Begin VB.CommandButton cmdDelete_All 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   9000
         Picture         =   "frm_mst_jamsostek.frx":1F89
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew_Master 
         Caption         =   "&New"
         Height          =   645
         Left            =   7980
         Picture         =   "frm_mst_jamsostek.frx":2513
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_select 
         Caption         =   "&Select"
         Height          =   645
         Left            =   6120
         Picture         =   "frm_mst_jamsostek.frx":2A9D
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Refresh"
         Height          =   645
         Left            =   6840
         Picture         =   "frm_mst_jamsostek.frx":3027
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save Detail"
         Height          =   645
         Left            =   1800
         Picture         =   "frm_mst_jamsostek.frx":35B1
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5040
         Picture         =   "frm_mst_jamsostek.frx":3B3B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   10080
         Picture         =   "frm_mst_jamsostek.frx":40C5
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New Detail"
         Height          =   645
         Left            =   720
         Picture         =   "frm_mst_jamsostek.frx":464F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_jamsostek.frx":4BD9
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Detail"
         Height          =   645
         Left            =   3960
         Picture         =   "frm_mst_jamsostek.frx":5163
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Detail"
         Height          =   645
         Left            =   2880
         Picture         =   "frm_mst_jamsostek.frx":56ED
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_jamsostek 
      Height          =   375
      Left            =   1590
      OleObjectBlob   =   "frm_mst_jamsostek.frx":5C77
      TabIndex        =   31
      Top             =   600
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_jamsostek 
      Height          =   375
      Left            =   1710
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4005
      Left            =   240
      TabIndex        =   0
      Top             =   1050
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7064
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "COMPANY CODE"
      Columns(0).DataField=   "company_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "COMPANY NAME"
      Columns(1).DataField=   "company_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "JKK"
      Columns(2).DataField=   "jkk"
      Columns(2).NumberFormat=   "General Number"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "JKM"
      Columns(3).DataField=   "jkm"
      Columns(3).NumberFormat=   "General Number"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "JHT (EMP)"
      Columns(4).DataField=   "tk"
      Columns(4).NumberFormat=   "General Number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "JHT (CO)"
      Columns(5).DataField=   "prs"
      Columns(5).NumberFormat=   "General Number"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "JP (EMP)"
      Columns(6).DataField=   "jp_emp"
      Columns(6).NumberFormat=   "General Number"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "JP (CO)"
      Columns(7).DataField=   "jp_co"
      Columns(7).NumberFormat=   "General Number"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "MAX SALARY"
      Columns(8).DataField=   "max_jp_salary"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=6297"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6218"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1614"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1535"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1561"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1482"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=1693"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1614"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=1640"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1561"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(5)._MinWidth=1768976485"
      Splits(0)._ColumnProps(33)=   "Column(6).Width=1614"
      Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=1535"
      Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=1614"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=1535"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(43)=   "Column(8).Width=2884"
      Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=2805"
      Splits(0)._ColumnProps(46)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(47)=   "Column(8).Order=9"
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
      Caption         =   "LIST OF TITLE"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(70)  =   "Named:id=33:Normal"
      _StyleDefs(71)  =   ":id=33,.parent=0"
      _StyleDefs(72)  =   "Named:id=34:Heading"
      _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=34,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=35:Footing"
      _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=36:Selected"
      _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=37:Caption"
      _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(81)  =   "Named:id=38:HighlightRow"
      _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=39:EvenRow"
      _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(85)  =   "Named:id=40:OddRow"
      _StyleDefs(86)  =   ":id=40,.parent=33"
      _StyleDefs(87)  =   "Named:id=41:RecordSelector"
      _StyleDefs(88)  =   ":id=41,.parent=34"
      _StyleDefs(89)  =   "Named:id=42:FilterBar"
      _StyleDefs(90)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "JAMSOSTEK TYPE"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MASTER JAMSOSTEK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5235
      TabIndex        =   29
      Top             =   0
      Width           =   3345
   End
End
Attribute VB_Name = "frm_mst_jamsostek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(company_code) as rec_count from m_jamsostek_detail " _
        & "where jamsostek_code = '" & TDBCombo_jamsostek.Text & "' AND company_code = '" _
        & Replace$(Trim$(TDBCombo_company.Text), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
TDBCombo_company.Text = ""
txt_company_name.Text = ""
TDBCombo_company.SetFocus
'If txt_jkk.Enabled = True Then txt_jkk.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not TDBCombo_company.Text = Adodc1.Recordset.Fields("company_code").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi JKK
If Trim(txt_jkk) = "" Then
    MsgBox "JKK Is Empty", vbOKOnly + vbInformation, headerMSG
    txt_jkk.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi JKM
If Trim(txt_jkm) = "" Then
    MsgBox "JKM Is Empty!", vbOKOnly + vbInformation, headerMSG
    txt_jkm.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi TK
If Trim(txt_tk) = "" Then
    MsgBox "TK Is Empty!", vbOKOnly + vbInformation, headerMSG
    txt_tk.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi PRS
If Trim(txt_prs) = "" Then
    MsgBox "PRS Is Empty!", vbOKOnly + vbInformation, headerMSG
    txt_prs.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi employee nick name
If Trim(TDBCombo_company) = "" Then
    MsgBox "Company is empty!", vbOKOnly + vbInformation, headerMSG
    TDBCombo_company.SetFocus
    check_validate_new = False
    Exit Function
End If

''validasi shiftable cbo
'If cbo_shiftable.ListIndex < 0 Then
'    MsgBox "Shiftable is not selected!", vbOKOnly + vbInformation, headerMSG
'    cbo_shiftable.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
''validasi description
'If Trim(txt_prs) = "" Then
'    MsgBox "Description is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_prs.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
End Function

'Private Sub load_data()
'timer1.Enabled = True
'End Sub

Private Sub cmd_refresh_Click()
Call load_data_jamsostek
End Sub

Private Sub cmd_select_Click()
Dim i As Integer
Dim item

If Not TDBGrid1.ApproxCount > 0 Then
    Exit Sub
End If
    
'Set SelBks = TDBGrid1.SelBookmarks
'i = 0

If public_int_mode = 1 Then
    With frm_mst_job_history
        .txt_title_code = Adodc1.Recordset.Fields("title_code").Value
        .txt_title_name = Adodc1.Recordset.Fields("title_name").Value
    End With
    
End If

Call CmdExit_Click

Exit Sub
err_capture:
MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
int_mode = 0
Call load_mode
CmdNew_Master.Caption = "&New"
End Sub

Private Sub cmdDelete_All_Click()
Dim i As Integer

On Error GoTo Err
    i = MsgBox("Are you sure want to delete data '" _
            & txt_jamsostek_name_type.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_jamsostek_detail where " _
                & "jamsostek_code = '" & TDBCombo_jamsostek.Text & "'"
        CnG.Execute "delete from m_jamsostek where jamsostek_code = " _
                & "'" & TDBCombo_jamsostek.Text & "'"
        
        CnG.CommitTrans
        
        Call load_mode
        Call load_data_jamsostek
        int_mode = 0
        Call load_mode
        
        TDBCombo_jamsostek.Text = ""
        txt_jamsostek_name_type.Text = ""
        Set TDBGrid1.DataSource = Nothing
        
        Exit Sub
Err:
MsgBox Err.Description, vbExclamation, headerMSG

End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & TDBGrid1.Columns("company_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_jamsostek_detail where company_code = '" & _
        TDBGrid1.Columns("company_code").Value & "' " & _
        "AND jamsostek_code = '" & TDBCombo_jamsostek.Text & "'"
CnG.CommitTrans

Call load_data
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
On Error Resume Next
With Adodc1.Recordset
    txt_jkk.Text = .Fields("jkk").Value
    txt_jkm.Text = .Fields("jkm").Value
    txt_prs.Text = .Fields("prs").Value
    txt_tk.Text = .Fields("tk").Value
    txt_jp_employee.Text = .Fields("jp_emp").Value
    txt_jp_company.Text = .Fields("jp_co").Value
    txt_max_jp.Text = FormatNumber(.Fields("max_jp_salary").Value)
    TDBCombo_company.Text = .Fields("company_code")
    txt_company_name.Text = .Fields("company_name")
End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select a.*,b.company_name " _
        & "from m_jamsostek_detail a left join m_company b on a.company_code = b.company_code " _
        & "where jamsostek_code = '" & TDBCombo_jamsostek.Text & "' AND a.company_code = '" _
& Adodc1.Recordset.Fields("company_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_jamsostek_detail where company_code = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdNew_Master_Click()
Dim strsql As String

If CmdNew_Master.Caption = "&New" Then
    CmdNew_Master.Caption = "&Save"
    Call set_buttons_enable(False, False, False, False, True, False, False)
    fra_entry1.Visible = True
    fra_entry.Visible = False
    
    txt_jamsostek_code.Text = ""
    txt_name_jamsostek.Text = ""
    txt_jamsostek_code.SetFocus
    
Else
    strsql = "INSERT INTO m_jamsostek(jamsostek_code,jamsostek_name) " _
            & "VALUES ('" & txt_jamsostek_code & "','" & txt_name_jamsostek & "')"
    CnG.Execute strsql
    
    Call set_buttons_enable(True, False, True, True, False, True, True)
    CmdNew_Master.Caption = "&New"
        
    fra_entry1.Visible = False
    
    txt_jamsostek_code.Text = ""
    txt_name_jamsostek.Text = ""
'    txt_pph_code.SetFocus

    Call load_data_jamsostek
'    Call load_data
End If
End Sub

Private Sub cmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
Dim strsql As String

CnG.BeginTrans

strsql = "INSERT INTO m_jamsostek_detail(jamsostek_code,company_code,jkk,jkm,tk," _
            & "prs,jp_emp,jp_co,max_jp_salary) " _
            & "VALUES " _
            & "('" & TDBCombo_jamsostek.Text & "','" & TDBCombo_company.Text & "','" & Trim(txt_jkk.Text) & "'," _
            & "'" & Trim(txt_jkm.Text) & "','" & Trim(txt_tk.Text) & "'," _
            & "'" & Trim(txt_prs.Text) & "','" & Trim(txt_jp_employee.Text) & "','" & Trim(txt_jp_company.Text) & "'," _
            & "'" & Val(DropAllComma(txt_max_jp.Text)) & "')"
CnG.Execute strsql

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
Dim strsql As String
'On Error GoTo err_capture

CnG.BeginTrans

strsql = "UPDATE m_jamsostek_detail SET company_code = '" & Trim(TDBCombo_company.Text) & "'," _
            & "jkk = '" & Trim(txt_jkk.Text) & "'," _
            & "jkm = '" & Trim(txt_jkm.Text) & "'," _
            & "prs = '" & Trim(txt_prs.Text) & "'," _
            & "tk = '" & Trim(txt_tk.Text) & "'," _
            & "jp_emp = '" & Trim(txt_jp_employee.Text) & "'," _
            & "jp_co = '" & Trim(txt_jp_company.Text) & "'," _
            & "jamsostek_code = '" & TDBCombo_jamsostek.Text & "'," _
            & "max_jp_salary = '" & Val(DropAllComma(txt_max_jp.Text)) & "' " _
            & "WHERE company_code = '" & TDBCombo_company.Text & "' AND jamsostek_code = '" & TDBCombo_jamsostek.Text & "'"
CnG.Execute strsql

CnG.CommitTrans

Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
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

Call load_data
int_mode = 0
Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
CmdNew.Enabled = a And blnUser_Add
CmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
CmdCancel.Enabled = e

CmdPrint.Enabled = F
cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        If Not LCase(Ctr.name) = "txt_jamsostek_name_type" Then Ctr.Text = ""
'    ElseIf TypeOf Ctr Is TDBCombo Then
'        If Not LCase(Ctr.name) = "tdbcombo_jamsostek" Then Ctr.Text = ""
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
TDBCombo_company.Text = ""
txt_company_name.Text = ""
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    TDBCombo_company.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If TDBCombo_company.Enabled = True Then
        TDBCombo_company.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    fra_entry1.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    TDBCombo_company.Enabled = False
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
Adodc_company.ConnectionString = strConn
Adodc_jamsostek.ConnectionString = strConn

Call load_data_jamsostek
Call load_data_company

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
End Sub

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
        str_inc_kode = Right("0000" & Trim(str(CLng(str_inc_kode) + 1)), 5)
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
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
Adodc1.RecordSource = "select a.*,b.company_name " _
        & "from m_jamsostek_detail a left join m_company b on a.company_code = b.company_code " _
        & "where jamsostek_code = '" & TDBCombo_jamsostek.Text & "' order by b.company_code"
Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub Timer1_Timer()
Call load_data_jamsostek
timer1.Enabled = False
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
End If
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_jamsostek()
Adodc_jamsostek.RecordSource = "select * from m_jamsostek order by jamsostek_code"
Adodc_jamsostek.Refresh

TDBCombo_jamsostek.RowSource = Adodc_jamsostek
End Sub

Private Sub TDBCombo_jamsostek_ItemChange()
If TDBCombo_jamsostek.ApproxCount > 0 Then
    TDBCombo_jamsostek.Text = TDBCombo_jamsostek.Columns("jamsostek_code").Value
    txt_jamsostek_name_type = TDBCombo_jamsostek.Columns("jamsostek_name").Value
    
    Call load_data
End If
End Sub

Private Sub txt_jamsostek_code_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_max_jp_Validate(Cancel As Boolean)
    If Not Trim(txt_max_jp.Text) = "" Then
        txt_max_jp.Text = FormatNumber(DropAllComma(txt_max_jp.Text))
    End If
End Sub
