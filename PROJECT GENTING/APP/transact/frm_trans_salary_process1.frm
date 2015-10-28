VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_salary_process 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PPH 21 SALARY PROCESS"
   ClientHeight    =   6690
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_trans_salary_process.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   10335
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   4980
         TabIndex        =   5
         Top             =   750
         Visible         =   0   'False
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Load Data"
         Height          =   645
         Left            =   2490
         Picture         =   "frm_trans_salary_process.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Height          =   645
         Left            =   330
         Picture         =   "frm_trans_salary_process.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   3780
         Picture         =   "frm_trans_salary_process.frx":08A0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_process 
         Caption         =   "&Process Att"
         Height          =   645
         Left            =   1410
         Picture         =   "frm_trans_salary_process.frx":0E2A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_proses_ot 
         Caption         =   "&Process Ot"
         Height          =   645
         Left            =   1410
         Picture         =   "frm_trans_salary_process.frx":13B4
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Please Click New"
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1905
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   180
      TabIndex        =   7
      Top             =   210
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ATTENDANCE"
      TabPicture(0)   =   "frm_trans_salary_process.frx":193E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Adodc1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TDBGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_entry"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "OVERTIME"
      TabPicture(1)   =   "frm_trans_salary_process.frx":195A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Adodc2"
      Tab(1).Control(1)=   "TDBGrid2"
      Tab(1).Control(2)=   "fra_entry1"
      Tab(1).ControlCount=   3
      Begin VB.Frame fra_entry1 
         Height          =   2325
         Left            =   -74820
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   10065
         Begin VB.TextBox txt_company_name1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   285
            Left            =   3030
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   30
            Top             =   390
            Width           =   6195
         End
         Begin MSComCtl2.DTPicker DTPicker_month1 
            Height          =   300
            Left            =   1320
            TabIndex        =   31
            Top             =   810
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   158597123
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from1 
            Height          =   300
            Left            =   4920
            TabIndex        =   32
            Top             =   810
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   158531587
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to1 
            Height          =   300
            Left            =   4920
            TabIndex        =   33
            Top             =   1170
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   158531587
            CurrentDate     =   39278
         End
         Begin MSAdodcLib.Adodc Adodc_company1 
            Height          =   330
            Left            =   1350
            Top             =   60
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
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
         Begin TrueOleDBList60.TDBCombo TDBCombo_company1 
            Height          =   375
            Left            =   1320
            OleObjectBlob   =   "frm_trans_salary_process.frx":1976
            TabIndex        =   34
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   450
            TabIndex        =   38
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE TO"
            Height          =   195
            Left            =   3600
            TabIndex        =   37
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE FROM"
            Height          =   195
            Left            =   3600
            TabIndex        =   36
            Top             =   810
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   630
            TabIndex        =   35
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame fra_entry 
         Height          =   2325
         Left            =   180
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   10065
         Begin VB.TextBox txt_company_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   285
            Left            =   3030
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            Top             =   390
            Width           =   6195
         End
         Begin MSComCtl2.DTPicker DTPicker_month 
            Height          =   300
            Left            =   1320
            TabIndex        =   20
            Top             =   810
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   158597123
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_from 
            Height          =   300
            Left            =   4920
            TabIndex        =   21
            Top             =   810
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   158531587
            CurrentDate     =   39278
         End
         Begin MSComCtl2.DTPicker DTPicker_periode_to 
            Height          =   300
            Left            =   4920
            TabIndex        =   22
            Top             =   1170
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   158531587
            CurrentDate     =   39278
         End
         Begin MSAdodcLib.Adodc Adodc_company 
            Height          =   330
            Left            =   1320
            Top             =   60
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
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
         Begin TrueOleDBList60.TDBCombo TDBCombo_company 
            Height          =   375
            Left            =   1320
            OleObjectBlob   =   "frm_trans_salary_process.frx":3935
            TabIndex        =   23
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label lbl_company 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   450
            TabIndex        =   27
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE TO"
            Height          =   195
            Left            =   3600
            TabIndex        =   26
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PERIODE FROM"
            Height          =   195
            Left            =   3600
            TabIndex        =   25
            Top             =   810
            Width           =   1140
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "MONTH"
            Height          =   195
            Left            =   630
            TabIndex        =   24
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   8
         Top             =   960
         Width           =   8415
         Begin VB.ComboBox cbo_monthly_company 
            Height          =   315
            ItemData        =   "frm_trans_salary_process.frx":58F3
            Left            =   1800
            List            =   "frm_trans_salary_process.frx":58FD
            TabIndex        =   13
            Text            =   "..."
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton cmd_monthly_browse_employee 
            Caption         =   "..."
            Height          =   300
            Left            =   4920
            TabIndex        =   12
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txt_yearly_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5340
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1200
            Width           =   2805
         End
         Begin VB.TextBox txt_yearly_employee_code 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox cbo_yearly_employee 
            Height          =   315
            ItemData        =   "frm_trans_salary_process.frx":5910
            Left            =   1800
            List            =   "frm_trans_salary_process.frx":591A
            TabIndex        =   9
            Text            =   "..."
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_yearly 
            Height          =   300
            Left            =   1800
            TabIndex        =   14
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   158531587
            UpDown          =   -1  'True
            CurrentDate     =   39278
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   1200
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "YEAR"
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "COMPANY"
            Height          =   195
            Left            =   720
            TabIndex        =   15
            Top             =   840
            Width           =   795
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4005
         Left            =   180
         TabIndex        =   28
         Top             =   600
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   7064
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "kd company"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "COMPANY"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "MONTH"
         Columns(2).DataField=   "month_"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "PERIODE FROM"
         Columns(3).DataField=   "periode_from_"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PERIODE TO"
         Columns(4).DataField=   "periode_to_"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=8811"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8731"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2646"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2566"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2646"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2566"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2646"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2566"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
         Caption         =   "LIST OF SALARY PROCESSED"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   0
         Top             =   750
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   4005
         Left            =   -74820
         TabIndex        =   39
         Top             =   600
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   7064
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "kd company"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "COMPANY"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "MONTH"
         Columns(2).DataField=   "month_"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "PERIODE FROM"
         Columns(3).DataField=   "periode_from_"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PERIODE TO"
         Columns(4).DataField=   "periode_to_"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=8811"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8731"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2646"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2566"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2646"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2566"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2646"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2566"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
         Caption         =   "LIST OF SALARY PROCESSED"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   -74970
         Top             =   750
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
   End
End
Attribute VB_Name = "frm_trans_salary_process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim str1 As String, strsql As String
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

'str_sql = "select count(income_code) as rec_count from m_other_income where income_code = '" _
'& Replace$(Trim$(txt_income_code), Chr$(39), Chr$(96)) & "'"
'rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
'
'If rs.Fields("rec_count").Value > 0 Then
'    check_validate_exist_new = True
'    Exit Function
'End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
DTPicker_month.Value = Null
If DTPicker_month.Enabled = True Then DTPicker_month.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not DTPicker_month.Value = Adodc1.Recordset.Fields("month").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'If Trim(txt_income_code) = "" Then
'    MsgBox "Department Code is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_income_code.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
'If Trim(txt_income_name) = "" Then
'    MsgBox "Department Name is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_income_name.SetFocus
'    check_validate_new = False
'    Exit Function
'End If


End Function

Private Sub cmd_refresh_Click()
    Call load_data
End Sub

Private Sub CmdCancel_Click()
    int_mode = 0
    'Call load_mode
End Sub

Private Sub cmd_proses_ot_Click()
Dim rsemployee As New ADODB.Recordset

strsql = "select employee_code, all_in, employee_name,no_jamsostek,npwp from m_employee " & _
        "where company_code = '" & TDBCombo_company1.Text & "'"
    rsemployee.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly


    If rsemployee.RecordCount > 0 Then
        ProgressBar1.Max = rsemployee.RecordCount
        ProgressBar1.Value = 0
        
        Label3.Visible = True
        ProgressBar1.Visible = True
        rsemployee.MoveFirst
        While Not rsemployee.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label3.Caption = "Data ke " & ProgressBar1.Value & " dari " & rsemployee.RecordCount
            
            strsql = "UPDATE h_salary set " _
                & "salary_value = f_get_ot_salary_genting('" & rsemployee!employee_code & "', " _
                & "'" & Format(DTPicker_periode_from1.Value, "yyyy-MM-dd") & "', " _
                & "'" & Format(DTPicker_periode_to1.Value, "yyyy-MM-dd") & "') " _
                & "WHERE month = '" & Format(DTPicker_month1.Value, "yyyy-MM") & "' AND " _
                & "salary_code = 'SU-053'"
            CnG.Execute strsql

            rsemployee.MoveNext
            DoEvents
        Wend
    
    MsgBox "Calculate Overtime Process Success...!!", vbInformation, "Information"
    fra_entry.Visible = False
    End If
    rsemployee.Close
End Sub

Private Sub cmdDelete_Click()
'    Dim i As Integer
'
'    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
'        MsgBox "No Data selected!", vbInformation, headerMSG
'        Exit Sub
'    End If
'
'    i = MsgBox("Are you sure want to delete data '" _
'        & Format(TDBGrid1.Columns("month").Value, "mm-yyyy") & "' ?", vbYesNo + vbQuestion, headerMSG)
'    If Not i = vbYes Then Exit Sub
'
'    CnG.Execute "delete from h_d_salary where left(month,7) = '" & Format(Adodc1.Recordset.Fields("month").Value, "yyyy-mm") & "'"
'    CnG.Execute "delete from h_m_salary where left(month,7) = '" & Format(Adodc1.Recordset.Fields("month").Value, "yyyy-mm") & "'"
    Call load_data
End Sub

Private Sub cmd_new_Click()
If SSTab1.Tab = 0 Then
    fra_entry.Visible = True
    DTPicker_month = Now
    DTPicker_periode_from = Now
    DTPicker_periode_to = Now
Else
    fra_entry1.Visible = True
    DTPicker_month1 = Now
    DTPicker_periode_from1 = Now
    DTPicker_periode_to1 = Now
End If
End Sub

Private Sub cmd_process_Click()
Dim rsemployee As New ADODB.Recordset
Dim d1, d2, dx As Date

d1 = Format(DTPicker_month.Value, "yyyy-MM-01"): dx = DateAdd("m", 1, d1)
d2 = Format(d1, "yyyy-MM-") & Format(DateDiff("d", d1, dx), "0#")

    str1 = "DELETE FROM h_salary WHERE LEFT(MONTH,7) = '" & Format(d1, "yyyy-MM") & "';"
    
    CnG.Execute str1
    
    ProgressBar1.Visible = True
    Label3.Visible = True
    
    strsql = "select employee_code, all_in, employee_name,no_jamsostek,npwp from m_employee " & _
        "where company_code = '" & TDBCombo_company.Text & "'"
    rsemployee.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly


    If rsemployee.RecordCount > 0 Then
        ProgressBar1.Max = rsemployee.RecordCount
        ProgressBar1.Value = 0
        
        Label3.Visible = True
        ProgressBar1.Visible = True
        
        rsemployee.MoveFirst
        While Not rsemployee.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label3.Caption = "Data ke " & ProgressBar1.Value & " dari " & rsemployee.RecordCount
            'Label11.Caption = "( " & rsemployee!employee_code & " ) " & rsemployee!employee_name

            Call proses_su(rsemployee!employee_code, _
                Format(DTPicker_periode_from.Value, "yyyy-MM-dd"), Format(DTPicker_periode_to.Value, "yyyy-MM-dd"), _
                IIf(IsNull(rsemployee!all_in), 0, rsemployee!all_in), IIf(IsNull(rsemployee!no_jamsostek), "", rsemployee!no_jamsostek), _
                IIf(IsNull(rsemployee!npwp), 0, rsemployee!npwp))
            
            '++++++++++++++++++++++++++++++ Update Data Loan +++++++++++++++++++++++++++
            
            strsql = "UPDATE td_loan SET flag_paid = 1 " _
                & "Where employee_code = '" & rsemployee!employee_code & "' " _
                & "AND Month(installment_month) = '" & month(DTPicker_month.Value) & "' " _
                & "AND Year(installment_month) = '" & year(DTPicker_month.Value) & "'"
            CnG.Execute (strsql)
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
            rsemployee.MoveNext
            DoEvents
        Wend
        
        strsql = "UPDATE h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                "Set a.salary_value = 0 " & _
                "WHERE a.month = LEFT('" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "', 7) AND a.salary_code = 'SU-052' AND b.all_in = 1"
        CnG.Execute strsql
        
        strsql = "DELETE FROM h_d_salary WHERE company_code = '" & TDBCombo_company.Text & "' and left(month,7) = '" & Format(DTPicker_month.Value, "yyyy-MM") & "'"
        CnG.Execute strsql
        
        strsql = "INSERT INTO h_d_salary (month,periode_from,periode_to,company_code,company_name) " & _
            "VALUES " & _
            "('" & Format(DTPicker_month.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," & _
            "'" & TDBCombo_company.Text & "','" & Replace(txt_company_name.Text, "'", "''") & "')"
        CnG.Execute strsql
        
        MsgBox "Calculate Salary Process Success...!!", vbInformation, "Information"
        fra_entry.Visible = False
    End If
End Sub

Private Sub process_delete()
CnG.Execute "delete from h_d_salary where left(month,7) = '" & Format(DTPicker_month, "yyyy-mm") & "'"
CnG.Execute "delete from h_m_salary where left(month,7) = '" & Format(DTPicker_month, "yyyy-mm") & "'"
End Sub

Private Sub CmdExit_Click()
Unload Me
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

Private Sub Form_Load()
    Adodc1.ConnectionString = strConn
    Adodc_company.ConnectionString = strConn
    Adodc2.ConnectionString = strConn
    Adodc_company1.ConnectionString = strConn
    
    DTPicker_month.Value = Date
    DTPicker_periode_from.Value = Date
    DTPicker_periode_to.Value = Date
    
    DTPicker_month1.Value = Date
    DTPicker_periode_from1.Value = Date
    DTPicker_periode_to1.Value = Date
    
    SSTab1.Tab = 0
    cmd_proses_ot.Visible = False

    Call load_data
    Call load_data_company
    
    Call load_data_user_access(Me)
    int_mode = 0
    'Call load_mode
 '   timer1.Enabled = True
End Sub

Public Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
TDBCombo_company1.RowSource = Adodc_company
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

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    cmd_process.Visible = True
    cmd_proses_ot.Visible = False
Else
    cmd_process.Visible = False
    cmd_proses_ot.Visible = True
End If
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
'
'    Call load_data_employee
'    Call load_data_department
End If
End Sub

Private Sub TDBCombo_company1_ItemChange()
If TDBCombo_company1.ApproxCount > 0 Then
    TDBCombo_company1.Text = TDBCombo_company1.Columns("company_code").Value
    txt_company_name1 = TDBCombo_company1.Columns("company_name").Value
'
'    Call load_data_employee
'    Call load_data_department
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
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
Adodc1.RecordSource = "select *, cast(left(month,7) as char) as month_, cast(left(periode_from,10) as char) as periode_from_, cast(left(periode_to,10) as char) as periode_to_ " _
& "from h_d_salary order by month asc"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
TDBGrid2.DataSource = Adodc1
End Sub

Private Sub proses_su(pEmployee_code As String, pTgl1 As String, _
    pTgl2 As String, pAll_in As Integer, pnoJamsostek As String, pNpwp As String)
Dim strsql As String
'Dim rsemployee As New ADODB.Recordset
Dim i As Integer
Dim clsCalcSUFormula As New clsCalcSUFormula

        strsql = "DELETE FROM h_salary WHERE flag_type = 'SU' AND LEFT(MONTH,7) = LEFT('" & pTgl2 & "',7) AND employee_code = '" & pEmployee_code & "'"
        CnG.Execute strsql

        strsql = "insert into h_salary " & _
            "(MONTH, employee_code, salary_code, company_code, salary_name," & _
            "date_from, date_to, flag_main_salary, flag_sign,flag_detail," & _
            "flag_use_formula, formula_salary_code, flag_ptkp, ptkp_salary_code, flag_pkp," & _
            "flag_pph21, pph21_number, flag_tax, tax_salary_code, flag_type," & _
            "flag_visible, salary_value, Description) " & _
            "SELECT " & _
                "LEFT('" & pTgl2 & "',7) AS MONTH, '" & pEmployee_code & "', salary_code," & _
                "'" & COMPANY_CODE & "', salary_name, '" & pTgl1 & "', '" & pTgl2 & "', " & _
                "flag_main_salary, flag_sign, flag_detail, flag_use_formula, " & _
                "formula_salary_code, flag_ptkp, ptkp_salary_code, flag_pkp," & _
                "flag_pph21, pph21_number, 0 AS flag_tax, '' AS tax_salary_code," & _
                "'SU' AS flag_type, flag_visible," & _
                "CASE WHEN '" & pAll_in & "' = 1 And salary_code = 'SU-052' " & _
                    "THEN 0 " & _
                    "ELSE f_get_sum_dsum('" & pEmployee_code & "',salary_code,'" & pTgl1 & "','" & pTgl2 & "') END," & _
                "Description " & _
            "FROM m_salary_summary;"
        CnG.Execute strsql

        'strSQl = "CALL sp_calc_su_formula('" & pTgl2 & "','" & pEmployee_code & "');"
        Call clsCalcSUFormula.CalcSuFormula(pTgl2, pEmployee_code, pnoJamsostek, pNpwp)
        'CnG.Execute strSQl
End Sub



