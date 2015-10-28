VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_standard 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER SALARY STANDARD"
   ClientHeight    =   10020
   ClientLeft      =   -15
   ClientTop       =   240
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
   Icon            =   "frm_mst_salary_standard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry 
      Height          =   5445
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   14175
      Begin VB.TextBox txt_jstk_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   65
         Top             =   2520
         Width           =   2835
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_jstk 
         Height          =   375
         Left            =   3960
         OleObjectBlob   =   "frm_mst_salary_standard.frx":000C
         TabIndex        =   63
         Top             =   2520
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker_salary 
         Height          =   315
         Left            =   3960
         TabIndex        =   61
         Top             =   720
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   90374147
         CurrentDate     =   40987
      End
      Begin VB.CheckBox chk_overtime 
         Caption         =   "NO"
         Height          =   195
         Left            =   3960
         TabIndex        =   59
         Top             =   2940
         Width           =   885
      End
      Begin VB.TextBox txt_skill_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2130
         Width           =   1575
      End
      Begin VB.TextBox txt_acting_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1770
         Width           =   1575
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_ptkp 
         Height          =   375
         Left            =   3960
         OleObjectBlob   =   "frm_mst_salary_standard.frx":1FCF
         TabIndex        =   53
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txt_ptkp_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         Top             =   2160
         Width           =   2835
      End
      Begin VB.TextBox txt_pph21_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Top             =   1800
         Width           =   2835
      End
      Begin VB.TextBox txt_employee_code 
         Height          =   315
         Left            =   6000
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txt_driver_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   10
         Top             =   4110
         Width           =   1575
      End
      Begin VB.TextBox txt_jstk_4psn_kawin_bwh 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4020
         TabIndex        =   14
         Top             =   5430
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_special_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         TabIndex        =   11
         Top             =   4470
         Width           =   1575
      End
      Begin VB.TextBox txt_meal_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   8
         Top             =   3390
         Width           =   1575
      End
      Begin VB.TextBox txt_number 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_presence_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3030
         Width           =   1575
      End
      Begin VB.TextBox txt_transport_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2670
         Width           =   1575
      End
      Begin VB.TextBox txt_phone_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3750
         Width           =   1575
      End
      Begin VB.TextBox txt_add_other 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   12
         Top             =   4830
         Width           =   1575
      End
      Begin VB.TextBox txt_functional_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1050
         Width           =   1575
      End
      Begin VB.TextBox txt_staff_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1410
         Width           =   1575
      End
      Begin VB.TextBox txt_main_salary 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   11220
         MaxLength       =   10
         TabIndex        =   1
         Top             =   690
         Width           =   1575
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "..."
         Height          =   320
         Left            =   5460
         TabIndex        =   33
         ToolTipText     =   "Browse employee data..."
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1440
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
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
         Height          =   375
         Left            =   3960
         OleObjectBlob   =   "frm_mst_salary_standard.frx":3F82
         TabIndex        =   50
         Top             =   1800
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_pph 
         Height          =   375
         Left            =   4080
         Top             =   1920
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
      Begin MSAdodcLib.Adodc Adodc_ptkp 
         Height          =   375
         Left            =   4080
         Top             =   2280
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
      Begin MSAdodcLib.Adodc Adodc_jstk 
         Height          =   375
         Left            =   4080
         Top             =   2640
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "TIPE JAMSOSTEK"
         Height          =   195
         Left            =   1740
         TabIndex        =   64
         Top             =   2610
         Width           =   1245
      End
      Begin VB.Label Label16 
         Caption         =   "* yyyy-MM-dd"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5460
         TabIndex        =   62
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL"
         Height          =   195
         Left            =   1740
         TabIndex        =   60
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "OVERTIME"
         Height          =   195
         Left            =   1740
         TabIndex        =   58
         Top             =   2940
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. SKILL"
         Height          =   195
         Left            =   9180
         TabIndex        =   57
         Top             =   2190
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. ACTING"
         Height          =   195
         Left            =   9180
         TabIndex        =   56
         Top             =   1830
         Width           =   1050
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "TIPE PTKP"
         Height          =   195
         Left            =   1740
         TabIndex        =   54
         Top             =   2250
         Width           =   735
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "TIPE PPH21"
         Height          =   195
         Left            =   1740
         TabIndex        =   51
         Top             =   1860
         Width           =   840
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. DRIVER"
         Height          =   195
         Left            =   9180
         TabIndex        =   46
         Top             =   4110
         Width           =   1035
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "JSTK 4,24% MARRIED UNDER 1,1 Jt (Rp)"
         Height          =   195
         Left            =   900
         TabIndex        =   45
         Top             =   5460
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. SPECIAL"
         Height          =   195
         Left            =   9180
         TabIndex        =   44
         Top             =   4500
         Width           =   1095
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. MEAL"
         Height          =   195
         Left            =   9180
         TabIndex        =   43
         Top             =   3390
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "NUMBER"
         Height          =   195
         Left            =   1740
         TabIndex        =   42
         Top             =   330
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. PRESENCE"
         Height          =   195
         Left            =   9180
         TabIndex        =   41
         Top             =   3060
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. TRANSPORT"
         Height          =   195
         Left            =   9180
         TabIndex        =   40
         Top             =   2700
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ.  PHONE"
         Height          =   195
         Left            =   9150
         TabIndex        =   39
         Top             =   3750
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. LAIN"
         Height          =   195
         Left            =   9180
         TabIndex        =   38
         Top             =   4890
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. FUNCTIONAL"
         Height          =   195
         Left            =   9180
         TabIndex        =   37
         Top             =   1110
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TUNJ. STAFF"
         Height          =   195
         Left            =   9180
         TabIndex        =   36
         Top             =   1470
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "UPAH POKOK"
         Height          =   195
         Left            =   9180
         TabIndex        =   34
         Top             =   750
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KODE KARYAWAN"
         Height          =   195
         Left            =   1740
         TabIndex        =   28
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAMA EMPLOYEE"
         Height          =   195
         Left            =   1740
         TabIndex        =   27
         Top             =   1500
         Width           =   1245
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NUMBER"
      Columns(0).DataField=   "number"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMP. CODE"
      Columns(1).DataField=   "employee_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMP. CODE"
      Columns(2).DataField=   "nik"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "EMP. NAME"
      Columns(3).DataField=   "employee_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "MAIN SALARY"
      Columns(4).DataField=   "main_salary"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "FUNCTIONAL"
      Columns(5).DataField=   "functional_allowance"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "STAFF"
      Columns(6).DataField=   "staff_allowance"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "ACTING"
      Columns(7).DataField=   "acting_allowance"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "SKILL"
      Columns(8).DataField=   "skill_allowance"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "TRANSPORT"
      Columns(9).DataField=   "transport_allowance"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "PRESENCE"
      Columns(10).DataField=   "presence_allowance"
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "MEAL"
      Columns(11).DataField=   "meal_allowance"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "PHONE"
      Columns(12).DataField=   "phone_allowance"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "DRIVER"
      Columns(13).DataField=   "driver_allowance"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "SPECIAL"
      Columns(14).DataField=   "special_allowance"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "OTHERS"
      Columns(15).DataField=   "other_allowance"
      Columns(15).NumberFormat=   "Standard"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2117"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2037"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=3519"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3440"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2117"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2037"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2037"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(5)._MinWidth=151003376"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=2117"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2037"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(6)._MinWidth=151012640"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=2117"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(45)=   "Column(7)._MinWidth=150996272"
      Splits(0)._ColumnProps(46)=   "Column(8).Width=2117"
      Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(49)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(8)._MinWidth=150995488"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=2117"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2037"
      Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(57)=   "Column(9)._MinWidth=151085888"
      Splits(0)._ColumnProps(58)=   "Column(10).Width=2117"
      Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=2037"
      Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(63)=   "Column(10)._MinWidth=151074416"
      Splits(0)._ColumnProps(64)=   "Column(11).Width=2117"
      Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=2037"
      Splits(0)._ColumnProps(67)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(68)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(69)=   "Column(11)._MinWidth=-1"
      Splits(0)._ColumnProps(70)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(71)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(73)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(74)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(75)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(76)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(78)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(79)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(80)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(81)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(83)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(85)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(88)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(89)=   "Column(15).Order=16"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      Caption         =   "LIST OF SALARY"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=0,.bgcolor=&H80000002&"
      _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000009&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=4,.fontname=Tahoma"
      _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(27)  =   "Splits(0).Style:id=99,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=116,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(29)  =   ":id=116,.fgcolor=&H80000009&"
      _StyleDefs(30)  =   "Splits(0).HeadingStyle:id=100,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(31)  =   ":id=100,.fgcolor=&H80000002&"
      _StyleDefs(32)  =   "Splits(0).FooterStyle:id=101,.parent=3"
      _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=102,.parent=5"
      _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=104,.parent=6"
      _StyleDefs(35)  =   "Splits(0).EditorStyle:id=103,.parent=7"
      _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=105,.parent=8"
      _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=106,.parent=9"
      _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=115,.parent=10"
      _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=117,.parent=11"
      _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=118,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=99,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=100"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=101"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=103"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=16,.parent=99"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=100"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=101"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=103"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=24,.parent=99"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=100"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=101"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=103"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=20,.parent=99"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=100"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=101"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=103"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=186,.parent=99,.alignment=1,.bold=0,.fontsize=825"
      _StyleDefs(58)  =   ":id=186,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=186,.fontname=Tahoma"
      _StyleDefs(60)  =   "Splits(0).Columns(4).HeadingStyle:id=183,.parent=100,.bold=0,.fontsize=825"
      _StyleDefs(61)  =   ":id=183,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(62)  =   ":id=183,.fontname=Tahoma"
      _StyleDefs(63)  =   "Splits(0).Columns(4).FooterStyle:id=184,.parent=101,.bold=0,.fontsize=825"
      _StyleDefs(64)  =   ":id=184,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(65)  =   ":id=184,.fontname=Tahoma"
      _StyleDefs(66)  =   "Splits(0).Columns(4).EditorStyle:id=185,.parent=103"
      _StyleDefs(67)  =   "Splits(0).Columns(5).Style:id=32,.parent=99,.alignment=1"
      _StyleDefs(68)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=100"
      _StyleDefs(69)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=101"
      _StyleDefs(70)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=103"
      _StyleDefs(71)  =   "Splits(0).Columns(6).Style:id=46,.parent=99,.alignment=1"
      _StyleDefs(72)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=100"
      _StyleDefs(73)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=101"
      _StyleDefs(74)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=103"
      _StyleDefs(75)  =   "Splits(0).Columns(7).Style:id=50,.parent=99,.alignment=1"
      _StyleDefs(76)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=100"
      _StyleDefs(77)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=101"
      _StyleDefs(78)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=103"
      _StyleDefs(79)  =   "Splits(0).Columns(8).Style:id=54,.parent=99,.alignment=1"
      _StyleDefs(80)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=100"
      _StyleDefs(81)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=101"
      _StyleDefs(82)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=103"
      _StyleDefs(83)  =   "Splits(0).Columns(9).Style:id=58,.parent=99,.alignment=1"
      _StyleDefs(84)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=100"
      _StyleDefs(85)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=101"
      _StyleDefs(86)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=103"
      _StyleDefs(87)  =   "Splits(0).Columns(10).Style:id=62,.parent=99,.alignment=1"
      _StyleDefs(88)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=100"
      _StyleDefs(89)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=101"
      _StyleDefs(90)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=103"
      _StyleDefs(91)  =   "Splits(0).Columns(11).Style:id=70,.parent=99,.alignment=1"
      _StyleDefs(92)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=100"
      _StyleDefs(93)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=101"
      _StyleDefs(94)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=103"
      _StyleDefs(95)  =   "Splits(0).Columns(12).Style:id=66,.parent=99"
      _StyleDefs(96)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=100"
      _StyleDefs(97)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=101"
      _StyleDefs(98)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=103"
      _StyleDefs(99)  =   "Splits(0).Columns(13).Style:id=74,.parent=99"
      _StyleDefs(100) =   "Splits(0).Columns(13).HeadingStyle:id=71,.parent=100"
      _StyleDefs(101) =   "Splits(0).Columns(13).FooterStyle:id=72,.parent=101"
      _StyleDefs(102) =   "Splits(0).Columns(13).EditorStyle:id=73,.parent=103"
      _StyleDefs(103) =   "Splits(0).Columns(14).Style:id=78,.parent=99"
      _StyleDefs(104) =   "Splits(0).Columns(14).HeadingStyle:id=75,.parent=100"
      _StyleDefs(105) =   "Splits(0).Columns(14).FooterStyle:id=76,.parent=101"
      _StyleDefs(106) =   "Splits(0).Columns(14).EditorStyle:id=77,.parent=103"
      _StyleDefs(107) =   "Splits(0).Columns(15).Style:id=82,.parent=99"
      _StyleDefs(108) =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=100"
      _StyleDefs(109) =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=101"
      _StyleDefs(110) =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=103"
      _StyleDefs(111) =   "Named:id=33:Normal"
      _StyleDefs(112) =   ":id=33,.parent=0"
      _StyleDefs(113) =   "Named:id=34:Heading"
      _StyleDefs(114) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(115) =   ":id=34,.wraptext=-1"
      _StyleDefs(116) =   "Named:id=35:Footing"
      _StyleDefs(117) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(118) =   "Named:id=36:Selected"
      _StyleDefs(119) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(120) =   "Named:id=37:Caption"
      _StyleDefs(121) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(122) =   "Named:id=38:HighlightRow"
      _StyleDefs(123) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(124) =   "Named:id=39:EvenRow"
      _StyleDefs(125) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(126) =   "Named:id=40:OddRow"
      _StyleDefs(127) =   ":id=40,.parent=33"
      _StyleDefs(128) =   "Named:id=41:RecordSelector"
      _StyleDefs(129) =   ":id=41,.parent=34"
      _StyleDefs(130) =   "Named:id=42:FilterBar"
      _StyleDefs(131) =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7200
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   750
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   30
      Top             =   690
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   24
      Top             =   7950
      Width           =   14175
      Begin VB.CommandButton cmdCekPrint 
         Caption         =   "&Print"
         Height          =   645
         Left            =   8940
         Picture         =   "frm_mst_salary_standard.frx":5F34
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   390
         Width           =   975
      End
      Begin VB.CommandButton cmd_import 
         Caption         =   "&Import"
         Height          =   645
         Left            =   7920
         Picture         =   "frm_mst_salary_standard.frx":64BE
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   390
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Load"
         Height          =   645
         Left            =   12420
         Picture         =   "frm_mst_salary_standard.frx":6A48
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2040
         Picture         =   "frm_mst_salary_standard.frx":6FD2
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5280
         Picture         =   "frm_mst_salary_standard.frx":755C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11400
         Picture         =   "frm_mst_salary_standard.frx":7AE6
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   960
         Picture         =   "frm_mst_salary_standard.frx":8070
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_salary_standard.frx":85FA
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4200
         Picture         =   "frm_mst_salary_standard.frx":8B84
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3120
         Picture         =   "frm_mst_salary_standard.frx":910E
         Style           =   1  'Graphical
         TabIndex        =   17
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_mst_salary_standard.frx":9698
      TabIndex        =   31
      Top             =   690
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1320
      Top             =   840
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
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MASTER SALARY"
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
      Left            =   6210
      TabIndex        =   55
      Top             =   0
      Width           =   2715
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frm_mst_salary_standard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim strsql As String
Dim v_main, v_staff, v_func, v_other, v_phone As Double
Dim v_trans, v_presence, v_main_daily, v_meal As Double
Dim v_driver, v_ot As Double
Dim v_salary_date As String


Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_nik) = "" Then
    MsgBox "Employee Code is empty!", vbOKOnly + vbInformation, headerMSG
    txt_nik.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi employee name
If Trim(txt_employee_name) = "" Then
    MsgBox "Employee Name is empty!", vbOKOnly + vbInformation, headerMSG
    txt_employee_name.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi salary
If Trim(TDBCombo_pph.Text) = "" Then
    MsgBox "PPh21 Type is empty!", vbOKOnly + vbInformation, headerMSG
    TDBCombo_pph.SetFocus
    check_validate_new = False
    Exit Function
End If

''validasi description
'If Trim(txt_description) = "" Then
'    MsgBox "Description is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_description.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
End Function

Private Sub load_data()
timer1.Enabled = True
End Sub

Private Sub chk_overtime_Click()
If chk_overtime.Value = 1 Then
    chk_overtime.Caption = "YES"
Else
    chk_overtime.Caption = "NO"
End If
End Sub

'Private Sub date_event()
'If cbo_date_to.ListIndex = 0 Then
'    DTPicker_date_to.Visible = False
'Else
'    DTPicker_date_to.Visible = True
'    DTPicker_date_to.Value = DTPicker_date_from.Value
'End If
'End Sub
'
'Private Sub cbo_date_to_Click()
'Call date_event
'End Sub

Private Sub cmd_browse_Click()
frm_lookup_mst_employee.public_int_mode = 77
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_import_Click()
frm_import_salary_standard.Show
End Sub

Private Sub cmd_refresh_Click()
Call generate_data_salary
Call load_data_salary
End Sub


Private Sub generate_data_salary()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim str1 As String
Dim i As Integer



i = MsgBox("Are you sure want to generate data main salary?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub


str1 = "select * from m_employee where company_code='" & TDBCombo_company.Columns("company_code").Value _
        & "' and employee_code not in (select distinct employee_code from m_salary_standard)"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

rs2.Open "select * from m_salary_standard where employee_code='uOu'", CnG, adOpenKeyset, adLockOptimistic

While Not rs1.EOF
    rs2.AddNew
    
    rs2.Fields("employee_code").Value = rs1.Fields("employee_code").Value
    rs2.Fields("salary_date").Value = Now 'rs1.Fields("salary_date").Value
    rs2.Fields("salary").Value = 0 'rs1.Fields("salary").Value
    rs2.Fields("over_time").Value = 0 'rs1.Fields("over_time").Value
    rs2.Fields("description").Value = "" 'rs1.Fields("description").Value
    
    rs2.Update
    
    rs1.MoveNext
Wend
End Sub


Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdCekPrint_Click()
Dim str_sql, str_param_periode, str_file, str1, str2  As String
Dim int_flag_company As Integer, str_company_code As String
Dim int_flag_employee As Integer, str_employee_code As String
Dim a As New frm_rpt
Dim int_process As Integer
Dim strsql As String
Dim rsemployee As New ADODB.Recordset
    
int_process = vbNo

str_file = "\report\rpt_salary_standard.rpt"

str_sql = "SELECT (SELECT date_format(salary_date,'%Y-%m-%d') FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) salary_date," & _
            "(SELECT employee_code FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) employee_code," & _
            "b.nik,b.employee_name,(SELECT pph21_type FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) pph21_type," & _
            "(SELECT ptkp_type FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) ptkp_type," & _
            "(SELECT jstk_type FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) jstk_type," & _
            "IFNULL((SELECT flag_ot FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1),0) jstk_type," & _
            "(SELECT main_salary FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) main_salary," & _
            "(SELECT functional_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) functional_allowance," & _
            "(SELECT staff_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) staff_allowance," & _
            "(SELECT acting_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) acting_allowance," & _
            "(SELECT skill_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) skill_allowance," & _
            "(SELECT transport_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) transport_allowance," & _
            "(SELECT presence_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) presence_allowance," & _
            "(SELECT meal_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) meal_allowance," & _
            "(SELECT phone_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) phone_allowance," & _
            "(SELECT driver_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) driver_allowance," & _
            "(SELECT special_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) special_allowance," & _
            "b.COMPANY_CODE , c.company_name, b.DEPARTMENT_CODE, d.department_name " & _
        "FROM m_employee b JOIN m_company c ON b.company_code = c.company_code " & _
            "JOIN m_department d ON b.company_code = d.company_code AND b.department_code = d.department_code " & _
        "WHERE b.flag_active <> 0 AND b.company_code = '" & TDBCombo_company.Text & "' " & _
            "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) ORDER BY 1"

str_param_periode = ""

Text1 = str_sql

Call a.Show

a.Caption = "REPORT SALARY STANDARD"
Call a.rpt_view(str_sql, str_file, str_param_periode)

End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & TDBGrid1.Columns("number").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_salary_standard where number = " & Adodc1.Recordset.Fields("number").Value
'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
CnG.CommitTrans

Call load_data_salary
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With rsBound
    txt_number.Text = .Fields("number").Value
    txt_employee_code.Text = .Fields("employee_code").Value
    txt_nik.Text = .Fields("nik").Value
    txt_employee_name.Text = .Fields("employee_name").Value
    
    txt_main_salary.Text = IIf(IsNull(.Fields("main_salary").Value), 0, .Fields("main_salary").Value)
    
    txt_functional_allowance.Text = IIf(IsNull(.Fields("functional_allowance").Value), 0, .Fields("functional_allowance").Value)
    txt_staff_allowance.Text = IIf(IsNull(.Fields("staff_allowance").Value), 0, .Fields("staff_allowance").Value)
    txt_acting_allowance.Text = IIf(IsNull(.Fields("acting_allowance").Value), 0, .Fields("acting_allowance").Value)
    txt_skill_allowance.Text = IIf(IsNull(.Fields("skill_allowance").Value), 0, .Fields("skill_allowance").Value)
    
    txt_transport_allowance.Text = IIf(IsNull(.Fields("transport_allowance").Value), 0, .Fields("transport_allowance").Value)
    txt_presence_allowance.Text = IIf(IsNull(.Fields("presence_allowance").Value), 0, .Fields("presence_allowance").Value)
    txt_meal_allowance.Text = IIf(IsNull(.Fields("meal_allowance").Value), 0, .Fields("meal_allowance").Value)
    txt_phone_allowance.Text = IIf(IsNull(.Fields("phone_allowance").Value), 0, .Fields("phone_allowance").Value)
    txt_driver_allowance.Text = IIf(IsNull(.Fields("driver_allowance").Value), 0, .Fields("driver_allowance").Value)
    txt_special_allowance.Text = IIf(IsNull(.Fields("special_allowance").Value), 0, .Fields("special_allowance").Value)
    txt_add_other.Text = IIf(IsNull(.Fields("other_allowance").Value), 0, .Fields("other_allowance").Value)
    
    DTPicker_salary.Value = Format(.Fields("salary_date").Value, "yyyy-MM-dd")
    TDBCombo_pph.Text = IIf(IsNull(.Fields("pph21_type").Value), "", .Fields("pph21_type").Value)
    txt_pph21_name.Text = IIf(IsNull(.Fields("pph21_name").Value), "", .Fields("pph21_type").Value)
    TDBCombo_ptkp.Text = IIf(IsNull(.Fields("ptkp_type").Value), "", .Fields("ptkp_type").Value)
    txt_ptkp_name.Text = IIf(IsNull(.Fields("ptkp_name").Value), "", .Fields("ptkp_name").Value)
    TDBCombo_jstk.Text = IIf(IsNull(.Fields("jstk_type").Value), "", .Fields("jstk_type").Value)
    txt_jstk_name.Text = IIf(IsNull(.Fields("jamsostek_name").Value), "", .Fields("jamsostek_name").Value)
    
    chk_overtime.Value = IIf(IsNull(.Fields("flag_ot").Value), 0, .Fields("flag_ot").Value)
    
End With

    v_main = txt_main_salary.Text
    v_staff = txt_staff_allowance.Text
    v_func = txt_functional_allowance.Text
    v_other = txt_add_other.Text
    v_phone = txt_phone_allowance.Text
    v_trans = txt_transport_allowance
    v_presence = txt_presence_allowance
    v_meal = txt_meal_allowance
    v_driver = txt_driver_allowance.Text
    v_salary_date = Format(DTPicker_salary.Value, "yyyy-MM-dd")
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close

If IsNull(Adodc1.Recordset.Fields("number").Value) Then
    Exit Sub
End If

rsBound.Open "select b.nik,b.employee_name,c.pph21_name,d.ptkp_name,e.jamsostek_name, a.* " & _
    "from m_salary_standard a join m_employee b on a.employee_code = b.employee_code " & _
    "left join m_pph21 c on a.pph21_type = c.pph21_code " & _
    "left join m_ptkp d on a.ptkp_type = d.ptkp_code " & _
    "left join m_jamsostek e on a.jstk_type = e.jamsostek_code " & _
    "where number = " & Adodc1.Recordset.Fields("number").Value, CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary_standard where number = -77", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub cmdPrint_Click()
'TDBGrid1.PrintInfo.PageSetup
'If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
'    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
'End If

'++++++++NUMPANG TOMBOL++++++++++++++++
'Dim rsempl As New ADODB.Recordset
'Dim nr As Integer
'
'    strSQL = "Select * from m_employee order by employee_code asc limit 214"
'    rsempl.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    nr = 0
'    rsempl.MoveFirst
'    While Not rsempl.EOF
'        nr = nr + 1
'        strSQL = "update m_salary_standard SET employee_code = '" & rsempl!employee_code & "' " & _
'            " WHERE number = '" & nr & "'"
'        CnG.Execute strSQL
'
'        rsempl.MoveNext
'    Wend
'
'    MsgBox "Update Berhasil"
End Sub

Private Sub insert_new_data()
Dim rsnumber As New ADODB.Recordset
'Dim nourut As Long
'
'strsql = "select ifnull(max(number),0) nourutdb from m_salary_standard"
'rsnumber.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'If rsnumber.RecordCount > 0 Then
'    nourut = rsnumber!nourutdb + 1
'Else
'    nourut = 1
'End If
'rsnumber.Close
CnG.BeginTrans

strsql = "INSERT INTO m_salary_standard (main_salary, staff_allowance," & _
        "functional_allowance, phone_allowance,transport_allowance,other_allowance," & _
        "presence_allowance,meal_allowance,special_allowance,employee_code," & _
        "driver_allowance,entry_date,user_entry,pph21_type,ptkp_type,flag_ot,salary_date," & _
        "acting_allowance,skill_allowance,jstk_type) " & _
        "VALUES " & _
        "(" & Val(txt_main_salary.Text) & "," & Val(txt_staff_allowance.Text) & "," & _
        "" & Val(txt_functional_allowance.Text) & "," & Val(txt_phone_allowance.Text) & "," & _
        "" & Val(txt_transport_allowance.Text) & "," & Val(txt_add_other.Text) & "," & _
        "" & Val(txt_presence_allowance.Text) & "," & Val(txt_meal_allowance.Text) & "," & _
        "" & Val(txt_special_allowance.Text) & ",'" & txt_employee_code.Text & "'," & _
        "" & Val(txt_driver_allowance.Text) & ", Now(), '" & LOGIN_CODE & "'," & _
        "'" & TDBCombo_pph.Text & "','" & TDBCombo_ptkp.Text & "'," & chk_overtime.Value & "," & _
        "'" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & Val(txt_acting_allowance.Text) & "," & _
        "" & Val(txt_skill_allowance.Text) & ",'" & TDBCombo_jstk.Text & "')"
    CnG.Execute strsql

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
CnG.CommitTrans

End Sub

Private Sub edit_old_data()
'On Error GoTo err_capture

CnG.BeginTrans

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
If v_main <> txt_main_salary.Text Or v_staff <> txt_staff_allowance.Text Or v_func <> txt_functional_allowance.Text _
    Or v_other <> txt_add_other.Text Or v_phone <> txt_phone_allowance.Text Or v_trans <> txt_transport_allowance _
    Or v_presence <> txt_presence_allowance Or v_meal <> txt_meal_allowance _
    Or v_driver <> txt_driver_allowance.Text Then
        
        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute strsql
End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
strsql = "UPDATE m_salary_standard SET " & _
    "employee_code = '" & txt_employee_code.Text & "'," & _
    "main_salary = '" & txt_main_salary.Text & "'," & _
    "staff_allowance = '" & txt_staff_allowance.Text & "'," & _
    "functional_allowance = '" & txt_functional_allowance.Text & "'," & _
    "phone_allowance = '" & txt_phone_allowance.Text & "'," & _
    "driver_allowance = '" & txt_driver_allowance.Text & "'," & _
    "transport_allowance = '" & txt_transport_allowance.Text & "'," & _
    "presence_allowance = '" & txt_presence_allowance.Text & "'," & _
    "other_allowance = '" & txt_add_other.Text & "'," & _
    "meal_allowance = '" & txt_meal_allowance.Text & "'," & _
    "special_allowance = '" & txt_special_allowance.Text & "'," & _
    "acting_allowance = '" & txt_acting_allowance.Text & "'," & _
    "skill_allowance = '" & txt_skill_allowance.Text & "'," & _
    "pph21_type = '" & TDBCombo_pph.Text & "'," & _
    "ptkp_type = '" & TDBCombo_ptkp.Text & "'," & _
    "jstk_type = '" & TDBCombo_jstk.Text & "'," & _
    "flag_ot = '" & chk_overtime.Value & "'," & _
    "salary_date = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "' " & _
    "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
        "AND date(salary_date) = '" & v_salary_date & "'"
CnG.Execute strsql

CnG.CommitTrans

Exit Sub
'err_capture:
'MsgBox err.Description
'rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub CmdSave_Click()
Dim clsFunc As New clsFunction

If int_mode = 1 Then
    If Not check_validate_new Then Exit Sub
    If check_validate_exist_new Then
        MsgBox "No valid data!", vbInformation, headerMSG
        Exit Sub
    End If
    Call insert_new_data
    clsFunc.InsertLog ("Insert Salary Standard : " & txt_employee_code.Text)
ElseIf int_mode = 2 Then
    If Not check_validate_new Then Exit Sub
'    If check_validate_exist_edit Then
'        Call check_invalid: Exit Sub
'    End If
    Call edit_old_data
    'Call insert_new_data
    clsFunc.InsertLog ("Edit Salary Standard : " & txt_employee_code.Text)
End If

Call load_data_salary
int_mode = 0
Call load_mode
End Sub

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(number) as rec_count from m_salary_standard where number = " _
& Val(txt_number)
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function


Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
cmdNew.Enabled = a And blnUser_Add
cmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
cmdCancel.Enabled = e

cmdPrint.Enabled = F
cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
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
txt_employee_code = ""
txt_employee_name = ""
DTPicker_salary.Value = Now
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    'txt_number.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_number.Enabled = True Then
        txt_number.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    txt_number.Enabled = False
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
Adodc_pph.ConnectionString = strConn
Adodc_ptkp.ConnectionString = strConn
Adodc_jstk.ConnectionString = strConn

Call load_data_company

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
timer1.Enabled = True

'cmdEdit.Enabled = False
'cmdDelete.Enabled = False
'CmdNew.Enabled = False

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub txt_driver_allowance_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 44, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select

End Sub

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

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call load_data_pph21
    Call load_data_ptkp
    Call load_data_pph21
    Call load_data_jstk
    Call load_data_salary
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

Private Sub load_data_salary()
'Adodc1.RecordSource = "select * from m_salary_standard_standard where company_code = '" _
'& TDBCombo_company.Columns("company_code").Value & "' order by salary_date desc"

'strSQL = "select b.employee_name,a.* " & _
'    "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
'    "WHERE b.level_code >= '" & DATA_LEVEL & "' order by 1"
    
strsql = "SELECT b.*," & _
            "(SELECT number FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) number," & _
            "(SELECT main_salary FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) main_salary," & _
            "(SELECT functional_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) functional_allowance," & _
            "(SELECT staff_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) staff_allowance," & _
            "(SELECT acting_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) acting_allowance," & _
            "(SELECT skill_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) skill_allowance," & _
            "(SELECT transport_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) transport_allowance," & _
            "(SELECT presence_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) presence_allowance," & _
            "(SELECT meal_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) meal_allowance," & _
            "(SELECT phone_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) phone_allowance," & _
            "(SELECT driver_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) driver_allowance," & _
            "(SELECT special_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) special_allowance," & _
            "(SELECT other_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) other_allowance " & _
        "FROM m_employee b WHERE b.flag_active <> 0 AND b.company_code = '" & TDBCombo_company.Text & "' AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) ORDER BY 1"
Adodc1.RecordSource = strsql

Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'CmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "DATE" Then
    Value = Format(Value, "yyyy-mm-dd")
End If
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_pph21()
Adodc_pph.RecordSource = "select * from m_pph21 order by pph21_code"
Adodc_pph.Refresh

TDBCombo_pph.RowSource = Adodc_pph
End Sub

Private Sub TDBCombo_pph_ItemChange()
If TDBCombo_pph.ApproxCount > 0 Then
    TDBCombo_pph.Text = TDBCombo_pph.Columns("pph21_code").Value
    txt_pph21_name = TDBCombo_pph.Columns("pph21_name").Value
    
End If
End Sub

Private Sub load_data_ptkp()
Adodc_ptkp.RecordSource = "select * from m_ptkp order by ptkp_code"
Adodc_ptkp.Refresh

TDBCombo_ptkp.RowSource = Adodc_ptkp
End Sub

Private Sub TDBCombo_ptkp_ItemChange()
If TDBCombo_ptkp.ApproxCount > 0 Then
    TDBCombo_ptkp.Text = TDBCombo_ptkp.Columns("ptkp_code").Value
    txt_ptkp_name = TDBCombo_ptkp.Columns("ptkp_name").Value
    
End If
End Sub

Private Sub load_data_jstk()
Adodc_jstk.RecordSource = "select * from m_jamsostek order by jamsostek_code"
Adodc_jstk.Refresh

TDBCombo_jstk.RowSource = Adodc_jstk
End Sub

Private Sub TDBCombo_jstk_ItemChange()
If TDBCombo_jstk.ApproxCount > 0 Then
    TDBCombo_jstk.Text = TDBCombo_jstk.Columns("jamsostek_code").Value
    txt_jstk_name = TDBCombo_jstk.Columns("jamsostek_name").Value
    
End If
End Sub

Private Sub txt_salary_Validate(Cancel As Boolean)
'If Not Trim(txt_salary) = "" Then
'    txt_salary = FormatNumber(DropAllComma(txt_salary))
'End If
End Sub
