VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_spl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OVERTIME FORM"
   ClientHeight    =   9405
   ClientLeft      =   -15
   ClientTop       =   300
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
   Icon            =   "frm_trans_spl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   3255
      Left            =   2160
      TabIndex        =   39
      Top             =   5070
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
   Begin VB.Frame fra_status_emp 
      Height          =   585
      Left            =   11010
      TabIndex        =   52
      Top             =   1230
      Width           =   3405
      Begin VB.OptionButton optApproval 
         Caption         =   "APPROVE"
         Height          =   225
         Index           =   1
         Left            =   1890
         TabIndex        =   54
         Top             =   240
         Width           =   1245
      End
      Begin VB.OptionButton optApproval 
         Caption         =   "NOT APPROVE"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   1365
      End
   End
   Begin MSComCtl2.DTPicker DTPicker_from 
      Height          =   315
      Left            =   1200
      TabIndex        =   47
      Top             =   1590
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94502915
      CurrentDate     =   41332
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2670
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   870
      Width           =   3855
   End
   Begin VB.Frame fra_entry 
      Height          =   3465
      Left            =   240
      TabIndex        =   16
      Top             =   4290
      Width           =   14175
      Begin VB.CheckBox chkChangeStatus 
         Caption         =   "Change Status Present to OT"
         Height          =   195
         Left            =   7860
         TabIndex        =   67
         Top             =   1530
         Width           =   2445
      End
      Begin VB.CheckBox chkShift3 
         Caption         =   "GET NIGHT ALLOW."
         Height          =   285
         Left            =   11100
         TabIndex        =   8
         Top             =   2430
         Width           =   2025
      End
      Begin VB.CheckBox chkShift2 
         Caption         =   "GET AFTERNOON ALLOW."
         Height          =   285
         Left            =   8850
         TabIndex        =   7
         Top             =   2430
         Width           =   2475
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frm_trans_spl.frx":058A
         Left            =   9480
         List            =   "frm_trans_spl.frx":058C
         TabIndex        =   64
         Text            =   "..."
         Top             =   420
         Width           =   645
      End
      Begin VB.CheckBox chkOnCall 
         Caption         =   "ON CALL"
         Height          =   285
         Left            =   7860
         TabIndex        =   6
         Top             =   2430
         Width           =   1275
      End
      Begin VB.TextBox txtTransport 
         Height          =   315
         Left            =   10980
         TabIndex        =   5
         Top             =   2100
         Width           =   435
      End
      Begin VB.CheckBox chkManual 
         Caption         =   "MANUAL"
         Height          =   195
         Left            =   11820
         TabIndex        =   58
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox txtMeal 
         Height          =   315
         Left            =   8280
         TabIndex        =   4
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox txtBreak 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11520
         TabIndex        =   56
         Text            =   "1"
         Top             =   1170
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkBreak 
         Caption         =   "Inc. Break"
         Height          =   195
         Left            =   10410
         TabIndex        =   55
         Top             =   1230
         Width           =   1125
      End
      Begin VB.TextBox txt_ot_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         Top             =   810
         Width           =   3975
      End
      Begin VB.TextBox txt6 
         Height          =   315
         Left            =   11850
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Top             =   450
         Width           =   1335
      End
      Begin TDBTime6Ctl.TDBTime TDBTime1 
         Height          =   315
         Left            =   7860
         TabIndex        =   2
         Top             =   1170
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "frm_trans_spl.frx":058E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_trans_spl.frx":05F3
         Spin            =   "frm_trans_spl.frx":0643
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "14:36"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.608530092592593
      End
      Begin VB.TextBox txt4 
         Height          =   315
         Left            =   10980
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox txt3 
         Height          =   315
         Left            =   10050
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox txt2 
         Height          =   315
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox txt15 
         Height          =   315
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1770
         Width           =   435
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   7860
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2730
         Width           =   5475
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin TDBTime6Ctl.TDBTime TDBTime2 
         Height          =   315
         Left            =   9300
         TabIndex        =   3
         Top             =   1170
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "frm_trans_spl.frx":066B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_trans_spl.frx":06D0
         Spin            =   "frm_trans_spl.frx":0720
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   1
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "14:36"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.608530092592593
      End
      Begin VB.TextBox txt_employee_code 
         Height          =   315
         Left            =   3720
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin prj_tpc.vbButton cmdBrowse 
         Height          =   315
         Left            =   3300
         TabIndex        =   40
         Top             =   450
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
         MICON           =   "frm_trans_spl.frx":0748
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_ot 
         Height          =   375
         Left            =   7860
         OleObjectBlob   =   "frm_trans_spl.frx":0764
         TabIndex        =   1
         Top             =   810
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   7860
         TabIndex        =   65
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   94502915
         CurrentDate     =   40823
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   10200
         TabIndex        =   66
         Top             =   420
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   94502915
         CurrentDate     =   40823
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "TRANSPORT"
         Height          =   195
         Left            =   9960
         TabIndex        =   61
         Top             =   2130
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "DAYS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   11430
         TabIndex        =   60
         Top             =   2190
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "DAYS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8760
         TabIndex        =   59
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "MEAL"
         Height          =   195
         Left            =   7860
         TabIndex        =   57
         Top             =   2130
         Width           =   390
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TYPE OT*"
         Height          =   195
         Left            =   6270
         TabIndex        =   46
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "6 X"
         Height          =   195
         Left            =   11520
         TabIndex        =   44
         Top             =   1830
         Width           =   225
      End
      Begin VB.Label lblVerify 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         TabIndex        =   42
         Top             =   1560
         Width           =   5265
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   195
         Left            =   8970
         TabIndex        =   30
         Top             =   1200
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "4 X"
         Height          =   195
         Left            =   10650
         TabIndex        =   29
         Top             =   1830
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "3 X"
         Height          =   195
         Left            =   9750
         TabIndex        =   28
         Top             =   1830
         Width           =   225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "2 X"
         Height          =   195
         Left            =   8850
         TabIndex        =   27
         Top             =   1830
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "1.5 X"
         Height          =   195
         Left            =   7860
         TabIndex        =   26
         Top             =   1830
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "START | END"
         Height          =   195
         Left            =   6300
         TabIndex        =   25
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DATE"
         Height          =   195
         Left            =   6300
         TabIndex        =   24
         Top             =   450
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   6300
         TabIndex        =   21
         Top             =   2790
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EMP. CODE"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EMP. NAME"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   840
         Width           =   825
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   7890
      Width           =   14175
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   660
         TabIndex        =   32
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
         MICON           =   "frm_trans_spl.frx":26AD
         PICN            =   "frm_trans_spl.frx":26C9
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
         Left            =   1680
         TabIndex        =   33
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
         MICON           =   "frm_trans_spl.frx":375B
         PICN            =   "frm_trans_spl.frx":3777
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
         Left            =   2700
         TabIndex        =   34
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
         MICON           =   "frm_trans_spl.frx":4809
         PICN            =   "frm_trans_spl.frx":4825
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
         Left            =   3720
         TabIndex        =   35
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
         MICON           =   "frm_trans_spl.frx":58B7
         PICN            =   "frm_trans_spl.frx":58D3
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
         Left            =   4740
         TabIndex        =   36
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
         MICON           =   "frm_trans_spl.frx":6965
         PICN            =   "frm_trans_spl.frx":6981
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
         Left            =   12750
         TabIndex        =   37
         Top             =   390
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
         MICON           =   "frm_trans_spl.frx":7A13
         PICN            =   "frm_trans_spl.frx":7A2F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton btnApprove 
         Height          =   705
         Left            =   7890
         TabIndex        =   41
         Top             =   390
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Approve"
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
         MICON           =   "frm_trans_spl.frx":8AC1
         PICN            =   "frm_trans_spl.frx":8ADD
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
      Height          =   5715
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10081
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "EMPLOYEE CODE"
      Columns(0).DataField=   "employee_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMP. CODE"
      Columns(1).DataField=   "nik"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NAME"
      Columns(2).DataField=   "employee_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TITLE CODE"
      Columns(3).DataField=   "title_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TITLE NAME"
      Columns(4).DataField=   "title_name"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DATE"
      Columns(5).DataField=   "date_"
      Columns(5).NumberFormat=   "dd-MM-yyyy"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "TYPE OT"
      Columns(6).DataField=   "ot_code"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "TYPE"
      Columns(7).DataField=   "type"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "START"
      Columns(8).DataField=   "start_time_"
      Columns(8).NumberFormat=   "HH:mm"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   4
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "BREAK"
      Columns(9).DataField=   "flag_break"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "END"
      Columns(10).DataField=   "end_time_"
      Columns(10).NumberFormat=   "HH:mm"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   4
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "ON CALL"
      Columns(11).DataField=   "flag_on_call"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   4
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "AFTERNOON"
      Columns(12).DataField=   "flag_shift2"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   4
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "NIGHT"
      Columns(13).DataField=   "flag_shift3"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   4
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "APPROVE"
      Columns(14).DataField=   "flag_approval"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "USER APPROVAL"
      Columns(15).DataField=   "user_approval"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "DESCRIPTION"
      Columns(16).DataField=   "description"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "DATETIME"
      Columns(17).DataField=   "date"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "SEQ"
      Columns(18).DataField=   "seq"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "CHANGE STATUS OT"
      Columns(19).DataField=   "flag_change_status"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   20
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=20"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2170"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2090"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=4498"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=4419"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=2355"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2143"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2064"
      Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(48)=   "Column(8).Width=1984"
      Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=1905"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=513"
      Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=1111"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=1032"
      Splits(0)._ColumnProps(56)=   "Column(9)._ColStyle=513"
      Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(58)=   "Column(10).Width=1931"
      Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=1852"
      Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(63)=   "Column(11).Width=1455"
      Splits(0)._ColumnProps(64)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(11)._WidthInPix=1376"
      Splits(0)._ColumnProps(66)=   "Column(11)._ColStyle=513"
      Splits(0)._ColumnProps(67)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(68)=   "Column(12).Width=1826"
      Splits(0)._ColumnProps(69)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(12)._WidthInPix=1746"
      Splits(0)._ColumnProps(71)=   "Column(12)._ColStyle=513"
      Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(73)=   "Column(13).Width=1402"
      Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=1323"
      Splits(0)._ColumnProps(76)=   "Column(13)._ColStyle=513"
      Splits(0)._ColumnProps(77)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(78)=   "Column(14).Width=1508"
      Splits(0)._ColumnProps(79)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(14)._WidthInPix=1429"
      Splits(0)._ColumnProps(81)=   "Column(14)._ColStyle=513"
      Splits(0)._ColumnProps(82)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(83)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(87)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(89)=   "Column(16).Width=5371"
      Splits(0)._ColumnProps(90)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(16)._WidthInPix=5292"
      Splits(0)._ColumnProps(92)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(93)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(94)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(95)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(97)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(98)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(99)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(100)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(101)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(103)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(104)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(105)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(106)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(107)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(108)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(109)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(110)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(111)=   "Column(19).Order=20"
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
      Caption         =   "LIST OF OVER TIME"
      TabAction       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   2
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
      _StyleDefs(21)  =   "Splits(0).Style:id=99,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=116,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(23)  =   ":id=116,.fgcolor=&H80000009&"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=100,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(25)  =   ":id=100,.fgcolor=&H80000002&"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=101,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=102,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=104,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=103,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=105,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=106,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=115,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=117,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=118,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=16,.parent=99"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=100"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=101"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=103"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=99,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=100"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=101"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=103"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=28,.parent=99"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=100"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=101"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=103"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=154,.parent=99"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=151,.parent=100"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=152,.parent=101"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=153,.parent=103"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=158,.parent=99"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=155,.parent=100"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=156,.parent=101"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=157,.parent=103"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=46,.parent=99,.alignment=2"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=100"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=101"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=103"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=70,.parent=99,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=100"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=101"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=103"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=66,.parent=99,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=100"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=101"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=103"
      _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=50,.parent=99,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=100"
      _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=101"
      _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=103"
      _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=54,.parent=99,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=100"
      _StyleDefs(73)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=101"
      _StyleDefs(74)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=103"
      _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=58,.parent=99,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=100"
      _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=101"
      _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=103"
      _StyleDefs(79)  =   "Splits(0).Columns(11).Style:id=74,.parent=99,.alignment=2"
      _StyleDefs(80)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=100"
      _StyleDefs(81)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=101"
      _StyleDefs(82)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=103"
      _StyleDefs(83)  =   "Splits(0).Columns(12).Style:id=90,.parent=99,.alignment=2"
      _StyleDefs(84)  =   "Splits(0).Columns(12).HeadingStyle:id=87,.parent=100"
      _StyleDefs(85)  =   "Splits(0).Columns(12).FooterStyle:id=88,.parent=101"
      _StyleDefs(86)  =   "Splits(0).Columns(12).EditorStyle:id=89,.parent=103"
      _StyleDefs(87)  =   "Splits(0).Columns(13).Style:id=94,.parent=99,.alignment=2"
      _StyleDefs(88)  =   "Splits(0).Columns(13).HeadingStyle:id=91,.parent=100"
      _StyleDefs(89)  =   "Splits(0).Columns(13).FooterStyle:id=92,.parent=101"
      _StyleDefs(90)  =   "Splits(0).Columns(13).EditorStyle:id=93,.parent=103"
      _StyleDefs(91)  =   "Splits(0).Columns(14).Style:id=20,.parent=99,.alignment=2"
      _StyleDefs(92)  =   "Splits(0).Columns(14).HeadingStyle:id=17,.parent=100"
      _StyleDefs(93)  =   "Splits(0).Columns(14).FooterStyle:id=18,.parent=101"
      _StyleDefs(94)  =   "Splits(0).Columns(14).EditorStyle:id=19,.parent=103"
      _StyleDefs(95)  =   "Splits(0).Columns(15).Style:id=24,.parent=99"
      _StyleDefs(96)  =   "Splits(0).Columns(15).HeadingStyle:id=21,.parent=100"
      _StyleDefs(97)  =   "Splits(0).Columns(15).FooterStyle:id=22,.parent=101"
      _StyleDefs(98)  =   "Splits(0).Columns(15).EditorStyle:id=23,.parent=103"
      _StyleDefs(99)  =   "Splits(0).Columns(16).Style:id=62,.parent=99"
      _StyleDefs(100) =   "Splits(0).Columns(16).HeadingStyle:id=59,.parent=100"
      _StyleDefs(101) =   "Splits(0).Columns(16).FooterStyle:id=60,.parent=101"
      _StyleDefs(102) =   "Splits(0).Columns(16).EditorStyle:id=61,.parent=103"
      _StyleDefs(103) =   "Splits(0).Columns(17).Style:id=78,.parent=99"
      _StyleDefs(104) =   "Splits(0).Columns(17).HeadingStyle:id=75,.parent=100"
      _StyleDefs(105) =   "Splits(0).Columns(17).FooterStyle:id=76,.parent=101"
      _StyleDefs(106) =   "Splits(0).Columns(17).EditorStyle:id=77,.parent=103"
      _StyleDefs(107) =   "Splits(0).Columns(18).Style:id=82,.parent=99"
      _StyleDefs(108) =   "Splits(0).Columns(18).HeadingStyle:id=79,.parent=100"
      _StyleDefs(109) =   "Splits(0).Columns(18).FooterStyle:id=80,.parent=101"
      _StyleDefs(110) =   "Splits(0).Columns(18).EditorStyle:id=81,.parent=103"
      _StyleDefs(111) =   "Splits(0).Columns(19).Style:id=86,.parent=99"
      _StyleDefs(112) =   "Splits(0).Columns(19).HeadingStyle:id=83,.parent=100"
      _StyleDefs(113) =   "Splits(0).Columns(19).FooterStyle:id=84,.parent=101"
      _StyleDefs(114) =   "Splits(0).Columns(19).EditorStyle:id=85,.parent=103"
      _StyleDefs(115) =   "Named:id=33:Normal"
      _StyleDefs(116) =   ":id=33,.parent=0"
      _StyleDefs(117) =   "Named:id=34:Heading"
      _StyleDefs(118) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(119) =   ":id=34,.wraptext=-1"
      _StyleDefs(120) =   "Named:id=35:Footing"
      _StyleDefs(121) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(122) =   "Named:id=36:Selected"
      _StyleDefs(123) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(124) =   "Named:id=37:Caption"
      _StyleDefs(125) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(126) =   "Named:id=38:HighlightRow"
      _StyleDefs(127) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(128) =   "Named:id=39:EvenRow"
      _StyleDefs(129) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(130) =   "Named:id=40:OddRow"
      _StyleDefs(131) =   ":id=40,.parent=33"
      _StyleDefs(132) =   "Named:id=41:RecordSelector"
      _StyleDefs(133) =   ":id=41,.parent=34"
      _StyleDefs(134) =   "Named:id=42:FilterBar"
      _StyleDefs(135) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_trans_spl.frx":9B6F
      TabIndex        =   15
      Top             =   870
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPicker_to 
      Height          =   315
      Left            =   2910
      TabIndex        =   49
      Top             =   1590
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94502915
      CurrentDate     =   41332
   End
   Begin prj_tpc.vbButton cmdSearch 
      Height          =   465
      Left            =   4320
      TabIndex        =   51
      Top             =   1530
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&View"
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
      MICON           =   "frm_trans_spl.frx":BAD5
      PICN            =   "frm_trans_spl.frx":BAF1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker_Periode 
      Height          =   315
      Left            =   1200
      TabIndex        =   62
      Top             =   1230
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM-yyyy"
      Format          =   94502915
      CurrentDate     =   40794
   End
   Begin VB.Label Label21 
      Caption         =   "PERIODE"
      Height          =   195
      Left            =   240
      TabIndex        =   63
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "DATE"
      Height          =   195
      Left            =   240
      TabIndex        =   50
      Top             =   1620
      Width           =   390
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   48
      Top             =   1620
      Width           =   285
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "OVERTIME FORM"
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
      TabIndex        =   38
      Top             =   150
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   930
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_spl.frx":CB83
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14760
   End
End
Attribute VB_Name = "frm_trans_spl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCompany As New ADODB.Recordset
Dim rsSPL As New ADODB.Recordset
Dim rsot As New ADODB.Recordset

Public int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim jmlJam As Double
Dim vHoliday As Integer
Dim vTimeIn As String
Dim vTimeOut As String
Dim vParam As String

Dim vOTMax_Meal As Double
Dim vFlagMeal As Integer
Dim vOTMax_Trans As Double
Dim vFlagTrans As Integer

Dim vShiftCode As String
Dim vGroupCode As String

Dim tglAwal, tglAkhir As Date
Dim vEndTime, vStartTime As String
Dim vSelisihJam As Double
Dim vChangeStatusHours As Double
Dim vSeq As Integer

Private Function check_validate_exist_new() As Boolean
    check_validate_exist_new = False
    
    If rs.State Then rs.Close
    SQL = "select count(employee_code) as rec_count from t_spl where employee_code = '" & txt_employee_code.Text & "' " & _
                "and left(date,10)= '" & Format(DTPicker1.Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
                "and company_code= '" & TDBCombo_company.Text & "'"
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly

    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        rs.Close
        Exit Function
    End If

    rs.Close
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    txt_nik = ""
    txt_employee_code.Text = ""
    txt_employee_name.Text = ""
    If txt_nik.Enabled = True Then txt_nik.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not txt_employee_code = rsSPL.Fields("employee_code").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    'validasi employee code
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
    
    'validasi employee name
    If Trim(TDBCombo_ot.Text) = "" Then
        MsgBox "Type OT is not selected...", vbOKOnly + vbInformation, headerMSG
        TDBCombo_ot.SetFocus
        check_validate_new = False
        Exit Function
    End If

End Function

Private Sub load_data()
    timer1.Enabled = True
End Sub

Private Sub btnApprove_Click()
Dim i As Integer
Dim item
    
On Error GoTo Err
    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid1.SelBookmarks
    i = MsgBox("Are you sure want to approve " _
        & SelBks.Count & " overtime's data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        SQL = "UPDATE t_spl SET flag_approval = 1," & _
                "user_approval = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "and date= '" & Format(TDBGrid1.Columns("date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "' " & _
                    "and company_code= '" & TDBCombo_company.Text & "' " & _
                    "and seq = " & TDBGrid1.Columns("seq").CellText(item) & ""
        CnG.Execute SQL
        
        If TDBGrid1.Columns("flag_change_status").CellText(item) <> 0 Then
            SQL = "UPDATE h_attendance SET status = 'OT', flag_manual = 1 " & _
                  "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "and att_date= '" & Format(TDBGrid1.Columns("date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "'"
            CnG.Execute SQL
        End If
                
    Next
    CnG.CommitTrans
    Call load_data_spl
    MsgBox i & " overtime's data are successfully approved", vbInformation, headerMSG
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmd_browse_Click()
    frm_lookup_mst_employee.public_int_mode = 161
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_refresh_Click()
    Call load_data_spl
End Sub

Private Sub chkBreak_Click()
    If chkBreak.Value = 0 Then
        txtBreak.Visible = False
        txtBreak.Text = 0
    Else
        txtBreak.Visible = True
        txtBreak.Text = 1
'        txtBreak.SetFocus
    End If
End Sub

Private Sub chkManual_Click()
    If chkManual.Value = 0 Then
        DTPicker1.CustomFormat = "yyyy-MM-dd"
        
        TDBTime1.Enabled = True
        TDBTime2.Enabled = True
        
        txt15.Locked = True
        txt2.Locked = True
        txt3.Locked = True
        txt4.Locked = True
        txt6.Locked = True
        txtMeal.Locked = True
        txtTransport.Locked = False
    Else
        DTPicker1.CustomFormat = "yyyy-MM"
        
        TDBTime1.Enabled = False
        TDBTime2.Enabled = False
        
        txt15.Locked = False
        txt2.Locked = False
        txt3.Locked = False
        txt4.Locked = False
        txt6.Locked = False
        txtMeal.Locked = False
        txtTransport.Locked = False
    End If
End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
    
    btnApprove.Visible = True
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim item
Dim vSelisihJam As Double
    
On Error GoTo Err
    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid1.SelBookmarks
    i = MsgBox("Are you sure want to delete " _
        & SelBks.Count & " overtime's data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        If TDBGrid1.Columns("ot_code").CellText(item) <> "HK" Then
            SQL = "UPDATE h_attendance SET status = 'P' " & _
                  "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "AND att_date = '" & Format(TDBGrid1.Columns("date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "'"
            CnG.Execute SQL
        End If
        
        vTimeIn = Format(TDBGrid1.Columns("start_time_").CellText(item), "HH:mm")
        vTimeOut = Format(TDBGrid1.Columns("end_time_").CellText(item), "HH:mm")
        
        vStartTime = Format(TDBGrid1.Columns("start_time_").CellText(item), "yyyy-MM-dd HH:mm:ss")
        If vTimeOut < vTimeIn Then
            vEndTime = Format(DateAdd("d", 1, TDBGrid1.Columns("end_time_").CellText(item)), "yyyy-MM-dd HH:mm:ss")
        Else
            vEndTime = Format(TDBGrid1.Columns("end_time_").CellText(item), "yyyy-MM-dd HH:mm:ss")
        End If

'        vStartTime = Format(TDBGrid1.Columns("start_time_").CellText(item), "yyyy-mm-dd hh:mm:ss")
'        vEndTime = Format(TDBGrid1.Columns("end_time_").CellText(item), "yyyy-mm-dd hh:mm:ss")
        
        vSelisihJam = DateDiff("s", vStartTime, vEndTime)
        vSelisihJam = vSelisihJam / 3600
        
        If vSelisihJam >= 7.9 Then
            SQL = "UPDATE h_attendance SET status = 'P' " & _
                  "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "and att_date= '" & Format(TDBGrid1.Columns("date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "'"
            CnG.Execute SQL
        End If
        
        CnG.Execute "delete from t_spl where employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "and date= '" & Format(TDBGrid1.Columns("date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "' " & _
                    "and company_code= '" & TDBCombo_company.Text & "' " & _
                    "and seq = " & TDBGrid1.Columns("seq").CellText(item) & ""
        
    Next
    CnG.CommitTrans
    Call load_data_spl
    MsgBox i & " overtime's data are successfully deleted", vbInformation, headerMSG
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
Dim v_verify As String
Dim v_level As String
    
    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsSPL
        Call load_data_ot
        
        SQL = "SELECT user_level FROM m_user WHERE user_name = '" & .Fields("user_approval").Value & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            v_level = rs!user_level
        Else
            v_level = 0
        End If
        rs.Close
        
        If v_level = 100 Then
            SQL = "SELECT full_name FROM m_user WHERE user_name = '" & .Fields("user_approval").Value & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                v_verify = rs!full_name
                lblVerify.Caption = "Verified by " & v_verify
            Else
                lblVerify.Caption = ""
            End If
        Else
            SQL = "SELECT employee_name FROM m_user a join m_employee b on a.employee_code = b.employee_code " & _
                    "WHERE user_name = '" & .Fields("user_approval").Value & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                v_verify = rs!EMPLOYEE_NAME
                lblVerify.Caption = "Verified by " & v_verify
            Else
                lblVerify.Caption = ""
            End If
        End If
        rs.Close
                
        txt_employee_code = .Fields("employee_code").Value
        txt_nik = .Fields("nik").Value
        txt_employee_name = .Fields("employee_name").Value
        DTPicker1 = .Fields("date").Value
        TDBTime1.Value = .Fields("start_time").Value
        TDBTime2.Value = .Fields("end_time").Value
        txt_description = .Fields("description").Value
        TDBCombo_ot.Text = .Fields("ot_code").Value
        txt_ot_name.Text = .Fields("ot_name").Value
        
        txt15.Text = .Fields("ot_15").Value
        txt2.Text = .Fields("ot_20").Value
        txt3.Text = .Fields("ot_30").Value
        txt4.Text = .Fields("ot_40").Value
        txt6.Text = IIf(IsNull(.Fields("ot_60").Value), 0, .Fields("ot_60").Value)
        
        chkBreak.Value = IIf(IsNull(.Fields("flag_break").Value), 0, .Fields("flag_break").Value)
        txtBreak.Text = IIf(IsNull(.Fields("int_break").Value), 0, .Fields("int_break").Value)
        
        chkManual.Value = IIf(IsNull(.Fields("flag_manual").Value), 0, .Fields("flag_manual").Value)
        txtMeal.Text = IIf(IsNull(.Fields("meal_allowance").Value), 0, .Fields("meal_allowance").Value)
        txtTransport.Text = IIf(IsNull(.Fields("transport_allowance").Value), 0, .Fields("transport_allowance").Value)
        
        chkOnCall.Value = IIf(IsNull(.Fields("flag_on_call").Value), 0, .Fields("flag_on_call").Value)
        chkShift2.Value = IIf(IsNull(.Fields("flag_shift2").Value), 0, .Fields("flag_shift2").Value)
        chkShift3.Value = IIf(IsNull(.Fields("flag_shift3").Value), 0, .Fields("flag_shift3").Value)
        
        vSeq = .Fields("seq").Value
        
    End With
End Sub

Private Sub cmdEdit_Click()
    int_mode = 2
    Call load_mode
    
    btnApprove.Visible = False
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdNew_Click()
    int_mode = 1
    vModeLoad = 0
    Call load_mode
    
    btnApprove.Visible = False
    lblVerify.Caption = ""
    lblVerify.Visible = False
    
    chkBreak.Value = 0
    Call load_data_ot
End Sub

Private Sub insert_new_data()
Dim v_total_ot As Double
Dim vGetMeal As Integer
Dim vDate As String
Dim totOT As Double
Dim vFlagLate As Integer
Dim vFlagAutoApprove As Integer

On Error GoTo Err
    'Mencari flag auto approve
    If rs.State Then rs.Close
    SQL = "SELECT flag_auto_approve FROM m_pref_OT"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vFlagAutoApprove = rs.Fields(0).Value
    Else
        vFlagAutoApprove = 0
    End If
    rs.Close
    
    'Mencari flag late
    SQL = "SELECT ifnull(flag_inc_late,0) flag_late FROM m_employee " & _
          "WHERE employee_code = '" & txt_employee_code.Text & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly
    
    If rscari.RecordCount > 0 Then
        vFlagLate = rscari!flag_late
    End If
    rscari.Close
    
    tglAwal = DTPicker1.Value
    If Combo3.Text = "TO" Then
        tglAkhir = DTPicker2.Value
    Else
        tglAkhir = DTPicker1.Value
    End If
    
    totOT = Val(txt15.Text) + Val(txt2.Text) + Val(txt3.Text) + Val(txt4.Text) + Val(txt6.Text)
    v_total_ot = (txt15.Text * 1.5) + (txt2.Text * 2) + (txt3.Text * 3) + (txt4.Text * 4) + (txt6.Text * 6)
    
    CnG.BeginTrans
    
    vStartTime = Format(TDBTime1.Value, "yyyy-mm-dd hh:mm:ss")
    If TDBTime2.Value < TDBTime1.Value Then
        vEndTime = Format(TDBTime2.Value + 1, "yyyy-mm-dd hh:mm:ss")
    Else
        vEndTime = Format(TDBTime2.Value, "yyyy-mm-dd hh:mm:ss")
    End If
    
    While tglAwal <= tglAkhir
        SQL = "SELECT MAX(seq) jmlSeq FROM t_spl WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND date = '" & Format(tglAwal, "yyyy-MM-dd HH:mm:ss") & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vSeq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
        Else
            vSeq = 1
        End If
        rs.Close
    
        If chkManual.Value = 0 Then
            vDate = Format(tglAwal, "yyyy-MM-dd HH:mm:ss")
        Else
            If TDBCombo_ot.Text = "HK" Then
                vDate = Format(tglAwal, "yyyy-MM-01 HH:mm:ss")
            ElseIf TDBCombo_ot.Text = "HL" Then
                vDate = Format(tglAwal, "yyyy-MM-02 HH:mm:ss")
            ElseIf TDBCombo_ot.Text = "HLL" Then
                vDate = Format(tglAwal, "yyyy-MM-03 HH:mm:ss")
            ElseIf TDBCombo_ot.Text = "HLSL" Then
                vDate = Format(tglAwal, "yyyy-MM-04 HH:mm:ss")
            End If
        End If
    
        If chkManual.Value = 0 Then
            SQL = "INSERT INTO t_spl(company_code,employee_code,description,date,start_time,end_time," & _
                    "flag_approval,ot_spl,ot_15,ot_20,ot_30,ot_40,ot_60,total_ot,ot_code,meal_allowance,transport_allowance," & _
                    "flag_break,int_break,flag_manual,flag_on_call,flag_shift2,flag_shift3,entry_date,entry_user,seq,flag_change_status) " & _
                  "VALUES( " & _
                    "'" & TDBCombo_company.Text & "','" & txt_employee_code.Text & "','" & txt_description.Text & "'," & _
                    "'" & vDate & "','" & vStartTime & "'," & _
                    "'" & vEndTime & "','" & IIf(vFlagAutoApprove <> 0, 1, 0) & "','" & jmlJam & "','" & txt15.Text & "'," & _
                    "'" & txt2.Text & "','" & txt3.Text & "','" & txt4.Text & "','" & txt6.Text & "'," & _
                    "'" & v_total_ot & "','" & TDBCombo_ot.Text & "','" & txtMeal.Text & "','" & txtTransport.Text & "'," & _
                    "'" & chkBreak.Value & "','" & Val(txtBreak.Text) & "','" & chkManual.Value & "','" & chkOnCall.Value & "'," & _
                    "'" & chkShift2.Value & "','" & chkShift3.Value & "',now(),'" & LOGIN_NAME & "'," & vSeq & ",'" & chkChangeStatus.Value & "')"
        Else
            SQL = "INSERT INTO t_spl(company_code,employee_code,description,date," & _
                    "flag_approval,ot_spl,ot_15,ot_20,ot_30,ot_40,ot_60,total_ot,ot_code,meal_allowance,transport_allowance," & _
                    "flag_break,int_break,flag_manual,flag_on_call,flag_shift2,flag_shift3,entry_date,entry_user, seq) " & _
                  "VALUES( " & _
                    "'" & TDBCombo_company.Text & "','" & txt_employee_code.Text & "','" & txt_description.Text & "'," & _
                    "'" & vDate & "','" & IIf(vFlagAutoApprove <> 0, 1, 0) & "','" & totOT & "','" & txt15.Text & "'," & _
                    "'" & txt2.Text & "','" & txt3.Text & "','" & txt4.Text & "','" & txt6.Text & "'," & _
                    "'" & v_total_ot & "','" & TDBCombo_ot.Text & "','" & txtMeal.Text & "','" & txtTransport.Text & "'," & _
                    "'" & chkBreak.Value & "','" & Val(txtBreak.Text) & "','" & chkManual.Value & "','" & chkOnCall.Value & "'," & _
                    "'" & chkShift2.Value & "','" & chkShift3.Value & "',now(),'" & LOGIN_NAME & "', " & vSeq & ",'" & chkChangeStatus.Value & "')"
        End If
        CnG.Execute SQL
        
        'Update status attendance
        Call cariShift(txt_employee_code.Text, Format(tglAwal, "yyyy-MM-dd"))
        
        If chkManual.Value = 0 Then
            If chkChangeStatus.Value <> 0 Then
                If rscari.State Then rscari.Close
                SQL = "select * from h_attendance " & _
                      "WHERE att_date = '" & Format(tglAwal, "yyyy-MM-dd HH:mm:ss") & "' " & _
                        "AND employee_code = '" & txt_employee_code.Text & "'"
                rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rscari.RecordCount > 0 Then
                    SQL = "UPDATE h_attendance SET status = 'OT', flag_manual = 1 " & _
                          "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                            "AND att_date = '" & Format(tglAwal, "yyyy-MM-dd HH:mm:ss") & "'"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO h_attendance (employee_code, att_date, enrollnumber, shift_number, group_code, shift_code, " & _
                            "`status`,time_in, time_out,flag_io, flag_present, flag_late, flag_early, description, entry_date, t_flag_day_over, flag_inc_late) " & _
                          "SELECT '" & txt_employee_code.Text & "', '" & Format(tglAwal, "yyyy-MM-dd") & "'," & Val(txt_employee_code.Text) & ", 1, '" & vGroupCode & "', '" & vShiftCode & "'," & _
                            "'OT','" & vStartTime & "','" & vEndTime & "',1, 1, 0, 0,'" & txt_description.Text & "',Now(), 1, " & vFlagLate & ""
                    CnG.Execute SQL
                End If
                rscari.Close
            Else
                
            End If
        End If
        
        tglAwal = tglAwal + 1
    Wend
    
    CnG.CommitTrans
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    SQL = "DELETE FROM t_spl WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND company_code = '" & TDBCombo_company.Text & "' " & _
            "AND date = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
            "AND seq = " & vSeq & ""
    CnG.Execute SQL
    
    Call insert_new_data
    
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
    
    Call load_data_spl
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
    
    cmdExit.Enabled = f
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            If Not LCase(Ctr.name) = "dtpicker_from" _
                And Not LCase(Ctr.name) = "dtpicker_to" Then Ctr.Value = Now
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
    txt_employee_code = ""
    txt_nik = ""
    txt_employee_name = ""
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    DTPicker2.Visible = False
    Combo3.Text = "..."
            
    txtBreak.Text = 0
    TDBTime1.Value = Now
    TDBTime2.Value = Now
    txt_description = ""
    chkOnCall.Value = 0
    
    txt15.Text = 0
    txt2.Text = 0
    txt3.Text = 0
    txt4.Text = 0
    txt6.Text = 0
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        If vModeLoad = 0 Then
            If TDBCombo_company.Text = "" Then
                MsgBox "Company Is Not Selected!", vbExclamation, headerMSG
                TDBCombo_company.SetFocus
                
                int_mode = 0
                Call load_mode
                Exit Sub
            End If
            
            Call clear_view_data
            fra_entry.Visible = True
            txt_nik.Enabled = True
            
            TDBGrid1.Enabled = False
            Call set_new_data
            
            If txt_nik.Enabled = True Then
                txt_nik.SetFocus
            End If
        End If
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry.Visible = False
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        txt_employee_code.Enabled = False
        
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

Private Sub cmdSearch_Click()
    Call load_data_spl
End Sub

Private Sub Combo3_Click()
    If Combo3.Text = "TO" Then
        DTPicker2.Value = DTPicker1.Value
        DTPicker2.Visible = True
    Else
        DTPicker2.Visible = False
    End If
End Sub

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_from, DTPicker_to)
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    Call cari_jam
End Sub

Private Sub Form_Load()
    Call createGridKar
    Call load_data_company
    
    Combo3.AddItem "..."
    Combo3.AddItem "TO"
    
    oClause = ""
    btnApprove.Visible = False
    optApproval(0).Value = True
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
    DTPicker_Periode.Value = Now
    chkManual.Value = 0
        
    Call load_data_user_access(Me)
    timer1.Enabled = True
    
    If vModeLoad = 0 Then
        int_mode = 0
    Else
        int_mode = int_mode
    End If
    
    Call load_mode
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsSPL.Filter = adFilterNone
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

Private Sub optApproval_Click(Index As Integer)
    Call load_data_spl
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsSPL.Filter = getFilter()
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

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_spl
        
        optApproval(0).Value = True
    End If
End Sub

Private Sub TDBCombo_ot_ItemChange()
    If TDBCombo_ot.ApproxCount > 0 Then
        TDBCombo_ot.Text = TDBCombo_ot.Columns("ot_code").Value
        txt_ot_name = TDBCombo_ot.Columns("ot_name").Value
    End If
    
    txt15.Text = 0
    txt2.Text = 0
    txt3.Text = 0
    txt4.Text = 0
    txt6.Text = 0
    txtMeal.Text = 0
    txtTransport.Text = 0
    
    Call cari_jam
    Call hitungLembur
End Sub

Private Sub load_data_spl()
    If rsSPL.State Then rsSPL.Close
    SQL = "select a.*,c.ot_name,b.nik,b.employee_name, date as date_, date, " & _
            "start_time as start_time_, " & _
            "end_time as end_time_, " & _
            "CASE WHEN ifnull(a.flag_manual,0) = 0 THEN 'AUTO' ELSE 'MANUAL' END type " & _
            "from t_spl a join m_employee b on a.employee_code = b.employee_code " & _
            "join m_ot c on a.ot_code = c.ot_code " & _
          "where b.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
            "and (date(date) between '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & _
            "and ifnull(flag_approval,0) = " & IIf(optApproval(0).Value, 0, 1) & " " & oClause
    rsSPL.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBGrid1.DataSource = rsSPL
End Sub

Public Sub load_data_ot()
    If rsot.State Then rsot.Close
    SQL = "select * from m_ot order by ot_name"
    rsot.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_ot.RowSource = rsot
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid1.Columns(ColIndex).Caption = "DATE FROM" _
    Or TDBGrid1.Columns(ColIndex).Caption = "DATE TO" Then
        Value = Format(Value, "yyyy-mm-dd")
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rs2 As New ADODB.Recordset
Dim vFlagApproval As Integer
Dim vUserApproval As String
Dim a As String
    
    a = IIf(IsNull(TDBGrid1.Columns("flag_approval").Value), 0, TDBGrid1.Columns("flag_approval").Value)
    vFlagApproval = IIf(a = "", 0, a)
    
    SQL = "SELECT a.employee_code FROM m_form_approval_dtl a join m_user b on a.employee_code = b.employee_code " & _
            "WHERE a.form_name = 'frm_trans_spl' and b.user_name = '" & LOGIN_NAME & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Or LOGIN_LEVEL = 100 Then
        If vFlagApproval = 0 And fra_entry.Visible = False Then
            btnApprove.Visible = True
            lblVerify.Visible = False
        ElseIf vFlagApproval = 1 And fra_entry.Visible = True Then
            btnApprove.Visible = False
            lblVerify.Visible = True
            
            SQL = "SELECT a.*,b.employee_name FROM m_user a LEFT JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "WHERE a.user_name = '" & TDBGrid1.Columns("user_approval").Value & "'"
            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs2.RecordCount > 0 Then
                If rs2!user_level = 100 Then
                    vUserApproval = "ADMINISTRATOR"
                Else
                    vUserApproval = rs2!EMPLOYEE_NAME
                End If
            Else
                vUserApproval = "ADMINISTRATOR"
            End If
            rs2.Close
            
            lblVerify.Caption = "Disetujui Oleh " & vUserApproval
        End If
    Else
        btnApprove.Visible = False
    End If
    rs.Close
End Sub

Private Sub TDBTime1_Change()
    Call hitungLembur
End Sub

Private Sub TDBTime2_Change()
    Call hitungLembur
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub hitungLembur()
Dim rscari As New ADODB.Recordset
Dim vTime1, vTime2 As String
Dim vFlagTransportEmp As Integer
    
    vTime1 = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(TDBTime1.Value, "hh:mm:ss")
    vTime2 = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(TDBTime2.Value, "hh:mm:ss")
    
    If Format(vTime2, "hh:mm") < Format(vTime1, "hh:mm") Then
        jmlJam = IIf(IsNull(DateDiff("n", vTime1, DateAdd("d", 1, vTime2))), 0, DateDiff("n", vTime1, DateAdd("d", 1, vTime2)))
    Else
        jmlJam = IIf(IsNull(DateDiff("n", vTime1, vTime2)), 0, DateDiff("n", vTime1, vTime2))
    End If
    
    If chkBreak.Value = 0 Then
        jmlJam = Round(jmlJam / 60, 2)
    Else
        jmlJam = Round(jmlJam / 60, 2) - Val(txtBreak.Text)
    End If
    
    SQL = "SELECT * FROM m_ot_detail WHERE ot_code = '" & TDBCombo_ot.Text & "' AND from_value < " & jmlJam & ""
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    If rscari.RecordCount > 0 Then
        While Not rscari.EOF
            If jmlJam > rscari!to_value Then
                If rscari!pengali = 1.5 Then
                    txt15.Text = rscari!to_value
                    txt2.Text = "0"
                    txt3.Text = "0"
                    txt4.Text = "0"
                    txt6.Text = "0"
                ElseIf rscari!pengali = 2 Then
                    If rscari!ot_number = 1 Then
                        txt2.Text = rscari!to_value
                    Else
                        txt2.Text = jmlJam - rscari!from_value
                    End If
                    txt3.Text = "0"
                    txt4.Text = "0"
                    txt6.Text = "0"
                ElseIf rscari!pengali = 3 Then
                    txt3.Text = "1"
                    txt4.Text = "0"
                    txt6.Text = "0"
                ElseIf rscari!pengali = 4 Then
                    If rscari!ot_number = 1 Then
                        txt4.Text = rscari!to_value
                    Else
                        txt4.Text = jmlJam - rscari!from_value
                    End If
                    txt6.Text = "0"
                ElseIf rscari!pengali = 6 Then
                    If rscari!ot_number = 1 Then
                        txt6.Text = rscari!to_value
                    Else
                        txt6.Text = jmlJam - rscari!from_value
                    End If
                End If
            Else
                If rscari!pengali = 1.5 Then
                    txt15.Text = jmlJam
                    txt2.Text = "0"
                    txt3.Text = "0"
                    txt4.Text = "0"
                    txt6.Text = "0"
                ElseIf rscari!pengali = 2 Then
                    txt2.Text = jmlJam - rscari!from_value
                    If Val(txt2.Text) < 0 Then txt2.Text = 0
                    txt3.Text = "0"
                    txt4.Text = "0"
                    txt6.Text = "0"
                ElseIf rscari!pengali = 3 Then
                    txt3.Text = jmlJam - txt2.Text - txt15.Text
                    If Val(txt3.Text) < 0 Then txt3.Text = 0
                    txt4.Text = "0"
                    txt6.Text = "0"
                ElseIf rscari!pengali = 4 Then
                    txt4.Text = jmlJam - txt3.Text - txt2.Text - txt15.Text
                    txt4.Text = Round(txt4.Text, 2)
                    If Val(txt4.Text) < 0 Then txt4.Text = 0
                    txt6.Text = "0"
                ElseIf rscari!pengali = 6 Then
                    txt6.Text = jmlJam - txt4.Text - txt3.Text - txt2.Text - txt15.Text
                    If Val(txt6.Text) < 0 Then txt6.Text = 0
                End If
            End If
        rscari.MoveNext
        Wend
    End If
    rscari.Close
    
    If chkManual.Value = 0 Then
        SQL = "SELECT ot_max_meal, ot_max_trans, flag_meal, flag_trans FROM m_pref_ot"
        rs.Open SQL, CnG, adOpenForwardOnly
        
        If rs.RecordCount > 0 Then
            vOTMax_Meal = rs!ot_max_meal
            vFlagMeal = rs!flag_meal
            vOTMax_Trans = rs!ot_max_trans
            vFlagTrans = rs!flag_trans
        Else
            vOTMax_Meal = 0
            vFlagMeal = 0
            vOTMax_Trans = 0
            vFlagTrans = 0
        End If
        rs.Close
        
        SQL = "SELECT flag_transport FROM m_employee WHERE employee_code = '" & txt_employee_code.Text & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vFlagTransportEmp = rs!flag_transport
        End If
        rs.Close
        
        If vFlagMeal <> 0 Then
            If jmlJam >= vOTMax_Meal Then
                txtMeal.Text = 1
            Else
                txtMeal.Text = 0
            End If
        Else
            txtMeal.Text = 0
        End If
        
        If vFlagTransportEmp <> 0 Then
            If vFlagTrans <> 0 Then
                If TDBCombo_ot.Text <> "HK" Then
                    If jmlJam > 0 Then
                        txtTransport.Text = 1
                    Else
                        txtTransport.Text = 0
                    End If
                Else
                    txtTransport.Text = 0
                End If
            Else
                txtTransport.Text = 0
            End If
        Else
            txtTransport = 0
        End If
    End If
    
    If rs.State Then rs.Close
    SQL = "SELECT change_status_hours FROM m_pref_ot"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        vChangeStatusHours = rs.Fields(0).Value
    Else
        vChangeStatusHours = 0
    End If
    rs.Close
    
    If TDBCombo_ot.Text <> "HK" Or jmlJam >= vChangeStatusHours Then
        chkChangeStatus.Value = 1
    Else
        chkChangeStatus.Value = 0
    End If
End Sub

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
Dim rsemp As New ADODB.Recordset

    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If
        
        rsemp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rsemp.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rsemp.MoveFirst
            While Not rsemp.EOF
                LynxGrid2.AddItem rsemp!nik & vbTab & rsemp!EMPLOYEE_NAME & vbTab & rsemp!employee_code
                rsemp.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rsemp.RecordCount = 1 Then
                rsemp.MoveFirst
                txt_employee_code.Text = rsemp!employee_code
                txt_employee_name.Text = rsemp!EMPLOYEE_NAME
                txt_nik.Text = rsemp!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
'        rs.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
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
        txt_employee_code.Text = ""
        txt_employee_name.Text = ""
    End If
    
    Call cari_jam
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub

Public Sub cari_jam()
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rsJam As New ADODB.Recordset

Dim vStatus As String
Dim v_shift As String
Dim vFlagHoliday As Integer
Dim vFlagShiftable As Integer
    
    SQL = "SELECT IFNULL(flag_shiftable,0) flag_shiftable FROM m_employee " & _
          "WHERE employee_code = '" & txt_employee_code.Text & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vFlagShiftable = rscari!flag_shiftable
    End If
    rscari.Close
    
    If vFlagShiftable = 0 Then
        SQL = "SELECT a.att_date,a.status,b.description FROM h_attendance a JOIN t_holiday b ON DATE(a.att_date) = DATE(b.holiday_date) " & _
              "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "' " & _
                "AND (WEEKDAY('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') = 5 OR WEEKDAY('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "') = 6) "
    Else
        SQL = "SELECT a.att_date,a.status,b.description FROM h_attendance a JOIN t_holiday b ON DATE(a.att_date) = DATE(b.holiday_date) " & _
              "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "' " & _
                "AND status = 'OF'"
    End If
    rsJam.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsJam.RecordCount > 0 Then
    
        SQL = "SELECT time_in, time_out FROM h_attendance " & _
              "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs2.RecordCount > 0 Then
            TDBTime1.Text = Format(rs2!time_in, "hh:mm")
            TDBTime2.Text = Format(rs2!time_out, "hh:mm")
        Else
            TDBTime1.Text = Format(Now, "hh:mm")
            TDBTime2.Text = Format(Now, "hh:mm")
            
            txt15.Text = 0
            txt2.Text = 0
            txt3.Text = 0
            txt4.Text = 0
            txt6.Text = 0
        End If
        rs2.Close
    Else
        SQL = "SELECT time_in, end_time, time_out, status FROM h_attendance " & _
              "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            
            If Format(DTPicker1.Value, "dddd") = "Sunday" Or Format(DTPicker1.Value, "dddd") = "Saturday" Then
                TDBTime1.Text = Format(rs2!time_in, "hh:mm")
                TDBTime2.Text = Format(rs2!time_out, "hh:mm")
            Else
                TDBTime1.Text = Format(rs2!end_time, "hh:mm")
                TDBTime2.Text = Format(rs2!time_out, "hh:mm")
            End If
        Else
            TDBTime1.Text = Format(Now, "hh:mm")
            TDBTime2.Text = Format(Now, "hh:mm")
            
            txt15.Text = 0
            txt2.Text = 0
            txt3.Text = 0
            txt4.Text = 0
            txt6.Text = 0
        End If
        rs2.Close
    End If
    
'    SQL = "SELECT a.holiday_date, a.ot_code, b.ot_name FROM t_holiday a join m_ot b on a.ot_code = b.ot_code " & _
'          "WHERE Date(holiday_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
'    rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rs2.RecordCount > 0 Then
'        TDBCombo_ot.Text = rs2!ot_code
'        txt_ot_name.Text = rs2!ot_name
'    Else
'        TDBCombo_ot.Text = ""
'        txt_ot_name.Text = ""
'    End If
'    rs2.Close
        
    rsJam.Close
End Sub

Private Sub txt_nik_Validate(Cancel As Boolean)
    Call cari_jam
End Sub


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid1.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid1.Columns(ColIndex).DataField
    Call load_data_spl

End Sub

Private Sub txtBreak_Change()
    Call hitungLembur
End Sub

Private Sub cariShift(vEmployee_Code As String, vTgl As String)
Dim vTeam As String
Dim vFlagShift As Integer
Dim vGroupRoll As String

    SQL = "SELECT ifnull(flag_shiftable,0) flag_shift FROM m_employee " & _
          "WHERE employee_code = '" & vEmployee_Code & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vFlagShift = rscari!flag_shift
    End If
    rscari.Close

    If vFlagShift = 0 Then
        SQL = "SELECT a.group_code,b.shift_code,a.shift_number " & _
              "FROM td_shift a JOIN tm_shift b ON a.shift_number = b.shift_number AND a.group_code = b.group_code " & _
              "WHERE DATE(b.start_date) <= DATE('" & vTgl & "') AND a.employee_code = '" & txt_employee_code.Text & "' " & _
              "ORDER BY b.start_date DESC LIMIT 1;"
        rscari.Open SQL, CnG, adOpenForwardOnly
        
        If rscari.RecordCount > 0 Then
            vGroupCode = rscari!group_code
            vShiftCode = rscari!shift_code
        End If
        rscari.Close
    Else
        SQL = "SELECT a.group_code " & _
              "FROM td_emp_group a JOIN tm_emp_group b ON a.group_number = b.emp_group_number AND a.emp_group_code = b.emp_group_code " & _
              "WHERE DATE(a.start_date) <= DATE('" & vTgl & "') AND a.employee_code = '" & vEmployee_Code & "' " & _
              "ORDER BY a.start_date DESC LIMIT 1;"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            vGroupCode = rscari!group_code
        End If
        rscari.Close
        
        SQL = "SELECT emp_group_code FROM td_emp_group WHERE employee_code = '" & vEmployee_Code & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            vGroupRoll = rscari!emp_group_code
        End If
        rscari.Close
        
        SQL = "SELECT shift_code FROM m_shift_new WHERE group_code = '" & vGroupRoll & "' AND DATE(shift_date) = DATE('" & vTgl & "')"
        rscari.Open SQL, CnG, adOpenForwardOnly
        
        If rscari.RecordCount > 0 Then
            vShiftCode = rscari!shift_code
        End If
        rscari.Close
        
        If vShiftCode = "OFF" Then
            vShiftCode = "S01"
        Else
            vShiftCode = vShiftCode
        End If
    End If
End Sub
