VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_spl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S P L"
   ClientHeight    =   9090
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
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin prj_panji.vbButton vbButton1 
      Height          =   375
      Left            =   4770
      TabIndex        =   41
      Top             =   1500
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Cari"
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
      MICON           =   "frm_trans_spl.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fra_status_emp 
      Height          =   585
      Left            =   11010
      TabIndex        =   34
      Top             =   1320
      Width           =   3405
      Begin VB.OptionButton optApproval 
         Caption         =   "NOT APPROVE"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1365
      End
      Begin VB.OptionButton optApproval 
         Caption         =   "APPROVE"
         Height          =   225
         Index           =   1
         Left            =   1890
         TabIndex        =   35
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.TextBox txt_division_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3420
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   1140
      Width           =   3855
   End
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   3465
      Left            =   2160
      TabIndex        =   27
      Top             =   4830
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   6112
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3420
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   750
      Width           =   3855
   End
   Begin VB.Frame fra_entry 
      Height          =   3405
      Left            =   240
      TabIndex        =   8
      Top             =   4050
      Width           =   14175
      Begin VB.OptionButton opt_fixed 
         Caption         =   "Fixed"
         Height          =   195
         Left            =   10530
         TabIndex        =   67
         Top             =   1290
         Width           =   1365
      End
      Begin VB.Frame fra_fixed 
         Height          =   1065
         Left            =   8400
         TabIndex        =   61
         Top             =   1470
         Width           =   5025
         Begin VB.TextBox txt_pengali2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   2130
            TabIndex        =   66
            Top             =   630
            Width           =   525
         End
         Begin VB.TextBox txt_value2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   360
            TabIndex        =   65
            Top             =   630
            Width           =   1485
         End
         Begin VB.TextBox txt_value1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   360
            TabIndex        =   62
            Top             =   180
            Width           =   1485
         End
         Begin VB.TextBox txt_pengali1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   2130
            TabIndex        =   64
            Top             =   180
            Width           =   525
         End
         Begin VB.TextBox txt_ot_value_fixed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   345
            Left            =   3240
            TabIndex        =   68
            Top             =   420
            Width           =   1515
         End
         Begin VB.Line Line3 
            X1              =   2700
            X2              =   3120
            Y1              =   810
            Y2              =   570
         End
         Begin VB.Line Line2 
            X1              =   2700
            X2              =   3120
            Y1              =   360
            Y2              =   570
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   225
            Index           =   1
            Left            =   1890
            TabIndex        =   69
            Top             =   690
            Width           =   285
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   225
            Index           =   0
            Left            =   1890
            TabIndex        =   63
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "+"
            Height          =   225
            Index           =   2
            Left            =   2760
            TabIndex        =   70
            Top             =   480
            Width           =   285
         End
      End
      Begin VB.Frame fra_manual 
         Height          =   1065
         Left            =   8400
         TabIndex        =   51
         Top             =   1470
         Width           =   5025
         Begin VB.TextBox txt_ot_value 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   345
            Left            =   3300
            TabIndex        =   57
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txt_pengali 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   2130
            TabIndex        =   55
            Top             =   420
            Width           =   525
         End
         Begin VB.TextBox txt_pembagi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   780
            TabIndex        =   53
            Top             =   630
            Width           =   675
         End
         Begin VB.TextBox txt_basic_salary 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   360
            TabIndex        =   52
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "JAM = "
            Height          =   225
            Left            =   2700
            TabIndex        =   56
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   225
            Left            =   1890
            TabIndex        =   54
            Top             =   480
            Width           =   285
         End
         Begin VB.Line Line1 
            X1              =   330
            X2              =   1890
            Y1              =   570
            Y2              =   570
         End
      End
      Begin VB.OptionButton opt_manual 
         Caption         =   "Manual"
         Height          =   195
         Left            =   9540
         TabIndex        =   60
         Top             =   1290
         Width           =   1365
      End
      Begin VB.OptionButton opt_auto 
         Caption         =   "Otomatis"
         Height          =   195
         Left            =   8430
         TabIndex        =   59
         Top             =   1290
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.Frame fra_auto 
         Height          =   675
         Left            =   8400
         TabIndex        =   42
         Top             =   1470
         Width           =   5025
         Begin VB.TextBox txt_15 
            Height          =   315
            Left            =   510
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   210
            Width           =   615
         End
         Begin VB.TextBox txt_20 
            Height          =   315
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   210
            Width           =   615
         End
         Begin VB.TextBox txt_30 
            Height          =   315
            Left            =   3030
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   210
            Width           =   615
         End
         Begin VB.TextBox txt_40 
            Height          =   315
            Left            =   4290
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "1.5 X"
            Height          =   195
            Left            =   90
            TabIndex        =   50
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "2 X"
            Height          =   195
            Left            =   1440
            TabIndex        =   49
            Top             =   270
            Width           =   225
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "3 X"
            Height          =   195
            Left            =   2700
            TabIndex        =   48
            Top             =   270
            Width           =   225
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "4 X"
            Height          =   195
            Left            =   3960
            TabIndex        =   47
            Top             =   270
            Width           =   225
         End
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
         Left            =   8400
         TabIndex        =   3
         Top             =   840
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "frm_trans_spl.frx":05A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_trans_spl.frx":060B
         Spin            =   "frm_trans_spl.frx":065B
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
      Begin VB.CheckBox chk_holiday 
         Caption         =   "Hari Libur"
         Height          =   285
         Left            =   9990
         TabIndex        =   2
         Top             =   480
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   8400
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   91488259
         CurrentDate     =   39270
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   8400
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2580
         Width           =   5475
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   840
         Width           =   3495
      End
      Begin TDBTime6Ctl.TDBTime TDBTime2 
         Height          =   315
         Left            =   9840
         TabIndex        =   4
         Top             =   840
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "frm_trans_spl.frx":0683
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_trans_spl.frx":06E8
         Spin            =   "frm_trans_spl.frx":0738
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
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin prj_panji.vbButton cmdBrowse 
         Height          =   315
         Left            =   3300
         TabIndex        =   28
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
         MICON           =   "frm_trans_spl.frx":0760
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TIPE OT"
         Height          =   195
         Left            =   7200
         TabIndex        =   58
         Top             =   1290
         Width           =   585
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
         Height          =   1065
         Left            =   660
         TabIndex        =   30
         Top             =   1560
         Width           =   5685
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   195
         Left            =   9510
         TabIndex        =   18
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MULAI | SELESAI"
         Height          =   195
         Left            =   6555
         TabIndex        =   17
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL"
         Height          =   195
         Left            =   7080
         TabIndex        =   16
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KETERANGAN"
         Height          =   195
         Left            =   6780
         TabIndex        =   13
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KODE KARY."
         Height          =   195
         Left            =   675
         TabIndex        =   12
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NAMA KARY."
         Height          =   195
         Left            =   615
         TabIndex        =   11
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   7560
      Width           =   14175
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   705
         Left            =   660
         TabIndex        =   20
         Top             =   360
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
         MICON           =   "frm_trans_spl.frx":077C
         PICN            =   "frm_trans_spl.frx":0798
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdSave 
         Height          =   705
         Left            =   1680
         TabIndex        =   21
         Top             =   360
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
         MICON           =   "frm_trans_spl.frx":182A
         PICN            =   "frm_trans_spl.frx":1846
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdEdit 
         Height          =   705
         Left            =   2700
         TabIndex        =   22
         Top             =   360
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
         MICON           =   "frm_trans_spl.frx":28D8
         PICN            =   "frm_trans_spl.frx":28F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdDelete 
         Height          =   705
         Left            =   3720
         TabIndex        =   23
         Top             =   360
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
         MICON           =   "frm_trans_spl.frx":3986
         PICN            =   "frm_trans_spl.frx":39A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdCancel 
         Height          =   705
         Left            =   4740
         TabIndex        =   24
         Top             =   360
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
         MICON           =   "frm_trans_spl.frx":4A34
         PICN            =   "frm_trans_spl.frx":4A50
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   9780
         TabIndex        =   25
         Top             =   390
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
         MICON           =   "frm_trans_spl.frx":5AE2
         PICN            =   "frm_trans_spl.frx":5AFE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton btnApprove 
         Height          =   705
         Left            =   7890
         TabIndex        =   29
         Top             =   390
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Setujui"
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
         MICON           =   "frm_trans_spl.frx":6B90
         PICN            =   "frm_trans_spl.frx":6BAC
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
      Height          =   5505
      Left            =   240
      TabIndex        =   6
      Top             =   1950
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9710
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
      Columns(1).Caption=   "KODE KARY."
      Columns(1).DataField=   "nik"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NAMA KARY."
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
      Columns(5).Caption=   "TANGGAL"
      Columns(5).DataField=   "date_"
      Columns(5).NumberFormat=   "yyyy-MM-dd"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "MULAI"
      Columns(6).DataField=   "start_time_"
      Columns(6).NumberFormat=   "HH:nn"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "SELESAI"
      Columns(7).DataField=   "end_time_"
      Columns(7).NumberFormat=   "HH:nn"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   4
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "DISETUJUI"
      Columns(8).DataField=   "flag_approval"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "KETERANGAN"
      Columns(9).DataField=   "description"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "USER APPROVAL"
      Columns(10).DataField=   "user_approval"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   1
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3307"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3228"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=5292"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5212"
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
      Splits(0)._ColumnProps(33)=   "Column(5).Width=3466"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=3387"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(48)=   "Column(8).Width=1508"
      Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=1429"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=513"
      Splits(0)._ColumnProps(52)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(54)=   "Column(9).Width=5371"
      Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=5292"
      Splits(0)._ColumnProps(57)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(59)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(63)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(64)=   "Column(10).Order=11"
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
      Caption         =   "DAFTAR SPL"
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
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=99"
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
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=50,.parent=99,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=100"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=101"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=103"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=58,.parent=99,.alignment=2"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=100"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=101"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=103"
      _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=20,.parent=99,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=17,.parent=100"
      _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=18,.parent=101"
      _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=19,.parent=103"
      _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=62,.parent=99"
      _StyleDefs(72)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=100"
      _StyleDefs(73)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=101"
      _StyleDefs(74)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=103"
      _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=24,.parent=99"
      _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=21,.parent=100"
      _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=22,.parent=101"
      _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=23,.parent=103"
      _StyleDefs(79)  =   "Named:id=33:Normal"
      _StyleDefs(80)  =   ":id=33,.parent=0"
      _StyleDefs(81)  =   "Named:id=34:Heading"
      _StyleDefs(82)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   ":id=34,.wraptext=-1"
      _StyleDefs(84)  =   "Named:id=35:Footing"
      _StyleDefs(85)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   "Named:id=36:Selected"
      _StyleDefs(87)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(88)  =   "Named:id=37:Caption"
      _StyleDefs(89)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(90)  =   "Named:id=38:HighlightRow"
      _StyleDefs(91)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(92)  =   "Named:id=39:EvenRow"
      _StyleDefs(93)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(94)  =   "Named:id=40:OddRow"
      _StyleDefs(95)  =   ":id=40,.parent=33"
      _StyleDefs(96)  =   "Named:id=41:RecordSelector"
      _StyleDefs(97)  =   ":id=41,.parent=34"
      _StyleDefs(98)  =   "Named:id=42:FilterBar"
      _StyleDefs(99)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1620
      OleObjectBlob   =   "frm_trans_spl.frx":7C3E
      TabIndex        =   7
      Top             =   750
      Width           =   1695
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_division 
      Height          =   375
      Left            =   1620
      OleObjectBlob   =   "frm_trans_spl.frx":9BFC
      TabIndex        =   32
      Top             =   1140
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker_from 
      Height          =   315
      Left            =   1620
      TabIndex        =   37
      Top             =   1530
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   91488259
      CurrentDate     =   41332
   End
   Begin MSComCtl2.DTPicker DTPicker_to 
      Height          =   315
      Left            =   3360
      TabIndex        =   38
      Top             =   1530
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   91488259
      CurrentDate     =   41332
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "s/d"
      Height          =   195
      Left            =   3060
      TabIndex        =   40
      Top             =   1590
      Width           =   225
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TANGGAL"
      Height          =   195
      Left            =   225
      TabIndex        =   39
      Top             =   1590
      Width           =   690
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIVISI"
      Height          =   195
      Left            =   240
      TabIndex        =   33
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "S P L"
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
      TabIndex        =   26
      Top             =   150
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   810
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_spl.frx":BBBB
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
Dim rsDivision As New ADODB.Recordset
Dim rsSPL As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim jmlJam As Double
Dim vHoliday As Integer
Dim vTimeIn As String
Dim vTimeOut As String

Private Function check_validate_exist_new() As Boolean
    check_validate_exist_new = False
    
    SQL = "select count(employee_code) as rec_count from t_spl where employee_code = '" & txt_employee_code.Text & "' " & _
                "and left(date,10)= '" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' " & _
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
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
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
        MsgBox "Kode Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_nik.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi employee name
    If Trim(txt_employee_name) = "" Then
        MsgBox "Nama Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_employee_name.SetFocus
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
    i = MsgBox("Apakah anda yakin menyetujui " _
        & SelBks.Count & " data SPL ini ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        SQL = "UPDATE t_spl SET flag_approval = 1," & _
                "user_approval = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "and left(date,10)= '" & Format(TDBGrid1.Columns("date_").CellText(item), "yyyy-MM-dd") & "' " & _
                    "and company_code= '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
                
    Next
    CnG.CommitTrans
    Call load_data_spl
    MsgBox i & " data SPL berhasil disetujui...", vbInformation, headerMSG
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub chk_holiday_Click()
    Call hitungLembur
End Sub

Private Sub cmd_browse_Click()
    frm_lookup_mst_employee.public_int_mode = 161
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_refresh_Click()
    Call load_data_spl
End Sub

Private Sub cmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim item
    
On Error GoTo Err
    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid1.SelBookmarks
    i = MsgBox("Apakah anda yakin menghapus " _
        & SelBks.Count & " sata SPL ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        CnG.Execute "delete from t_spl where employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                    "and left(date,10)= '" & Format(TDBGrid1.Columns("date_").CellText(item), "yyyy-MM-dd") & "' " & _
                    "and company_code= '" & TDBCombo_company.Text & "'"
                
    Next
    CnG.CommitTrans
    Call load_data_spl
    MsgBox i & " data SPL berhasil dihapus...", vbInformation, headerMSG
    
'    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
'    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
'    CnG.Execute SQL
'    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
Dim v_verify As String
Dim v_level As String
Dim vFlagType As Integer

    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsSPL
        
        txt_employee_code = .Fields("employee_code").Value
        txt_nik = .Fields("nik").Value
        txt_employee_name = .Fields("employee_name").Value
        DTPicker1 = .Fields("date").Value
        TDBTime1.Value = .Fields("start_time").Value
        TDBTime2.Value = .Fields("end_time").Value
        txt_description = .Fields("description").Value
        
        vFlagType = .Fields("flag_type").Value
        
        opt_auto.Value = IIf(vFlagType = 0, True, False)
        opt_manual.Value = IIf(vFlagType = 1, True, False)
        opt_fixed.Value = IIf(vFlagType = 2, True, False)
        
        If vFlagType = 0 Then
            txt_15.Text = .Fields("ot_15").Value
            txt_20.Text = .Fields("ot_20").Value
            txt_30.Text = .Fields("ot_30").Value
            txt_40.Text = .Fields("ot_40").Value
            
            txt_basic_salary.Text = 0
            txt_pengali.Text = 0
            txt_pembagi.Text = 0
            txt_ot_value.Text = 0
            
            txt_value1.Text = 0
            txt_pengali1.Text = 0
            txt_value2.Text = 0
            txt_pengali2.Text = 0
            txt_ot_value_fixed.Text = 0
        ElseIf vFlagType = 1 Then
            txt_basic_salary.Text = FormatNumber(.Fields("ot_basic_salary").Value)
            txt_pengali.Text = .Fields("ot_pengali").Value
            txt_pembagi.Text = .Fields("ot_pembagi").Value
            txt_ot_value.Text = FormatNumber(.Fields("ot_hasil").Value)
            
            txt_15.Text = 0
            txt_20.Text = 0
            txt_30.Text = 0
            txt_40.Text = 0
            
            txt_value1.Text = 0
            txt_pengali1.Text = 0
            txt_value2.Text = 0
            txt_pengali2.Text = 0
            txt_ot_value_fixed.Text = 0
        Else
            txt_basic_salary.Text = 0
            txt_pengali.Text = 0
            txt_pembagi.Text = 0
            txt_ot_value.Text = 0
            
            txt_15.Text = 0
            txt_20.Text = 0
            txt_30.Text = 0
            txt_40.Text = 0
            
            txt_value1.Text = FormatNumber(.Fields("ot_value1").Value)
            txt_pengali1.Text = FormatNumber(.Fields("ot_pengali1").Value)
            txt_value2.Text = FormatNumber(.Fields("ot_value2").Value)
            txt_pengali2.Text = FormatNumber(.Fields("ot_pengali2").Value)
            txt_ot_value_fixed.Text = FormatNumber(.Fields("ot_hasil_fixed").Value)
        End If
        
        chk_holiday.Value = IIf(IsNull(.Fields("flag_holiday").Value), 0, .Fields("flag_holiday").Value)
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

Private Sub cmdNew_Click()
    int_mode = 1
    Call load_mode
    
    btnApprove.Visible = False
    lblVerify.Caption = ""
End Sub

Private Sub insert_new_data()
Dim v_total_ot As Double
Dim vEndTime As String

On Error GoTo Err
    v_total_ot = (txt_15.Text * 1.5) + (txt_20.Text * 2) + (txt_30.Text * 3) + (txt_40.Text * 4)
    
    CnG.BeginTrans
    
    If TDBTime2.Value < TDBTime1.Value Then
        vEndTime = Format(TDBTime2.Value + 1, "yyyy-mm-dd hh:mm:ss")
    Else
        vEndTime = Format(TDBTime2.Value, "yyyy-mm-dd hh:mm:ss")
    End If
    
    If opt_auto.Value Then
        SQL = "INSERT INTO t_spl(company_code,employee_code,description,date,start_time,end_time," & _
                "flag_approval,ot_spl,ot_15,ot_20,ot_30,ot_40,total_ot,flag_holiday,flag_type,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & TDBCombo_company.Text & "','" & txt_employee_code.Text & "','" & txt_description.Text & "'," & _
                "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(TDBTime1.Value, "yyyy-mm-dd hh:mm:ss") & "'," & _
                "'" & vEndTime & "',0,'" & jmlJam & "','" & txt_15.Text & "'," & _
                "'" & txt_20.Text & "','" & txt_30.Text & "','" & txt_40.Text & "'," & _
                "'" & v_total_ot & "','" & chk_holiday.Value & "','" & IIf(opt_auto.Value, 0, IIf(opt_manual.Value, 1, 2)) & "',now(),'" & LOGIN_NAME & "')"
    ElseIf opt_manual.Value Then
        SQL = "INSERT INTO t_spl(company_code,employee_code,description,date,start_time,end_time," & _
              "flag_approval,ot_spl,flag_type, ot_basic_salary,ot_pembagi,ot_pengali,ot_hasil,entry_date,entry_user) " & _
            "VALUES( " & _
              "'" & TDBCombo_company.Text & "','" & txt_employee_code.Text & "','" & txt_description.Text & "'," & _
              "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(TDBTime1.Value, "yyyy-mm-dd hh:mm:ss") & "'," & _
              "'" & vEndTime & "',0,'" & Val(DropAllComma(txt_pengali.Text)) & "','" & IIf(opt_auto.Value, 0, IIf(opt_manual.Value, 1, 2)) & "'," & _
              "'" & DropAllComma(txt_basic_salary.Text) & "','" & txt_pembagi.Text & "'," & _
              "'" & txt_pengali.Text & "','" & DropAllComma(txt_ot_value.Text) & "',now(),'" & LOGIN_NAME & "')"
    Else
        SQL = "INSERT INTO t_spl(company_code,employee_code,description,date,start_time,end_time," & _
              "flag_approval,ot_spl,flag_type, ot_value1,ot_pengali1,ot_value2,ot_pengali2,ot_hasil_fixed,entry_date,entry_user) " & _
            "VALUES( " & _
              "'" & TDBCombo_company.Text & "','" & txt_employee_code.Text & "','" & txt_description.Text & "'," & _
              "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & Format(TDBTime1.Value, "yyyy-mm-dd hh:mm:ss") & "'," & _
              "'" & vEndTime & "',0,'" & (Val(DropAllComma(txt_pengali1.Text)) + Val(DropAllComma(txt_pengali2.Text))) & "','" & IIf(opt_auto.Value, 0, IIf(opt_manual.Value, 1, 2)) & "'," & _
              "'" & DropAllComma(txt_value1.Text) & "','" & DropAllComma(txt_pengali1.Text) & "'," & _
              "'" & DropAllComma(txt_value2.Text) & "','" & DropAllComma(txt_pengali2.Text) & "','" & DropAllComma(txt_ot_value_fixed.Text) & "',now(),'" & LOGIN_NAME & "')"
    End If
    CnG.Execute SQL

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
Dim vTgl As String
On Error GoTo Err
    
    vTgl = Format(TDBGrid1.Columns("date_").Value, "yyyy-MM-dd")
    SQL = "DELETE FROM t_spl WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND company_code = '" & TDBCombo_company.Text & "' " & _
            "AND date(date) = '" & vTgl & "'"
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
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" And Not LCase(Ctr.name) = "txt_division_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" And Not LCase(Ctr.name) = "tdbcombo_division" Then Ctr.Text = ""
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
    chk_holiday.Value = 0
    DTPicker1.Value = Now
    TDBTime1.Value = Now
    TDBTime2.Value = Now
    txt_description = ""
    
    txt_15.Text = 0
    txt_20.Text = 0
    txt_30.Text = 0
    txt_40.Text = 0
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        If TDBCombo_company.Text = "" Then
            MsgBox "Perusahaan Belum Dipilih...", vbExclamation, headerMSG
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

Private Sub DTPicker1_Validate(Cancel As Boolean)
    Call cari_jam
End Sub

Private Sub Form_Load()
    Call createGridKar
    Call load_data_company
    
    optApproval(0).Value = True
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
    
    fra_auto.Visible = True
    fra_manual.Visible = False
    fra_fixed.Visible = False
    btnApprove.Visible = False
        
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
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

Private Sub opt_auto_Click()
    txt_basic_salary.Text = 0
    txt_pengali.Text = 0
    txt_pembagi.Text = 0
    txt_ot_value.Text = 0
    
    txt_value1.Text = 0
    txt_pengali1.Text = 0
    txt_value2.Text = 0
    txt_pengali2.Text = 0
    txt_ot_value_fixed.Text = 0
            
    fra_auto.Visible = True
    fra_manual.Visible = False
    fra_fixed.Visible = False
    
    Call hitungLembur
End Sub

Private Sub opt_manual_Click()
    txt_15.Text = 0
    txt_20.Text = 0
    txt_30.Text = 0
    txt_40.Text = 0
    
    txt_value1.Text = 0
    txt_pengali1.Text = 0
    txt_value2.Text = 0
    txt_pengali2.Text = 0
    txt_ot_value_fixed.Text = 0
            
    fra_auto.Visible = False
    fra_manual.Visible = True
    fra_fixed.Visible = False
    
    Call hitungLembur
End Sub

Private Sub opt_fixed_Click()
    txt_15.Text = 0
    txt_20.Text = 0
    txt_30.Text = 0
    txt_40.Text = 0
    
    txt_basic_salary.Text = 0
    txt_pengali.Text = 0
    txt_pembagi.Text = 0
    txt_ot_value.Text = 0
    
    txt_value1.Text = 0
    txt_pengali1.Text = 0
    txt_value2.Text = 0
    txt_pengali2.Text = 0
    txt_ot_value_fixed.Text = 0
            
    fra_auto.Visible = False
    fra_manual.Visible = False
    fra_fixed.Visible = True
    
    Call hitungLembur
End Sub

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
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_division
'        Call load_data_spl
    End If
End Sub

Private Sub tdbcombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name = TDBCombo_division.Columns("division_name").Value
        
'        Call load_data_spl
    End If
End Sub

Private Sub tdbcombo_division_Change()
    If TDBCombo_division.Text = "" Then
        txt_division_name.Text = ""
        Call load_data_spl
    Else
        Call load_data_spl
    End If
End Sub

Private Sub load_data_spl()
    If rsSPL.State Then rsSPL.Close
    
    If TDBCombo_division.Text = "" Then
        SQL = "select a.*,b.nik,b.employee_name, date as date_, " & _
                "start_time as start_time_, " & _
                "end_time as end_time_ " & _
              "from t_spl a join m_employee b on a.employee_code = b.employee_code " & _
              "where b.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                "and (date(date) between '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' and '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & _
                "and ifnull(flag_approval,0) = " & IIf(optApproval(0).Value, 0, 1) & " " & _
              "order by date asc"
    Else
        SQL = "select a.*,b.nik,b.employee_name, date as date_, " & _
                "start_time as start_time_, " & _
                "end_time as end_time_ " & _
              "from t_spl a join m_employee b on a.employee_code = b.employee_code " & _
              "where b.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                "and (date(date) between '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' and '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & _
                "and b.division_code = '" & TDBCombo_division.Text & "' " & _
                "and ifnull(flag_approval,0) = " & IIf(optApproval(0).Value, 0, 1) & " " & _
              "order by date asc"
    End If
    rsSPL.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsSPL
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_division()
    If rsDivision.State Then rsDivision.Close
    SQL = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by division_code"
    rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDivision
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
    
    a = IIf(IsNull(TDBGrid1.Columns("flag_approval").Value), "0", TDBGrid1.Columns("flag_approval").Value)
    vFlagApproval = IIf(a = "", 0, a)
    
    SQL = "SELECT a.employee_code FROM m_form_approval_dtl a join m_user b on a.employee_code = b.employee_code " & _
            "WHERE a.form_name = 'frm_trans_spl' and b.user_name = '" & LOGIN_NAME & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Or LOGIN_LEVEL = 100 Then
        If vFlagApproval = 0 Then
            btnApprove.Visible = True
            lblVerify.Visible = False
        Else
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

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub hitungLembur()
    
    If opt_auto.Value Then
        If TDBTime2.Value < TDBTime1.Value Then
            jmlJam = IIf(IsNull(DateDiff("n", TDBTime1.Value, DateAdd("d", 1, TDBTime2.Value))), 0, DateDiff("n", TDBTime1.Value, DateAdd("d", 1, TDBTime2.Value)))
        Else
            jmlJam = IIf(IsNull(DateDiff("n", TDBTime1.Value, TDBTime2.Value)), 0, DateDiff("n", TDBTime1.Value, TDBTime2.Value))
        End If
        
            jmlJam = Round(jmlJam / 60, 2)
            
        If chk_holiday.Value = 0 Then
            txt_30.Text = 0
            txt_40.Text = 0
            
            If jmlJam >= 1 Then
                txt_15.Text = 1
                txt_20.Text = jmlJam - 1
            Else
                txt_15.Text = jmlJam
                txt_20.Text = 0
            End If
        Else
            txt_15.Text = 0
            If jmlJam = 7 Then
                txt_20.Text = 7
                txt_30.Text = 0
                txt_40.Text = 0
            ElseIf jmlJam > 7 And jmlJam < 8 Then
                txt_20.Text = 7
                txt_30.Text = jmlJam - 7
                txt_40.Text = 0
            ElseIf jmlJam = 8 Then
                txt_20.Text = 7
                txt_30.Text = 1
                txt_40.Text = 0
            ElseIf jmlJam > 8 Then
                txt_20.Text = 7
                txt_30.Text = 1
                txt_40.Text = jmlJam - 8
            Else
                txt_20.Text = jmlJam
                txt_30.Text = 0
                txt_40.Text = 0
            End If
        End If
    ElseIf opt_manual.Value Then
        SQL = "SELECT CASE WHEN DAYNAME(NOW()) = 'Sunday' THEN " & _
                        "CASE WHEN flag_basic = 1 THEN (basic_salary_sunday/30) ELSE basic_salary_sunday END " & _
                      "Else " & _
                        "CASE WHEN flag_basic = 1 THEN (basic_salary/30) ELSE basic_salary END " & _
                      "END basic_salary " & _
              "FROM m_salary_standard " & _
              "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
              "ORDER BY salary_date DESC LIMIT 1"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            txt_basic_salary.Text = FormatNumber(rs.Fields(0).Value)
        Else
            txt_basic_salary.Text = FormatNumber(0)
        End If
        rs.Close
        
        SQL = "SELECT ROUND(TIME_TO_SEC(TIMEDIFF(end_time,start_time)) / 3600) jam_pembagi " & _
              "FROM td_shift a JOIN tm_shift b ON a.group_code = b.group_code AND a.shift_number = b.shift_number " & _
              "JOIN m_shift c ON b.group_code = c.group_code AND b.shift_code = c.shift_code " & _
              "WHERE a.employee_code = '" & txt_employee_code.Text & "' LIMIT 1"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            txt_pembagi.Text = rs.Fields(0).Value
        Else
            txt_pembagi.Text = 1
        End If
        rs.Close
        
        If TDBTime2.Value < TDBTime1.Value Then
            jmlJam = IIf(IsNull(DateDiff("n", TDBTime1.Value, DateAdd("d", 1, TDBTime2.Value))), 0, DateDiff("n", TDBTime1.Value, DateAdd("d", 1, TDBTime2.Value)))
        Else
            jmlJam = IIf(IsNull(DateDiff("n", TDBTime1.Value, TDBTime2.Value)), 0, DateDiff("n", TDBTime1.Value, TDBTime2.Value))
        End If
        
        jmlJam = Round(jmlJam / 60, 2)
        txt_pengali.Text = jmlJam
        
        txt_ot_value.Text = (Val(DropAllComma(txt_basic_salary.Text)) / Val(txt_pembagi)) * Val(txt_pengali.Text)
        txt_ot_value.Text = FormatNumber(txt_ot_value.Text)
    Else
        jmlJam = Val(txt_pengali1.Text) + Val(txt_pengali2.Text)
        
        txt_ot_value_fixed.Text = (Val(DropAllComma(txt_value1.Text)) * Val(txt_pengali1)) + (Val(DropAllComma(txt_value2.Text)) * Val(txt_pengali2))
        txt_ot_value_fixed.Text = FormatNumber(txt_ot_value_fixed.Text)
    End If
End Sub


Private Sub txt_basic_salary_Validate(Cancel As Boolean)
    If txt_basic_salary <> "" Then
        txt_ot_value.Text = (Val(DropAllComma(txt_basic_salary.Text)) / Val(txt_pembagi)) * Val(txt_pengali.Text)
        txt_ot_value.Text = FormatNumber(txt_ot_value.Text)
        
        txt_basic_salary.Text = FormatNumber(txt_basic_salary.Text)
    End If
End Sub

Private Sub txt_pembagi_Validate(Cancel As Boolean)
    If txt_pembagi.Text <> "" Then
        txt_ot_value.Text = (Val(DropAllComma(txt_basic_salary.Text)) / Val(txt_pembagi)) * Val(txt_pengali.Text)
        txt_ot_value.Text = FormatNumber(txt_ot_value.Text)
        
        txt_pembagi.Text = FormatNumber(txt_pembagi.Text)
    End If
End Sub

Private Sub txt_pengali_Validate(Cancel As Boolean)
    If txt_pengali.Text <> "" Then
        txt_ot_value.Text = (Val(DropAllComma(txt_basic_salary.Text)) / Val(txt_pembagi)) * Val(txt_pengali.Text)
        txt_ot_value.Text = FormatNumber(txt_ot_value.Text)
        
        txt_pengali.Text = FormatNumber(txt_pengali.Text)
    End If
End Sub

Private Sub txt_value1_Validate(Cancel As Boolean)
    If txt_value1.Text <> "" Then
        txt_ot_value_fixed.Text = (Val(DropAllComma(txt_value1.Text)) * Val(txt_pengali1)) + (Val(DropAllComma(txt_value2.Text)) * Val(txt_pengali2))
        txt_ot_value_fixed.Text = FormatNumber(txt_ot_value_fixed.Text)
        
        txt_value1.Text = FormatNumber(txt_value1.Text)
    End If
End Sub

Private Sub txt_pengali1_Validate(Cancel As Boolean)
    If txt_pengali1.Text <> "" Then
        txt_ot_value_fixed.Text = (Val(DropAllComma(txt_value1.Text)) * Val(txt_pengali1)) + (Val(DropAllComma(txt_value2.Text)) * Val(txt_pengali2))
        txt_ot_value_fixed.Text = FormatNumber(txt_ot_value_fixed.Text)
        
        txt_pengali1.Text = FormatNumber(txt_pengali1.Text)
    End If
End Sub

Private Sub txt_value2_Validate(Cancel As Boolean)
    If txt_value2.Text <> "" Then
        txt_ot_value_fixed.Text = (Val(DropAllComma(txt_value1.Text)) * Val(txt_pengali1)) + (Val(DropAllComma(txt_value2.Text)) * Val(txt_pengali2))
        txt_ot_value_fixed.Text = FormatNumber(txt_ot_value_fixed.Text)
        
        txt_value2.Text = FormatNumber(txt_value2.Text)
    End If
End Sub

Private Sub txt_pengali2_Validate(Cancel As Boolean)
    If txt_pengali2.Text <> "" Then
        txt_ot_value_fixed.Text = (Val(DropAllComma(txt_value1.Text)) * Val(txt_pengali1)) + (Val(DropAllComma(txt_value2.Text)) * Val(txt_pengali2))
        txt_ot_value_fixed.Text = FormatNumber(txt_ot_value_fixed.Text)
        
        txt_pengali2.Text = FormatNumber(txt_pengali2.Text)
    End If
End Sub

Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "KODE KARY.", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 2000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "DIVISI", 1300, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "JABATAN", 1300, , , , , , , , , True
      .AddColumn "Employee Code", 3000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        If LOGIN_LEVEL = 100 Then
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_nik.Text & "%') " _
                        & "AND a.flag_active <> 0"
        Else
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_nik.Text & "%') " _
                        & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                        & "ORDER BY a.employee_name ASC"

        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME _
                                & vbTab & rs!division_code & vbTab & rs!division_name _
                                & vbTab & rs!title_code & vbTab & rs!title_name _
                                & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
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
            txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
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
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub

Private Sub cari_jam()
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

Dim vStatus As String
Dim v_shift As String
Dim vFlagHoliday As Integer

    SQL = "SELECT a.att_date,a.status FROM h_attendance a JOIN t_holiday b ON DATE(a.att_date) = DATE(b.holiday_date) " & _
          "WHERE Date(a.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
            "AND a.employee_code = '" & txt_employee_code.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
    
        '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT SBG PRESENT++++
        vStatus = IIf(IsNull(rs!Status), "H", rs!Status)
        
        If vStatus <> "H" Then
            v_shift = IIf(rs!Status = "A", "ALPHA", IIf(rs!Status = "T", "TUGAS DINAS", _
                    IIf(rs!Status = "L", "LIBUR", IIf(rs!Status = "I", "IJIN", "SAKIT"))))
            
            MsgBox "Karyawan Dalam Status " & v_shift & "! Bukan Dalam Status PRESENT!", vbCritical
            txt_employee_code.Text = "": txt_nik.Text = "": txt_employee_name.Text = ""
            rs.Close
            Exit Sub
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
        SQL = "SELECT time_in, time_out FROM h_attendance " & _
              "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs2.RecordCount > 0 Then
            chk_holiday.Value = 1
            TDBTime1.Text = Format(rs2!time_in, "hh:mm")
            TDBTime2.Text = Format(rs2!time_out, "hh:mm")
        Else
            chk_holiday.Value = 1
            TDBTime1.Text = Format(Now, "hh:mm")
            TDBTime2.Text = Format(Now, "hh:mm")
            
            txt_15.Text = 0
            txt_20.Text = 0
            txt_30.Text = 0
            txt_40.Text = 0
        End If
        rs2.Close
    Else
        SQL = "SELECT holiday_date FROM t_holiday " & _
              "WHERE Date(holiday_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs2.RecordCount > 0 Then
            vFlagHoliday = 1
        Else
            vFlagHoliday = 0
        End If
        rs2.Close
        
        SQL = "SELECT time_in, end_time, time_out, status FROM h_attendance " & _
              "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs2.RecordCount > 0 Then
            
            '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT SBG PRESENT++++
            vStatus = IIf(IsNull(rs2!Status), "H", rs2!Status)
            
            If vStatus <> "H" Then
                v_shift = IIf(rs2!Status = "A", "ALPHA", IIf(rs2!Status = "T", "TUGAS DINAS", _
                        IIf(rs2!Status = "L", "LIBUR", IIf(rs2!Status = "I", "IJIN", "SAKIT"))))
                
                MsgBox "Karyawan Dalam Status " & v_shift & "! Bukan Dalam Status PRESENT!", vbCritical
                txt_employee_code.Text = "": txt_nik.Text = "": txt_employee_name.Text = ""
                rs2.Close
                rs.Close
                Exit Sub
            End If
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            If Format(DTPicker1.Value, "dddd") = "Sunday" Then
                chk_holiday.Value = 1
                TDBTime1.Text = Format(rs2!time_in, "hh:mm")
                TDBTime2.Text = Format(rs2!time_out, "hh:mm")
            Else
                chk_holiday.Value = vFlagHoliday
                TDBTime1.Text = Format(rs2!end_time, "hh:mm")
                TDBTime2.Text = Format(rs2!time_out, "hh:mm")
            End If
        Else
            chk_holiday.Value = vFlagHoliday
            TDBTime1.Text = Format(Now, "hh:mm")
            TDBTime2.Text = Format(Now, "hh:mm")
            
            txt_15.Text = 0
            txt_20.Text = 0
            txt_30.Text = 0
            txt_40.Text = 0
        End If
        rs2.Close
    End If
    rs.Close
End Sub

Private Sub txt_nik_Validate(Cancel As Boolean)
    Call cari_jam
End Sub

Private Sub vbButton1_Click()
    Call load_data_spl
End Sub
