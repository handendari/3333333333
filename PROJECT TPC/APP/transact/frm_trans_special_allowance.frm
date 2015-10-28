VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_special_allowance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPECIAL ALLOWANCE"
   ClientHeight    =   8925
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_employee_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   5250
      TabIndex        =   119
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox txt_employee_code 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8610
      TabIndex        =   122
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3330
      TabIndex        =   120
      Top             =   3360
      Width           =   1515
   End
   Begin prj_tpc.vbButton cmdBrowse 
      Height          =   285
      Left            =   4890
      TabIndex        =   117
      Top             =   3360
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "..."
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
      MICON           =   "frm_trans_special_allowance.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   3330
      TabIndex        =   118
      Top             =   3660
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5159
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
      GridLines       =   2
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3090
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   36
      Top             =   690
      Width           =   3855
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   90
      Top             =   8190
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10650
      TabIndex        =   19
      Top             =   8190
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
      MICON           =   "frm_trans_special_allowance.frx":001C
      PICN            =   "frm_trans_special_allowance.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1320
      OleObjectBlob   =   "frm_trans_special_allowance.frx":10CA
      TabIndex        =   37
      Top             =   690
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker_special_from 
      Height          =   315
      Left            =   1320
      TabIndex        =   91
      Top             =   1410
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   92798979
      CurrentDate     =   39270
   End
   Begin MSComCtl2.DTPicker DTPicker_special_to 
      Height          =   315
      Left            =   3240
      TabIndex        =   98
      Top             =   1410
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   92798979
      CurrentDate     =   39270
   End
   Begin prj_tpc.vbButton cmdSearch 
      Height          =   465
      Left            =   5010
      TabIndex        =   116
      Top             =   1110
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
      MICON           =   "frm_trans_special_allowance.frx":3030
      PICN            =   "frm_trans_special_allowance.frx":304C
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
      Left            =   1320
      TabIndex        =   202
      Top             =   1050
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM-yyyy"
      Format          =   92798979
      CurrentDate     =   39270
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "MARRIAGE ALLOW."
      TabPicture(0)   =   "frm_trans_special_allowance.frx":40DE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid_Marriage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_entry_marriage"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DELIVER BABY ALLOW."
      TabPicture(1)   =   "frm_trans_special_allowance.frx":40FA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_entry_baby"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TDBGrid_Baby"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "CONDOLENCE ALLOW."
      TabPicture(2)   =   "frm_trans_special_allowance.frx":4116
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fra_entry_condolence"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "TDBGrid_Condolence"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "DEATH ALLOW."
      TabPicture(3)   =   "frm_trans_special_allowance.frx":4132
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra_entry_death"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TDBGrid_Death"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "TRAVELLING ALLOW."
      TabPicture(4)   =   "frm_trans_special_allowance.frx":414E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fra_entry_travelling"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "TDBGrid_Travelling"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "ANNUALY LEAVE COMPENSATION"
      TabPicture(5)   =   "frm_trans_special_allowance.frx":416A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "fra_entry_leave"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "TDBGrid_Leave"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame fra_entry_death 
         Height          =   4005
         Left            =   -74850
         TabIndex        =   152
         Top             =   840
         Width           =   11295
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   3
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   196
            Top             =   3210
            Width           =   2835
         End
         Begin VB.TextBox txt_death_salary 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   61
            Top             =   1050
            Width           =   1545
         End
         Begin VB.TextBox txt_death_allowance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   64
            Top             =   2130
            Width           =   1515
         End
         Begin VB.CommandButton Command4 
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
            TabIndex        =   153
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_death_age 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   63
            Top             =   1770
            Width           =   1515
         End
         Begin VB.TextBox txt_death_words 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   65
            Top             =   2490
            Width           =   7575
         End
         Begin VB.TextBox txt_death_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   66
            Top             =   2850
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker_trans_death 
            Height          =   315
            Left            =   3060
            TabIndex        =   60
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin MSComCtl2.DTPicker DTPicker_Death 
            Height          =   315
            Left            =   3060
            TabIndex        =   62
            Top             =   1410
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin VB.Frame Frame10 
            Height          =   465
            Left            =   3060
            TabIndex        =   195
            Top             =   3120
            Width           =   2955
            Begin VB.OptionButton opt_formula 
               Caption         =   "FORMULA"
               Height          =   195
               Index           =   3
               Left            =   1350
               TabIndex        =   68
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton opt_manual 
               Caption         =   "MANUAL"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   67
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_pph21_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   6060
            MaxLength       =   50
            TabIndex        =   69
            Top             =   3210
            Width           =   1245
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
            Height          =   375
            Index           =   3
            Left            =   6060
            OleObjectBlob   =   "frm_trans_special_allowance.frx":4186
            TabIndex        =   76
            Top             =   3210
            Width           =   1275
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SALARY"
            Height          =   195
            Left            =   2310
            TabIndex        =   184
            Top             =   1110
            Width           =   570
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANS DATE"
            Height          =   195
            Left            =   1950
            TabIndex        =   161
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   1995
            TabIndex        =   160
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "WORKING AGE"
            Height          =   195
            Left            =   1770
            TabIndex        =   159
            Top             =   1800
            Width           =   1080
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   1830
            TabIndex        =   158
            Top             =   2160
            Width           =   1020
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DATE OF PASSED AWAY"
            Height          =   195
            Left            =   1125
            TabIndex        =   157
            Top             =   1470
            Width           =   1755
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BY WORDS"
            Height          =   195
            Left            =   2055
            TabIndex        =   156
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   155
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   2355
            TabIndex        =   154
            Top             =   3240
            Width           =   495
         End
      End
      Begin VB.Frame fra_entry_marriage 
         Height          =   4005
         Left            =   150
         TabIndex        =   39
         Top             =   840
         Width           =   11295
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   0
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   191
            Top             =   3210
            Width           =   2835
         End
         Begin VB.TextBox txt_marriage_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2850
            Width           =   3495
         End
         Begin VB.TextBox txt_marriage_words 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   124
            Top             =   2490
            Width           =   7575
         End
         Begin VB.TextBox txt_marriage_wife_husband 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1770
            Width           =   3495
         End
         Begin VB.TextBox txt_marriage_place 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1050
            Width           =   3495
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
            TabIndex        =   50
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_marriage_allowance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2130
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker DTPicker_trans_marriage 
            Height          =   315
            Left            =   3060
            TabIndex        =   1
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin MSComCtl2.DTPicker DTPicker_marriage 
            Height          =   315
            Left            =   3060
            TabIndex        =   3
            Top             =   1410
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin VB.Frame Frame7 
            Height          =   465
            Left            =   3060
            TabIndex        =   189
            Top             =   3120
            Width           =   2955
            Begin VB.OptionButton opt_formula 
               Caption         =   "FORMULA"
               Height          =   195
               Index           =   0
               Left            =   1350
               TabIndex        =   8
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton opt_manual 
               Caption         =   "MANUAL"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_pph21_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   6060
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3210
            Width           =   1245
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
            Height          =   375
            Index           =   0
            Left            =   6060
            OleObjectBlob   =   "frm_trans_special_allowance.frx":60E5
            TabIndex        =   10
            Top             =   3210
            Width           =   1275
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   2355
            TabIndex        =   126
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   125
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BY WORDS"
            Height          =   195
            Left            =   2055
            TabIndex        =   123
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DATE / SERIES"
            Height          =   195
            Left            =   1815
            TabIndex        =   121
            Top             =   1470
            Width           =   1065
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   1830
            TabIndex        =   77
            Top             =   2160
            Width           =   1020
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "WIFE/HUSBAND NAME"
            Height          =   195
            Left            =   1230
            TabIndex        =   70
            Top             =   1800
            Width           =   1620
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MARRIAGE PLACE"
            Height          =   195
            Left            =   1575
            TabIndex        =   59
            Top             =   1110
            Width           =   1305
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   1995
            TabIndex        =   58
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANS DATE"
            Height          =   195
            Left            =   1950
            TabIndex        =   57
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   17
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew 
            Height          =   705
            Index           =   5
            Left            =   360
            TabIndex        =   110
            Top             =   390
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
            MICON           =   "frm_trans_special_allowance.frx":8044
            PICN            =   "frm_trans_special_allowance.frx":8060
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
            Index           =   5
            Left            =   1380
            TabIndex        =   111
            Top             =   390
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
            MICON           =   "frm_trans_special_allowance.frx":90F2
            PICN            =   "frm_trans_special_allowance.frx":910E
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
            Index           =   5
            Left            =   2400
            TabIndex        =   112
            Top             =   390
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
            MICON           =   "frm_trans_special_allowance.frx":A1A0
            PICN            =   "frm_trans_special_allowance.frx":A1BC
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
            Index           =   5
            Left            =   3420
            TabIndex        =   113
            Top             =   390
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
            MICON           =   "frm_trans_special_allowance.frx":B24E
            PICN            =   "frm_trans_special_allowance.frx":B26A
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
            Index           =   5
            Left            =   4440
            TabIndex        =   114
            Top             =   390
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
            MICON           =   "frm_trans_special_allowance.frx":C2FC
            PICN            =   "frm_trans_special_allowance.frx":C318
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
      Begin VB.Frame fra_entry_leave 
         Height          =   4005
         Left            =   -74850
         TabIndex        =   176
         Top             =   840
         Width           =   11295
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   5
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   199
            Top             =   2250
            Width           =   2835
         End
         Begin VB.TextBox txt_leave_allowance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5790
            MaxLength       =   50
            TabIndex        =   103
            Top             =   1170
            Width           =   1845
         End
         Begin VB.TextBox txt_leave_pengali 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   102
            Top             =   1170
            Width           =   525
         End
         Begin VB.TextBox txt_leave_salary 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   101
            Top             =   1170
            Width           =   1515
         End
         Begin VB.CommandButton Command6 
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
            TabIndex        =   177
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_leave_words 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   104
            Top             =   1530
            Width           =   7575
         End
         Begin VB.TextBox txt_leave_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   105
            Top             =   1890
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker_trans_leave 
            Height          =   315
            Left            =   3060
            TabIndex        =   100
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin VB.Frame Frame12 
            Height          =   465
            Left            =   3060
            TabIndex        =   200
            Top             =   2160
            Width           =   2955
            Begin VB.OptionButton opt_manual 
               Caption         =   "MANUAL"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   106
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton opt_formula 
               Caption         =   "FORMULA"
               Height          =   195
               Index           =   5
               Left            =   1350
               TabIndex        =   107
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_pph21_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   6060
            MaxLength       =   50
            TabIndex        =   108
            Top             =   2250
            Width           =   1245
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
            Height          =   375
            Index           =   5
            Left            =   6060
            OleObjectBlob   =   "frm_trans_special_allowance.frx":D3AA
            TabIndex        =   115
            Top             =   2250
            Width           =   1275
         End
         Begin VB.Label Label54 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5550
            TabIndex        =   188
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label52 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            TabIndex        =   187
            Top             =   1230
            Width           =   195
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANS DATE"
            Height          =   195
            Left            =   1950
            TabIndex        =   183
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   1995
            TabIndex        =   182
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   1830
            TabIndex        =   181
            Top             =   1200
            Width           =   1020
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BY WORDS"
            Height          =   195
            Left            =   2055
            TabIndex        =   180
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   179
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   2355
            TabIndex        =   178
            Top             =   2280
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   174
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew 
            Height          =   705
            Index           =   4
            Left            =   360
            TabIndex        =   92
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
            MICON           =   "frm_trans_special_allowance.frx":F309
            PICN            =   "frm_trans_special_allowance.frx":F325
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
            Index           =   4
            Left            =   1380
            TabIndex        =   93
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
            MICON           =   "frm_trans_special_allowance.frx":103B7
            PICN            =   "frm_trans_special_allowance.frx":103D3
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
            Index           =   4
            Left            =   2400
            TabIndex        =   94
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
            MICON           =   "frm_trans_special_allowance.frx":11465
            PICN            =   "frm_trans_special_allowance.frx":11481
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
            Index           =   4
            Left            =   3420
            TabIndex        =   95
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
            MICON           =   "frm_trans_special_allowance.frx":12513
            PICN            =   "frm_trans_special_allowance.frx":1252F
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
            Index           =   4
            Left            =   4440
            TabIndex        =   96
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
            MICON           =   "frm_trans_special_allowance.frx":135C1
            PICN            =   "frm_trans_special_allowance.frx":135DD
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
      Begin VB.Frame fra_entry_travelling 
         Height          =   4035
         Left            =   -74850
         TabIndex        =   164
         Top             =   810
         Width           =   11295
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   4
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   197
            Top             =   3570
            Width           =   2835
         End
         Begin VB.TextBox txt_travel_requirement 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   84
            Top             =   2130
            Width           =   3495
         End
         Begin VB.TextBox txt_travel_address 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   82
            Top             =   1410
            Width           =   3495
         End
         Begin VB.TextBox txt_travel_allowance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   85
            Top             =   2490
            Width           =   1515
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   165
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_travel_long 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   81
            Top             =   1050
            Width           =   1515
         End
         Begin VB.TextBox txt_travel_country 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   83
            Top             =   1770
            Width           =   3495
         End
         Begin VB.TextBox txt_travel_words 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   86
            Top             =   2850
            Width           =   7575
         End
         Begin VB.TextBox txt_travel_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   87
            Top             =   3210
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker_trans_travel 
            Height          =   315
            Left            =   3060
            TabIndex        =   80
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin VB.Frame Frame11 
            Height          =   465
            Left            =   3060
            TabIndex        =   198
            Top             =   3480
            Width           =   2955
            Begin VB.OptionButton opt_manual 
               Caption         =   "MANUAL"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   88
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton opt_formula 
               Caption         =   "FORMULA"
               Height          =   195
               Index           =   4
               Left            =   1350
               TabIndex        =   89
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_pph21_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   6060
            MaxLength       =   50
            TabIndex        =   90
            Top             =   3570
            Width           =   1245
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
            Height          =   375
            Index           =   4
            Left            =   6060
            OleObjectBlob   =   "frm_trans_special_allowance.frx":1466F
            TabIndex        =   97
            Top             =   3570
            Width           =   1275
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "REQUIREMENT"
            Height          =   195
            Left            =   1785
            TabIndex        =   186
            Top             =   2190
            Width           =   1080
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESTINATION ADDRESS"
            Height          =   195
            Left            =   1095
            TabIndex        =   185
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANS DATE"
            Height          =   195
            Left            =   1950
            TabIndex        =   173
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   1995
            TabIndex        =   172
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LONG DUTY"
            Height          =   195
            Left            =   2040
            TabIndex        =   171
            Top             =   1110
            Width           =   840
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESTINATION COUNTRY"
            Height          =   195
            Left            =   1065
            TabIndex        =   170
            Top             =   1800
            Width           =   1785
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   1830
            TabIndex        =   169
            Top             =   2520
            Width           =   1020
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BY WORDS"
            Height          =   195
            Left            =   2055
            TabIndex        =   168
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   167
            Top             =   3240
            Width           =   1020
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   2355
            TabIndex        =   166
            Top             =   3600
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   162
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew 
            Height          =   705
            Index           =   3
            Left            =   360
            TabIndex        =   71
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
            MICON           =   "frm_trans_special_allowance.frx":165CE
            PICN            =   "frm_trans_special_allowance.frx":165EA
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
            Index           =   3
            Left            =   1380
            TabIndex        =   72
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
            MICON           =   "frm_trans_special_allowance.frx":1767C
            PICN            =   "frm_trans_special_allowance.frx":17698
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
            Index           =   3
            Left            =   2400
            TabIndex        =   73
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
            MICON           =   "frm_trans_special_allowance.frx":1872A
            PICN            =   "frm_trans_special_allowance.frx":18746
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
            Index           =   3
            Left            =   3420
            TabIndex        =   74
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
            MICON           =   "frm_trans_special_allowance.frx":197D8
            PICN            =   "frm_trans_special_allowance.frx":197F4
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
            Index           =   3
            Left            =   4440
            TabIndex        =   75
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
            MICON           =   "frm_trans_special_allowance.frx":1A886
            PICN            =   "frm_trans_special_allowance.frx":1A8A2
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
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   150
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew 
            Height          =   705
            Index           =   2
            Left            =   330
            TabIndex        =   51
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
            MICON           =   "frm_trans_special_allowance.frx":1B934
            PICN            =   "frm_trans_special_allowance.frx":1B950
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
            Index           =   2
            Left            =   1350
            TabIndex        =   52
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
            MICON           =   "frm_trans_special_allowance.frx":1C9E2
            PICN            =   "frm_trans_special_allowance.frx":1C9FE
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
            Index           =   2
            Left            =   2370
            TabIndex        =   53
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
            MICON           =   "frm_trans_special_allowance.frx":1DA90
            PICN            =   "frm_trans_special_allowance.frx":1DAAC
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
            Index           =   2
            Left            =   3390
            TabIndex        =   54
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
            MICON           =   "frm_trans_special_allowance.frx":1EB3E
            PICN            =   "frm_trans_special_allowance.frx":1EB5A
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
            Index           =   2
            Left            =   4410
            TabIndex        =   55
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
            MICON           =   "frm_trans_special_allowance.frx":1FBEC
            PICN            =   "frm_trans_special_allowance.frx":1FC08
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
      Begin VB.Frame fra_entry_condolence 
         Height          =   4005
         Left            =   -74850
         TabIndex        =   139
         Top             =   840
         Width           =   11295
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   2
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   194
            Top             =   3210
            Width           =   2835
         End
         Begin VB.TextBox txt_condolence_allowance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   44
            Top             =   2130
            Width           =   1515
         End
         Begin VB.CommandButton Command3 
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
            TabIndex        =   140
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_condolence_place 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   42
            Top             =   1410
            Width           =   3495
         End
         Begin VB.TextBox txt_condolence_family 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   41
            Top             =   1050
            Width           =   3495
         End
         Begin VB.TextBox txt_condolence_words 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   45
            Top             =   2490
            Width           =   7575
         End
         Begin VB.TextBox txt_condolence_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   46
            Top             =   2850
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker_trans_condolence 
            Height          =   315
            Left            =   3060
            TabIndex        =   40
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin MSComCtl2.DTPicker DTPicker_Condolence 
            Height          =   315
            Left            =   3060
            TabIndex        =   43
            Top             =   1770
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin VB.Frame Frame9 
            Height          =   465
            Left            =   3060
            TabIndex        =   193
            Top             =   3120
            Width           =   2955
            Begin VB.OptionButton opt_manual 
               Caption         =   "MANUAL"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   47
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton opt_formula 
               Caption         =   "FORMULA"
               Height          =   195
               Index           =   2
               Left            =   1350
               TabIndex        =   48
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_pph21_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   6060
            MaxLength       =   50
            TabIndex        =   49
            Top             =   3210
            Width           =   1245
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
            Height          =   375
            Index           =   2
            Left            =   6060
            OleObjectBlob   =   "frm_trans_special_allowance.frx":20C9A
            TabIndex        =   56
            Top             =   3210
            Width           =   1275
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANS DATE"
            Height          =   195
            Left            =   1950
            TabIndex        =   149
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   1995
            TabIndex        =   148
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DEATH PLACE"
            Height          =   195
            Left            =   1875
            TabIndex        =   147
            Top             =   1470
            Width           =   1005
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FAMILY NAME"
            Height          =   195
            Left            =   1845
            TabIndex        =   146
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   1830
            TabIndex        =   145
            Top             =   2160
            Width           =   1020
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DATE"
            Height          =   195
            Left            =   2490
            TabIndex        =   144
            Top             =   1830
            Width           =   390
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BY WORDS"
            Height          =   195
            Left            =   2055
            TabIndex        =   143
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   142
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   2355
            TabIndex        =   141
            Top             =   3240
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74850
         TabIndex        =   137
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew 
            Height          =   705
            Index           =   1
            Left            =   330
            TabIndex        =   31
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
            MICON           =   "frm_trans_special_allowance.frx":22BF9
            PICN            =   "frm_trans_special_allowance.frx":22C15
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
            Index           =   1
            Left            =   1350
            TabIndex        =   32
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
            MICON           =   "frm_trans_special_allowance.frx":23CA7
            PICN            =   "frm_trans_special_allowance.frx":23CC3
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
            Index           =   1
            Left            =   2370
            TabIndex        =   33
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
            MICON           =   "frm_trans_special_allowance.frx":24D55
            PICN            =   "frm_trans_special_allowance.frx":24D71
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
            Index           =   1
            Left            =   3390
            TabIndex        =   34
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
            MICON           =   "frm_trans_special_allowance.frx":25E03
            PICN            =   "frm_trans_special_allowance.frx":25E1F
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
            Index           =   1
            Left            =   4410
            TabIndex        =   35
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
            MICON           =   "frm_trans_special_allowance.frx":26EB1
            PICN            =   "frm_trans_special_allowance.frx":26ECD
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
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   150
         TabIndex        =   78
         Top             =   4830
         Width           =   11295
         Begin prj_tpc.vbButton cmdNew 
            Height          =   705
            Index           =   0
            Left            =   300
            TabIndex        =   11
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
            MICON           =   "frm_trans_special_allowance.frx":27F5F
            PICN            =   "frm_trans_special_allowance.frx":27F7B
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
            Index           =   0
            Left            =   1320
            TabIndex        =   12
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
            MICON           =   "frm_trans_special_allowance.frx":2900D
            PICN            =   "frm_trans_special_allowance.frx":29029
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
            Index           =   0
            Left            =   2340
            TabIndex        =   13
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
            MICON           =   "frm_trans_special_allowance.frx":2A0BB
            PICN            =   "frm_trans_special_allowance.frx":2A0D7
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
            Index           =   0
            Left            =   3360
            TabIndex        =   14
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
            MICON           =   "frm_trans_special_allowance.frx":2B169
            PICN            =   "frm_trans_special_allowance.frx":2B185
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
            Index           =   0
            Left            =   4380
            TabIndex        =   15
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
            MICON           =   "frm_trans_special_allowance.frx":2C217
            PICN            =   "frm_trans_special_allowance.frx":2C233
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Marriage 
         Height          =   4245
         Left            =   150
         TabIndex        =   79
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TRANS DATE"
         Columns(0).DataField=   "trans_date"
         Columns(0).NumberFormat=   "yyyy-MM-dd"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMP. ID. NO"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "EMPLOYEE NAME"
         Columns(2).DataField=   "employee_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "MARRIAGE PLACE"
         Columns(3).DataField=   "place"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DATE/SERIES"
         Columns(4).DataField=   "date"
         Columns(4).NumberFormat=   "dd-MM-yyyy"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "WIFE / HUSBAND"
         Columns(5).DataField=   "wife_husband_name"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ALLOWANCE"
         Columns(6).DataField=   "sisa_allowance"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "DESCRIPTION"
         Columns(7).DataField=   "description"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "EMP. CODE"
         Columns(8).DataField=   "employee_code"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3651"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3572"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2672"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2593"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2963"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2884"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2196"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2117"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=4763"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4683"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
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
         Caption         =   "LIST OF MARIAGE ALLOWANCE"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=3"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Condolence 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   151
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TRANS DATE"
         Columns(0).DataField=   "trans_date"
         Columns(0).NumberFormat=   "dd-MM-yyyy"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMP. ID. NO"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "EMPLOYEE NAME"
         Columns(2).DataField=   "employee_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "FAMILY NAME"
         Columns(3).DataField=   "family_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DEATH PLACE"
         Columns(4).DataField=   "death_place"
         Columns(4).NumberFormat=   "yyyy-MM-dd"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DATE"
         Columns(5).DataField=   "death_date"
         Columns(5).NumberFormat=   "dd-MM-yyyy"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ALLOWANCE"
         Columns(6).DataField=   "sisa_allowance"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "DESCRIPTION"
         Columns(7).DataField=   "description"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "EMP. CODE"
         Columns(8).DataField=   "employee_code"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3757"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3678"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2196"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2117"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=4763"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4683"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
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
         Caption         =   "LIST OF CONDOLENCE ALLOWANCE"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=3"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Death 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   163
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TRANS DATE"
         Columns(0).DataField=   "trans_date"
         Columns(0).NumberFormat=   "dd-MM-yyyy"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMP. ID. NO"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "EMPLOYEE NAME"
         Columns(2).DataField=   "employee_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "WORKING AGE"
         Columns(3).DataField=   "masa_kerja"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DATE OF PASSED AWAY"
         Columns(4).DataField=   "date_pass_away"
         Columns(4).NumberFormat=   "dd-MM-yyyy"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ALLOWANCE"
         Columns(5).DataField=   "sisa_allowance"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DESCRIPTION"
         Columns(6).DataField=   "description"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "EMP. CODE"
         Columns(7).DataField=   "employee_code"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3757"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3678"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2196"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2117"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=4763"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=4683"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
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
         Caption         =   "LIST OF DEATH ALLOWANCE"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=3"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Travelling 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   175
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TRANS DATE"
         Columns(0).DataField=   "trans_date"
         Columns(0).NumberFormat=   "dd-MM-yyyy"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMP. ID. NO"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "EMPLOYEE NAME"
         Columns(2).DataField=   "employee_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "LONG DUTY"
         Columns(3).DataField=   "lama_tugas"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DEST. ADDRESS"
         Columns(4).DataField=   "alamat_tujuan"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DEST. COUNTRY"
         Columns(5).DataField=   "negara_tujuan"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "REQUIREMENT"
         Columns(6).DataField=   "keperluan"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "ALLOWANCE"
         Columns(7).DataField=   "sisa_allowance"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "DESCRIPTION"
         Columns(8).DataField=   "description"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "EMP. CODE"
         Columns(9).DataField=   "employee_code"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3757"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3678"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2963"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2884"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2196"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2117"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=4763"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=4683"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(50)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
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
         Caption         =   "LIST OF TRAVELLING ALLOWANCE"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=3"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(74)  =   "Named:id=33:Normal"
         _StyleDefs(75)  =   ":id=33,.parent=0"
         _StyleDefs(76)  =   "Named:id=34:Heading"
         _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   ":id=34,.wraptext=-1"
         _StyleDefs(79)  =   "Named:id=35:Footing"
         _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(81)  =   "Named:id=36:Selected"
         _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(83)  =   "Named:id=37:Caption"
         _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(85)  =   "Named:id=38:HighlightRow"
         _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(87)  =   "Named:id=39:EvenRow"
         _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(89)  =   "Named:id=40:OddRow"
         _StyleDefs(90)  =   ":id=40,.parent=33"
         _StyleDefs(91)  =   "Named:id=41:RecordSelector"
         _StyleDefs(92)  =   ":id=41,.parent=34"
         _StyleDefs(93)  =   "Named:id=42:FilterBar"
         _StyleDefs(94)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Leave 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   18
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TRANS DATE"
         Columns(0).DataField=   "trans_date"
         Columns(0).NumberFormat=   "dd-MM-yyyy"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMP. ID. NO"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "EMPLOYEE NAME"
         Columns(2).DataField=   "employee_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "BASIC SALARY"
         Columns(3).DataField=   "basic_salary"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MULTIPLIERS"
         Columns(4).DataField=   "pengali"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ALLOWANCE"
         Columns(5).DataField=   "sisa_allowance"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DESCRIPTION"
         Columns(6).DataField=   "description"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "EMP. CODE"
         Columns(7).DataField=   "employee_code"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3757"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3678"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2196"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2117"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=4763"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=4683"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
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
         Caption         =   "LIST OF ANNUALY LEAVE COMPENSATION"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=3"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
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
      Begin VB.Frame fra_entry_baby 
         Height          =   4005
         Left            =   -74850
         TabIndex        =   127
         Top             =   840
         Width           =   11295
         Begin VB.TextBox txt_baby_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1050
            Width           =   3495
         End
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   1
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   192
            Top             =   3210
            Width           =   2835
         End
         Begin VB.TextBox txt_baby_allowance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   24
            Top             =   2130
            Width           =   1515
         End
         Begin VB.CommandButton Command2 
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
            TabIndex        =   128
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_baby_place 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1410
            Width           =   3495
         End
         Begin VB.TextBox txt_baby_words 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   25
            Top             =   2490
            Width           =   7575
         End
         Begin VB.TextBox txt_baby_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            MaxLength       =   50
            TabIndex        =   26
            Top             =   2850
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPicker_trans_baby 
            Height          =   315
            Left            =   3060
            TabIndex        =   20
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin MSComCtl2.DTPicker DTPicker_baby 
            Height          =   315
            Left            =   3060
            TabIndex        =   23
            Top             =   1770
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   92798979
            CurrentDate     =   41213
         End
         Begin VB.Frame Frame8 
            Height          =   465
            Index           =   0
            Left            =   3060
            TabIndex        =   190
            Top             =   3120
            Width           =   2955
            Begin VB.OptionButton opt_manual 
               Caption         =   "MANUAL"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton opt_formula 
               Caption         =   "FORMULA"
               Height          =   195
               Index           =   1
               Left            =   1350
               TabIndex        =   28
               Top             =   180
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_pph21_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   6060
            MaxLength       =   50
            TabIndex        =   29
            Top             =   3210
            Width           =   1245
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph21 
            Height          =   375
            Index           =   1
            Left            =   6060
            OleObjectBlob   =   "frm_trans_special_allowance.frx":2D2C5
            TabIndex        =   30
            Top             =   3210
            Width           =   1275
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CHILDREN NAME"
            Height          =   195
            Left            =   1665
            TabIndex        =   201
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TRANS DATE"
            Height          =   195
            Left            =   1950
            TabIndex        =   136
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE*"
            Height          =   195
            Left            =   1995
            TabIndex        =   135
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BIRTH PLACE"
            Height          =   195
            Left            =   1920
            TabIndex        =   134
            Top             =   1470
            Width           =   960
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALLOWANCE*"
            Height          =   195
            Left            =   1830
            TabIndex        =   133
            Top             =   2160
            Width           =   1020
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DATE OF BIRTH"
            Height          =   195
            Left            =   1740
            TabIndex        =   132
            Top             =   1830
            Width           =   1140
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BY WORDS"
            Height          =   195
            Left            =   2055
            TabIndex        =   131
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1830
            TabIndex        =   130
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PPh 21"
            Height          =   195
            Left            =   2355
            TabIndex        =   129
            Top             =   3240
            Width           =   495
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Baby 
         Height          =   4245
         Left            =   -74850
         TabIndex        =   138
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TRANS DATE"
         Columns(0).DataField=   "trans_date"
         Columns(0).NumberFormat=   "dd-MM-yyyy"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "EMP. ID. NO"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "EMPLOYEE NAME"
         Columns(2).DataField=   "employee_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "CHIDREN NAME"
         Columns(3).DataField=   "child_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "BIRTH PLACE"
         Columns(4).DataField=   "birth_place"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DATE OF BIRTH"
         Columns(5).DataField=   "date_birth"
         Columns(5).NumberFormat=   "dd-MM-yyyy"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ALLOWANCE"
         Columns(6).DataField=   "sisa_allowance"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "DESCRIPTION"
         Columns(7).DataField=   "description"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "EMP. CODE"
         Columns(8).DataField=   "employee_code"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3757"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3678"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2196"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2117"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=4763"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4683"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
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
         Caption         =   "LIST OF DELIVER BABY ALLOWANCE"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=3"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
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
   End
   Begin VB.Label Label58 
      AutoSize        =   -1  'True
      Caption         =   "PERIODE"
      Height          =   195
      Left            =   150
      TabIndex        =   203
      Top             =   1110
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2910
      TabIndex        =   109
      Top             =   1500
      Width           =   270
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      Caption         =   "RANGE DATE"
      Height          =   195
      Left            =   150
      TabIndex        =   99
      Top             =   1470
      Width           =   945
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   150
      TabIndex        =   38
      Top             =   750
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SPECIAL ALLOWANCE"
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
      TabIndex        =   0
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   -30
      Picture         =   "frm_trans_special_allowance.frx":2F224
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frm_trans_special_allowance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsMarriage As New ADODB.Recordset
Dim rsBaby As New ADODB.Recordset
Dim rsCondolence As New ADODB.Recordset
Dim rsDeath As New ADODB.Recordset
Dim rsTravel As New ADODB.Recordset
Dim rsLeave As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsPPh As New ADODB.Recordset

Dim x As Integer
Dim vTransDate As String
Dim vTitleCode As String
Dim vBasicSalary As String
Dim vStartWorking As String

Dim oClause As String

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False

    If SSTab1.Tab = 0 Then
        If rs.State Then rs.Close
        str_sql = "select count(employee_code) as rec_count from t_special_marriage " & _
                    "where employee_code = '" & txt_employee_code.Text & "' " & _
                        "and date(trans_date) = '" & Format(DTPicker_trans_marriage.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If

        rs.Close
    ElseIf SSTab1.Tab = 1 Then
        If rs.State Then rs.Close
        str_sql = "select count(employee_code) as rec_count from t_special_baby " & _
                    "where employee_code = '" & txt_employee_code.Text & "' " & _
                        "and date(trans_date) = '" & Format(DTPicker_trans_baby.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If

        rs.Close
    ElseIf SSTab1.Tab = 2 Then
        If rs.State Then rs.Close
        str_sql = "select count(employee_code) as rec_count from t_special_condolence " & _
                    "where employee_code = '" & txt_employee_code.Text & "' " & _
                        "and date(trans_date) = '" & Format(DTPicker_trans_condolence.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If

        rs.Close
    ElseIf SSTab1.Tab = 3 Then
        If rs.State Then rs.Close
        str_sql = "select count(employee_code) as rec_count from t_special_death " & _
                    "where employee_code = '" & txt_employee_code.Text & "' " & _
                        "and date(trans_date) = '" & Format(DTPicker_trans_death.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If

        rs.Close
    ElseIf SSTab1.Tab = 4 Then
        If rs.State Then rs.Close
        str_sql = "select count(employee_code) as rec_count from t_special_travelling " & _
                    "where employee_code = '" & txt_employee_code.Text & "' " & _
                        "and date(trans_date) = '" & Format(DTPicker_trans_travel.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If

        rs.Close
    ElseIf SSTab1.Tab = 5 Then
        If rs.State Then rs.Close
        str_sql = "select count(employee_code) as rec_count from t_special_leave " & _
                    "where employee_code = '" & txt_employee_code.Text & "' " & _
                        "and date(trans_date) = '" & Format(DTPicker_trans_leave.Value, "yyyy-MM-dd") & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            rs.Close
            Exit Function
        End If

        rs.Close
    End If
End Function

Private Sub check_invalid()
    If SSTab1.Tab = 0 Then
        MsgBox "Data Found...", vbCritical, headerMSG
        DTPicker_trans_marriage.Value = Now
        If DTPicker_trans_marriage.Enabled = True Then txt_nik.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        MsgBox "Data Found...", vbCritical, headerMSG
        DTPicker_trans_baby.Value = Now
        If DTPicker_trans_baby.Enabled = True Then txt_nik.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        MsgBox "Data Found...", vbCritical, headerMSG
        DTPicker_trans_condolence.Value = Now
        If DTPicker_trans_condolence.Enabled = True Then txt_nik.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        MsgBox "Data Found...", vbCritical, headerMSG
        DTPicker_trans_death.Value = Now
        If DTPicker_trans_death.Enabled = True Then txt_nik.SetFocus
    ElseIf SSTab1.Tab = 4 Then
        MsgBox "Data Found...", vbCritical, headerMSG
        DTPicker_trans_travel.Value = Now
        If DTPicker_trans_travel.Enabled = True Then txt_nik.SetFocus
    ElseIf SSTab1.Tab = 5 Then
        MsgBox "Data Found...", vbCritical, headerMSG
        DTPicker_trans_leave.Value = Now
        If DTPicker_trans_leave.Enabled = True Then txt_nik.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False

    If SSTab1.Tab = 0 Then
        If Not DTPicker_trans_marriage.Value = rsMarriage.Fields("trans_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not DTPicker_trans_baby.Value = rsBaby.Fields("trans_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Not DTPicker_trans_condolence.Value = rsCondolence.Fields("trans_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        If Not DTPicker_trans_death.Value = rsDeath.Fields("trans_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
        If Not DTPicker_trans_travel.Value = rsTravel.Fields("trans_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 5 Then
        If Not DTPicker_trans_leave.Value = rsLeave.Fields("trans_date").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True

    If SSTab1.Tab = 0 Then
        If Trim(txt_nik.Text) = "" Then
            MsgBox "NIK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If

        If Trim(txt_marriage_allowance.Text) = "" Then
            MsgBox "Marriage Allowance Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_marriage_allowance.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Trim(txt_nik.Text) = "" Then
            MsgBox "NIK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If

        If Trim(txt_baby_allowance.Text) = "" Then
            MsgBox "Deliver Baby Allowance Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_baby_allowance.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Trim(txt_nik.Text) = "" Then
            MsgBox "NIK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If

        If Trim(txt_condolence_allowance.Text) = "" Then
            MsgBox "Condolence Allowance Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_condolence_allowance.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        If Trim(txt_nik.Text) = "" Then
            MsgBox "NIK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If

        If Trim(txt_death_allowance.Text) = "" Then
            MsgBox "Death Allowance Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_death_allowance.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
        If Trim(txt_nik.Text) = "" Then
            MsgBox "NIK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If

        If Trim(txt_travel_allowance.Text) = "" Then
            MsgBox "Travel Allowance Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_travel_allowance.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 5 Then
        If Trim(txt_nik.Text) = "" Then
            MsgBox "NIK Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_nik.SetFocus
            check_validate_new = False
            Exit Function
        End If

        If Trim(txt_leave_allowance.Text) = "" Then
            MsgBox "Leave Allowance Is Empty...", vbOKOnly + vbInformation, headerMSG
            txt_leave_allowance.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If
End Function

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
    Dim i As Integer
    Dim vWhere As String
    Dim vCompany As String
    Dim vDivision As String
    Dim vFlagDivision As Integer
    Dim vEffectiveDate As String

    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Marriage.ApproxCount > 0 And TDBGrid_Marriage.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If

        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Marriage.Columns("employee_name").Value & "' - '" & TDBGrid_Marriage.Columns("trans_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub

        CnG.BeginTrans
            CnG.Execute "delete from t_special_marriage " & _
                "where employee_code = '" & TDBGrid_Marriage.Columns("employee_code").Value & "' " & _
                    "AND date(trans_date) = '" & Format(TDBGrid_Marriage.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Baby.ApproxCount > 0 And TDBGrid_Baby.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If

        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Baby.Columns("employee_name").Value & "' - '" & TDBGrid_Baby.Columns("trans_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub

        CnG.BeginTrans
            CnG.Execute "delete from t_special_baby " & _
                "where employee_code = '" & TDBGrid_Baby.Columns("employee_code").Value & "' " & _
                    "AND date(trans_date) = '" & Format(TDBGrid_Baby.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Condolence.ApproxCount > 0 And TDBGrid_Condolence.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If

        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Condolence.Columns("employee_name").Value & "' - '" & TDBGrid_Condolence.Columns("trans_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub

        CnG.BeginTrans
            CnG.Execute "delete from t_special_condolence " & _
                "where employee_code = '" & TDBGrid_Condolence.Columns("employee_code").Value & "' " & _
                    "AND date(trans_date) = '" & Format(TDBGrid_Condolence.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Death.ApproxCount > 0 And TDBGrid_Death.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If

        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Death.Columns("employee_name").Value & "' - '" & TDBGrid_Death.Columns("trans_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub

        CnG.BeginTrans
            CnG.Execute "delete from t_special_death " & _
                "where employee_code = '" & TDBGrid_Death.Columns("employee_code").Value & "' " & _
                    "AND date(trans_date) = '" & Format(TDBGrid_Death.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Travelling.ApproxCount > 0 And TDBGrid_Travelling.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If

        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Travelling.Columns("employee_name").Value & "' - '" & TDBGrid_Travelling.Columns("trans_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub

        CnG.BeginTrans
            CnG.Execute "delete from t_special_travelling " & _
                "where employee_code = '" & TDBGrid_Travelling.Columns("employee_code").Value & "' " & _
                    "AND date(trans_date) = '" & Format(TDBGrid_Travelling.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    ElseIf SSTab1.Tab = 5 Then
        If Not (TDBGrid_Leave.ApproxCount > 0 And TDBGrid_Leave.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            Exit Sub
        End If

        i = MsgBox("Are You Sure to Delete '" _
            & TDBGrid_Leave.Columns("employee_name").Value & "' - '" & TDBGrid_Leave.Columns("trans_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub

        CnG.BeginTrans
            CnG.Execute "delete from t_special_leave " & _
                "where employee_code = '" & TDBGrid_Leave.Columns("employee_code").Value & "' " & _
                    "AND date(trans_date) = '" & Format(TDBGrid_Leave.Columns("trans_date").Value, "yyyy-MM-dd") & "'"
        CnG.CommitTrans
    End If

    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
'On Error GoTo Err

Dim vFlagOTMethod As Integer
Dim vFlagCalcStart As Integer
Dim vFlagPPh As Integer

    vSetData = 1

    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Marriage.ApproxCount > 0 And TDBGrid_Marriage.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If

        With rsMarriage
            DTPicker_trans_marriage.Value = .Fields("trans_date").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name = .Fields("employee_name").Value

            txt_marriage_place.Text = .Fields("place").Value
            DTPicker_marriage = .Fields("date").Value
            txt_marriage_wife_husband.Text = .Fields("wife_husband_name").Value
            txt_marriage_allowance.Text = FormatNumber(.Fields("allowance").Value)
            txt_marriage_words.Text = UCase(TerbilangInggris(.Fields("allowance").Value))
            txt_marriage_description.Text = .Fields("description").Value

            vFlagPPh = .Fields("flag_pph21").Value
            opt_manual(0).Value = IIf(vFlagPPh = 0, 1, 0)
            opt_formula(0).Value = IIf(vFlagPPh = 0, 0, 1)

            If vFlagPPh = 0 Then
                txt_pph21_value(0).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
            Else
                txt_pph21_value(0).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
                TDBCombo_pph21(0).Text = .Fields("pph21_code").Value
                txt_pph21_name(0).Text = .Fields("pph21_name").Value
            End If
        End With
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Baby.ApproxCount > 0 And TDBGrid_Baby.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If

        With rsBaby
            DTPicker_trans_baby.Value = .Fields("trans_date").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name = .Fields("employee_name").Value
            
            txt_baby_name.Text = .Fields("child_name").Value
            txt_baby_place.Text = .Fields("birth_place").Value
            DTPicker_baby = .Fields("date_birth").Value
            txt_baby_allowance.Text = FormatNumber(.Fields("allowance").Value)
            txt_baby_words.Text = UCase(TerbilangInggris(.Fields("allowance").Value))
            txt_baby_description.Text = .Fields("description").Value

            vFlagPPh = .Fields("flag_pph21").Value
            opt_manual(1).Value = IIf(vFlagPPh = 0, 1, 0)
            opt_formula(1).Value = IIf(vFlagPPh = 0, 0, 1)

            If vFlagPPh = 0 Then
                txt_pph21_value(1).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
            Else
                txt_pph21_value(1).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
                TDBCombo_pph21(1).Text = .Fields("pph21_code").Value
                txt_pph21_name(1).Text = .Fields("pph21_name").Value
            End If
        End With
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Condolence.ApproxCount > 0 And TDBGrid_Condolence.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If

        With rsCondolence
            DTPicker_trans_condolence.Value = .Fields("trans_date").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name = .Fields("employee_name").Value

            txt_condolence_family.Text = .Fields("family_name").Value
            txt_condolence_place.Text = .Fields("death_place").Value
            DTPicker_Condolence = .Fields("death_date").Value
            txt_condolence_allowance.Text = FormatNumber(.Fields("allowance").Value)
            txt_condolence_words.Text = UCase(TerbilangInggris(.Fields("allowance").Value))
            txt_condolence_description.Text = .Fields("description").Value

            vFlagPPh = .Fields("flag_pph21").Value
            opt_manual(2).Value = IIf(vFlagPPh = 0, 1, 0)
            opt_formula(2).Value = IIf(vFlagPPh = 0, 0, 1)

            If vFlagPPh = 0 Then
                txt_pph21_value(2).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
            Else
                txt_pph21_value(2).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
                TDBCombo_pph21(2).Text = .Fields("pph21_code").Value
                txt_pph21_name(2).Text = .Fields("pph21_name").Value
            End If
        End With
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Death.ApproxCount > 0 And TDBGrid_Death.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If

        With rsDeath
            DTPicker_trans_death.Value = .Fields("trans_date").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name = .Fields("employee_name").Value

            txt_death_salary.Text = .Fields("basic_salary").Value
            DTPicker_Death = .Fields("date_pass_away").Value
            txt_death_age.Text = .Fields("masa_kerja").Value
            txt_death_allowance.Text = FormatNumber(.Fields("allowance").Value)
            txt_death_words.Text = UCase(TerbilangInggris(.Fields("allowance").Value))
            txt_death_description.Text = .Fields("description").Value

            vFlagPPh = .Fields("flag_pph21").Value
            opt_manual(3).Value = IIf(vFlagPPh = 0, 1, 0)
            opt_formula(3).Value = IIf(vFlagPPh = 0, 0, 1)

            If vFlagPPh = 0 Then
                txt_pph21_value(3).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
            Else
                txt_pph21_value(3).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
                TDBCombo_pph21(3).Text = .Fields("pph21_code").Value
                txt_pph21_name(3).Text = .Fields("pph21_name").Value
            End If
        End With
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Travelling.ApproxCount > 0 And TDBGrid_Travelling.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If

        With rsTravel
            DTPicker_trans_travel.Value = .Fields("trans_date").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name = .Fields("employee_name").Value

            txt_travel_long.Text = .Fields("lama_tugas").Value
            txt_travel_address.Text = .Fields("alamat_tujuan").Value
            txt_travel_country.Text = .Fields("negara_tujuan").Value
            txt_travel_requirement.Text = .Fields("keperluan").Value
            txt_travel_allowance.Text = FormatNumber(.Fields("allowance").Value)
            txt_travel_words.Text = UCase(TerbilangInggris(.Fields("allowance").Value))
            txt_travel_description.Text = .Fields("description").Value

            vFlagPPh = .Fields("flag_pph21").Value
            opt_manual(4).Value = IIf(vFlagPPh = 0, 1, 0)
            opt_formula(4).Value = IIf(vFlagPPh = 0, 0, 1)

            If vFlagPPh = 0 Then
                txt_pph21_value(4).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
            Else
                txt_pph21_value(4).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
                TDBCombo_pph21(4).Text = .Fields("pph21_code").Value
                txt_pph21_name(4).Text = .Fields("pph21_name").Value
            End If
        End With
    ElseIf SSTab1.Tab = 5 Then
        If Not (TDBGrid_Leave.ApproxCount > 0 And TDBGrid_Leave.Bookmark > 0) Then
            MsgBox "There Is No Data Selected...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If

        With rsLeave
            DTPicker_trans_leave.Value = .Fields("trans_date").Value
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name = .Fields("employee_name").Value

            txt_leave_salary.Text = .Fields("basic_salary").Value
            txt_leave_pengali.Text = .Fields("pengali").Value
            txt_leave_allowance.Text = .Fields("allowance").Value
            txt_leave_allowance.Text = FormatNumber(.Fields("allowance").Value)
            txt_leave_words.Text = UCase(TerbilangInggris(.Fields("allowance").Value))
            txt_leave_description.Text = .Fields("description").Value

            vFlagPPh = .Fields("flag_pph21").Value
            opt_manual(5).Value = IIf(vFlagPPh = 0, 1, 0)
            opt_formula(5).Value = IIf(vFlagPPh = 0, 0, 1)

            If vFlagPPh = 0 Then
                txt_pph21_value(5).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
            Else
                txt_pph21_value(5).Text = FormatNumber(IIf(IsNull(.Fields("pph21_value").Value), 0, .Fields("pph21_value").Value))
                TDBCombo_pph21(5).Text = .Fields("pph21_code").Value
                txt_pph21_name(5).Text = .Fields("pph21_name").Value
            End If
        End With
    End If

    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub new_data()
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()
On Error GoTo Err

Dim vNo As Integer
Dim vSisa As Double

    CnG.BeginTrans

    If SSTab1.Tab = 0 Then
        vSisa = DropAllComma(txt_marriage_allowance.Text) - DropAllComma(txt_pph21_value(0).Text)
        SQL = "INSERT INTO t_special_marriage(employee_code,trans_date,place,date," & _
                "wife_husband_name,allowance,description,flag_pph21,pph21_value," & _
                "pph21_code,sisa_allowance,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "','" & Format(DTPicker_trans_marriage.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_marriage_place & "','" & Format(DTPicker_marriage, "yyyy-MM-dd") & "','" & txt_marriage_wife_husband.Text & "'," & _
                "'" & DropAllComma(txt_marriage_allowance.Text) & "','" & txt_marriage_description.Text & "','" & IIf(opt_manual(0).Value, 0, 1) & "'," & _
                "'" & DropAllComma(txt_pph21_value(0).Text) & "','" & TDBCombo_pph21(0).Text & "','" & vSisa & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        vSisa = DropAllComma(txt_baby_allowance.Text) - DropAllComma(txt_pph21_value(1).Text)
        SQL = "INSERT INTO t_special_baby(employee_code,trans_date,child_name,birth_place," & _
                "date_birth,allowance,description,flag_pph21,pph21_value,pph21_code,sisa_allowance," & _
                "entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "','" & Format(DTPicker_trans_baby.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_baby_name.Text & "','" & txt_baby_name.Text & "','" & Format(DTPicker_baby, "yyyy-MM-dd") & "'," & _
                "'" & DropAllComma(txt_baby_allowance.Text) & "','" & txt_baby_description.Text & "','" & IIf(opt_manual(1).Value, 0, 1) & "'," & _
                "'" & DropAllComma(txt_pph21_value(1).Text) & "','" & TDBCombo_pph21(1).Text & "','" & vSisa & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        vSisa = DropAllComma(txt_condolence_allowance.Text) - DropAllComma(txt_pph21_value(2).Text)
        SQL = "INSERT INTO t_special_condolence(employee_code,trans_date,family_name,death_place," & _
                "death_date,allowance,description,flag_pph21,pph21_value,pph21_code,sisa_allowance," & _
                "entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "','" & Format(DTPicker_trans_condolence.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_condolence_family.Text & "','" & txt_condolence_place.Text & "','" & Format(DTPicker_Condolence, "yyyy-MM-dd") & "'," & _
                "'" & DropAllComma(txt_condolence_allowance.Text) & "','" & txt_condolence_description.Text & "','" & IIf(opt_manual(2).Value, 0, 1) & "'," & _
                "'" & DropAllComma(txt_pph21_value(2).Text) & "','" & TDBCombo_pph21(2).Text & "','" & vSisa & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        vSisa = DropAllComma(txt_death_allowance.Text) - DropAllComma(txt_pph21_value(3).Text)
        SQL = "INSERT INTO t_special_death(employee_code,trans_date,date_pass_away,masa_kerja," & _
                "basic_salary,allowance,description,flag_pph21,pph21_value,pph21_code,sisa_allowance," & _
                "entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "','" & Format(DTPicker_trans_death.Value, "yyyy-MM-dd") & "'," & _
                "'" & Format(DTPicker_Death, "yyyy-MM-dd") & "','" & txt_death_age.Text & "','" & DropAllComma(txt_death_salary.Text) & "'," & _
                "'" & DropAllComma(txt_death_allowance.Text) & "','" & txt_death_description.Text & "','" & IIf(opt_manual(3).Value, 0, 1) & "'," & _
                "'" & DropAllComma(txt_pph21_value(3).Text) & "','" & TDBCombo_pph21(3).Text & "','" & vSisa & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        vSisa = DropAllComma(txt_travel_allowance.Text) - DropAllComma(txt_pph21_value(4).Text)
        SQL = "INSERT INTO t_special_travelling(employee_code,trans_date,lama_tugas,alamat_tujuan," & _
                "negara_tujuan,keperluan,allowance,description,flag_pph21,pph21_value,pph21_code,sisa_allowance," & _
                "entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "','" & Format(DTPicker_trans_travel.Value, "yyyy-MM-dd") & "'," & _
                "'" & txt_travel_long.Text & "','" & txt_travel_address.Text & "','" & txt_travel_country & "'," & _
                "'" & txt_travel_requirement.Text & "','" & DropAllComma(txt_travel_allowance.Text) & "'," & _
                "'" & txt_travel_description.Text & "','" & IIf(opt_manual(4).Value, 0, 1) & "'," & _
                "'" & DropAllComma(txt_pph21_value(4).Text) & "','" & TDBCombo_pph21(4).Text & "','" & vSisa & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        vSisa = DropAllComma(txt_leave_allowance.Text) - DropAllComma(txt_pph21_value(5).Text)
        SQL = "INSERT INTO t_special_leave(employee_code,trans_date,basic_salary,pengali," & _
                "allowance,description,flag_pph21,pph21_value,pph21_code,sisa_allowance," & _
                "entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "','" & Format(DTPicker_trans_leave.Value, "yyyy-MM-dd") & "'," & _
                "'" & DropAllComma(txt_leave_salary.Text) & "','" & DropAllComma(txt_leave_pengali.Text) & "'," & _
                "'" & DropAllComma(txt_leave_allowance.Text) & "','" & txt_leave_description.Text & "'," & _
                "'" & IIf(opt_manual(5).Value, 0, 1) & "'," & _
                "'" & DropAllComma(txt_pph21_value(5).Text) & "','" & TDBCombo_pph21(5).Text & "','" & vSisa & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
'On Error GoTo Err
    CnG.BeginTrans

    If SSTab1.Tab = 0 Then
        SQL = "DELETE FROM t_special_marriage WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(trans_date) = '" & Format(DTPicker_trans_marriage.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "DELETE FROM t_special_baby WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(trans_date) = '" & Format(DTPicker_trans_baby.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "DELETE FROM t_special_condolence WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(trans_date) = '" & Format(DTPicker_trans_condolence.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "DELETE FROM t_special_death WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(trans_date) = '" & Format(DTPicker_trans_death.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "DELETE FROM t_special_travelling WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(trans_date) = '" & Format(DTPicker_trans_travel.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        SQL = "DELETE FROM t_special_leave WHERE employee_code = '" & txt_employee_code.Text & "' " & _
                "AND DATE(trans_date) = '" & Format(DTPicker_trans_leave.Value, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    End If
    CnG.CommitTrans
       
    Call insert_new_data
    
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub simpan_data()
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
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean, ByVal Index As Integer)
    cmdNew(Index).Enabled = a And blnUser_Add
    cmdSave(Index).Enabled = b
    cmdEdit(Index).Enabled = c And blnUser_Edit
    cmdDelete(Index).Enabled = d And blnUser_Delete
    cmdCancel(Index).Enabled = e
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            If Not LCase(Ctr.name) = "dtpicker_special_from" And _
                Not LCase(Ctr.name) = "dtpicker_special_to" Then Ctr.Value = Now
        ElseIf TypeOf Ctr Is CheckBox Then
            Ctr.Value = 0
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
    txt_nik.Visible = True
    txt_nik.Enabled = True
    txt_employee_name.Visible = True
    CmdBrowse.Visible = True
    
    opt_manual(SSTab1.Tab).Value = True
    txt_pph21_value(SSTab1.Tab).Text = FormatNumber("0")
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_marriage.Visible = True
            DTPicker_trans_marriage.Enabled = True
            TDBGrid_Marriage.Enabled = False
            Call set_new_data

'            If DTPicker_trans_marriage.Enabled = True Then
'                txt_nik.SetFocus
'            End If
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_baby.Visible = True
            DTPicker_trans_baby.Enabled = True
            TDBGrid_Baby.Enabled = False
            Call set_new_data

'            If DTPicker_trans_baby.Enabled = True Then
'                txt_nik.SetFocus
'            End If
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_condolence.Visible = True
            DTPicker_trans_condolence.Enabled = True
            TDBGrid_Condolence.Enabled = False
            Call set_new_data

'            If DTPicker_trans_condolence.Enabled = True Then
'                txt_nik.SetFocus
'            End If
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_death.Visible = True
            DTPicker_trans_death.Enabled = True
            TDBGrid_Death.Enabled = False
            Call set_new_data

'            If DTPicker_trans_death.Enabled = True Then
'                txt_nik.SetFocus
'            End If
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_travelling.Visible = True
            DTPicker_trans_travel.Enabled = True
            TDBGrid_Travelling.Enabled = False
            Call set_new_data

'            If DTPicker_trans_travel.Enabled = True Then
'                txt_nik.SetFocus
'            End If
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_leave.Visible = True
            DTPicker_trans_leave.Enabled = True
            TDBGrid_Leave.Enabled = False
            Call set_new_data

'            If DTPicker_trans_leave.Enabled = True Then
'                txt_nik.SetFocus
'            End If
        End If

    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        txt_nik.Visible = False
        txt_employee_name.Visible = False
        CmdBrowse.Visible = False
    
        If SSTab1.Tab = 0 Then
            fra_entry_marriage.Visible = False
            TDBGrid_Marriage.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_baby.Visible = False
            TDBGrid_Baby.Enabled = True
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_condolence.Visible = False
            TDBGrid_Condolence.Enabled = True
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_death.Visible = False
            TDBGrid_Death.Enabled = True
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_travelling.Visible = False
            TDBGrid_Travelling.Enabled = True
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_leave.Visible = False
            TDBGrid_Leave.Enabled = True
        End If

    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data

        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        txt_nik.Visible = True
        txt_employee_name.Visible = True
        CmdBrowse.Visible = True
    
        If SSTab1.Tab = 0 Then
            DTPicker_trans_marriage.Enabled = False
            fra_entry_marriage.Visible = True
            TDBGrid_Marriage.Enabled = False
            txt_nik.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            DTPicker_trans_baby.Enabled = False
            fra_entry_baby.Visible = True
            TDBGrid_Baby.Enabled = False
            txt_nik.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            DTPicker_trans_condolence.Enabled = False
            fra_entry_condolence.Visible = True
            TDBGrid_Condolence.Enabled = False
            txt_nik.Enabled = False
        ElseIf SSTab1.Tab = 3 Then
            DTPicker_trans_death.Enabled = False
            fra_entry_death.Visible = True
            TDBGrid_Death.Enabled = False
            txt_nik.Enabled = False
        ElseIf SSTab1.Tab = 4 Then
            DTPicker_trans_travel.Enabled = False
            fra_entry_travelling.Visible = True
            TDBGrid_Travelling.Enabled = False
            txt_nik.Enabled = False
        ElseIf SSTab1.Tab = 5 Then
            DTPicker_trans_leave.Enabled = False
            fra_entry_leave.Visible = True
            TDBGrid_Leave.Enabled = False
            txt_nik.Enabled = False
        End If
    End If
End Sub

Private Sub load_mode()
    If int_mode = 1 Then        ' new
        Call set_buttons_enable(False, True, False, False, True, False, False, SSTab1.Tab)
    ElseIf int_mode = 0 Then    ' view
        Call set_buttons_enable(True, False, True, True, False, True, True, SSTab1.Tab)
    ElseIf int_mode = 2 Then    ' edit/revise
        Call set_buttons_enable(False, True, False, False, True, False, False, SSTab1.Tab)
    End If

    Call set_data_mode
End Sub

Private Sub cmdSearch_Click()
    Call load_data
End Sub

Private Sub DTPicker_Death_Change()
    txt_death_age = Trim(str(year(DTPicker_Death.Value) - year(vStartWorking)))
End Sub

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_special_from, DTPicker_special_to)
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    oClause = ""
    opt_manual(SSTab1.Tab).Value = True

'    Call load_data
    Call load_data_company
    Call createGridKar
    
    Timer1.Enabled = True
    
    txt_nik.Visible = False
    txt_employee_name.Visible = False
    CmdBrowse.Visible = False
    
    DTPicker_Periode.Value = Now
    DTPicker_special_from.Value = Now
    DTPicker_special_to.Value = Now
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Marriage.Columns
            Col.FilterText = ""
        Next Col
        rsMarriage.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_Baby.Columns
            Col.FilterText = ""
        Next Col
        rsBaby.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_Condolence.Columns
            Col.FilterText = ""
        Next Col
        rsCondolence.Filter = adFilterNone
    ElseIf SSTab1.Tab = 3 Then
        For Each Col In TDBGrid_Death.Columns
            Col.FilterText = ""
        Next Col
        rsDeath.Filter = adFilterNone
    ElseIf SSTab1.Tab = 4 Then
        For Each Col In TDBGrid_Travelling.Columns
            Col.FilterText = ""
        Next Col
        rsTravel.Filter = adFilterNone
    ElseIf SSTab1.Tab = 5 Then
        For Each Col In TDBGrid_Leave.Columns
            Col.FilterText = ""
        Next Col
        rsLeave.Filter = adFilterNone
    End If
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
    Set frm_trans_special_allowance = Nothing
End Sub

Private Sub filter_change()
On Error GoTo Err

    Dim i As Integer

    If SSTab1.Tab = 0 Then
        Set Cols = TDBGrid_Marriage.Columns
        i = TDBGrid_Marriage.Col
        TDBGrid_Marriage.HoldFields

        rsMarriage.Filter = getFilter()
        TDBGrid_Marriage.Col = i
        TDBGrid_Marriage.EditActive = True

        TDBGrid_Marriage.SelStart = Len(TDBGrid_Marriage.Columns(i).FilterText)
        If TDBGrid_Marriage.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Marriage.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_Baby.Columns
        i = TDBGrid_Baby.Col
        TDBGrid_Baby.HoldFields

        rsBaby.Filter = getFilter()
        TDBGrid_Baby.Col = i
        TDBGrid_Baby.EditActive = True

        TDBGrid_Baby.SelStart = Len(TDBGrid_Baby.Columns(i).FilterText)
        If TDBGrid_Baby.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Baby.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_Condolence.Columns
        i = TDBGrid_Condolence.Col
        TDBGrid_Condolence.HoldFields

        rsCondolence.Filter = getFilter()
        TDBGrid_Condolence.Col = i
        TDBGrid_Condolence.EditActive = True

        TDBGrid_Condolence.SelStart = Len(TDBGrid_Condolence.Columns(i).FilterText)
        If TDBGrid_Condolence.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Condolence.Col = i
        End If
    ElseIf SSTab1.Tab = 3 Then
        Set Cols = TDBGrid_Death.Columns
        i = TDBGrid_Death.Col
        TDBGrid_Death.HoldFields

        rsDeath.Filter = getFilter()
        TDBGrid_Death.Col = i
        TDBGrid_Death.EditActive = True

        TDBGrid_Death.SelStart = Len(TDBGrid_Death.Columns(i).FilterText)
        If TDBGrid_Death.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Death.Col = i
        End If
    ElseIf SSTab1.Tab = 4 Then
        Set Cols = TDBGrid_Travelling.Columns
        i = TDBGrid_Travelling.Col
        TDBGrid_Travelling.HoldFields

        rsTravel.Filter = getFilter()
        TDBGrid_Travelling.Col = i
        TDBGrid_Travelling.EditActive = True

        TDBGrid_Travelling.SelStart = Len(TDBGrid_Travelling.Columns(i).FilterText)
        If TDBGrid_Travelling.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Travelling.Col = i
        End If
    ElseIf SSTab1.Tab = 5 Then
        Set Cols = TDBGrid_Leave.Columns
        i = TDBGrid_Leave.Col
        TDBGrid_Leave.HoldFields

        rsLeave.Filter = getFilter()
        TDBGrid_Leave.Col = i
        TDBGrid_Leave.EditActive = True

        TDBGrid_Leave.SelStart = Len(TDBGrid_Leave.Columns(i).FilterText)
        If TDBGrid_Leave.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Leave.Col = i
        End If
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsMarriage.State Then rsMarriage.Close
        SQL = "SELECT a.*,b.nik, b.employee_name, c.pph21_name " & _
                "FROM t_special_marriage a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "LEFT JOIN m_pph21 c ON a.pph21_code = c.pph21_code " & _
                "WHERE DATE(trans_date) BETWEEN '" & Format(DTPicker_special_from.Value, "yyyy-MM-dd") & "' " & _
                    "AND '" & Format(DTPicker_special_to.Value, "yyyy-MM-dd") & "' " & oClause
        rsMarriage.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        TDBGrid_Marriage.DataSource = rsMarriage
    ElseIf SSTab1.Tab = 1 Then
        If rsBaby.State Then rsBaby.Close
        SQL = "SELECT a.*,b.nik, b.employee_name, c.pph21_name " & _
                "FROM t_special_baby a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "LEFT JOIN m_pph21 c ON a.pph21_code = c.pph21_code " & _
                "WHERE DATE(trans_date) BETWEEN '" & Format(DTPicker_special_from.Value, "yyyy-MM-dd") & "' " & _
                    "AND '" & Format(DTPicker_special_to.Value, "yyyy-MM-dd") & "' " & oClause
        rsBaby.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        TDBGrid_Baby.DataSource = rsBaby
    ElseIf SSTab1.Tab = 2 Then
        If rsCondolence.State Then rsCondolence.Close
        SQL = "SELECT a.*,b.nik, b.employee_name, c.pph21_name " & _
                "FROM t_special_condolence a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "LEFT JOIN m_pph21 c ON a.pph21_code = c.pph21_code " & _
                "WHERE DATE(trans_date) BETWEEN '" & Format(DTPicker_special_from.Value, "yyyy-MM-dd") & "' " & _
                    "AND '" & Format(DTPicker_special_to.Value, "yyyy-MM-dd") & "' " & oClause
        rsCondolence.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        TDBGrid_Condolence.DataSource = rsCondolence
    ElseIf SSTab1.Tab = 3 Then
        If rsDeath.State Then rsDeath.Close
        SQL = "SELECT a.*,b.nik, b.employee_name, c.pph21_name " & _
                "FROM t_special_death a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "LEFT JOIN m_pph21 c ON a.pph21_code = c.pph21_code " & _
                "WHERE DATE(trans_date) BETWEEN '" & Format(DTPicker_special_from.Value, "yyyy-MM-dd") & "' " & _
                    "AND '" & Format(DTPicker_special_to.Value, "yyyy-MM-dd") & "' " & oClause
        rsDeath.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        TDBGrid_Death.DataSource = rsDeath
    ElseIf SSTab1.Tab = 4 Then
        If rsTravel.State Then rsTravel.Close
        SQL = "SELECT a.*,b.nik, b.employee_name, c.pph21_name " & _
                "FROM t_special_travelling a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "LEFT JOIN m_pph21 c ON a.pph21_code = c.pph21_code " & _
                "WHERE DATE(trans_date) BETWEEN '" & Format(DTPicker_special_from.Value, "yyyy-MM-dd") & "' " & _
                    "AND '" & Format(DTPicker_special_to.Value, "yyyy-MM-dd") & "' " & oClause
        rsTravel.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        TDBGrid_Travelling.DataSource = rsTravel
    ElseIf SSTab1.Tab = 5 Then
        If rsLeave.State Then rsLeave.Close
        SQL = "SELECT a.*,b.nik, b.employee_name, c.pph21_name " & _
                "FROM t_special_leave a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "LEFT JOIN m_pph21 c ON a.pph21_code = c.pph21_code " & _
                "WHERE DATE(trans_date) BETWEEN '" & Format(DTPicker_special_from.Value, "yyyy-MM-dd") & "' " & _
                    "AND '" & Format(DTPicker_special_to.Value, "yyyy-MM-dd") & "' " & oClause
        rsLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        TDBGrid_Leave.DataSource = rsLeave
    End If
End Sub

Private Sub opt_manual_Click(Index As Integer)
    If opt_manual(Index).Value Then
        txt_pph21_value(Index).Visible = True
        txt_pph21_value(Index).Text = FormatNumber("0")
        
        TDBCombo_pph21(Index).Visible = False
        txt_pph21_name(Index).Visible = False
    Else
        txt_pph21_value(Index).Visible = False
        
        Call load_data_pph21(Index)
        
        TDBCombo_pph21(Index).Visible = True
        txt_pph21_name(Index).Visible = True
    End If
End Sub

Private Sub opt_formula_Click(Index As Integer)
    If opt_formula(Index).Value Then
        txt_pph21_value(Index).Visible = False
        
        Call load_data_pph21(Index)
        
        TDBCombo_pph21(Index).Visible = True
        txt_pph21_name(Index).Visible = True
    Else
        txt_pph21_value(Index).Visible = True
        txt_pph21_value(Index).Text = FormatNumber("0")
        
        TDBCombo_pph21(Index).Visible = False
        txt_pph21_name(Index).Visible = False
    End If
End Sub

Private Sub load_data_pph21(Index As Integer)
    If rsPPh.State Then rsPPh.Close
    SQL = "select * from m_pph21 order by pph21_code"
    rsPPh.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_pph21(Index).RowSource = rsPPh
End Sub

Private Sub TDBCombo_pph21_ItemChange(Index As Integer)
    If TDBCombo_pph21(Index).ApproxCount > 0 Then
        TDBCombo_pph21(Index).Text = TDBCombo_pph21(Index).Columns("pph21_code").Value
        txt_pph21_name(Index).Text = TDBCombo_pph21(Index).Columns("pph21_name").Value
        
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    oClause = ""
    opt_manual(SSTab1.Tab).Value = True
'    Call load_data
    
    If SSTab1.Tab = 0 Then
        vTransDate = Format(DTPicker_trans_marriage.Value, "yyyy-MM-dd")
    ElseIf SSTab1.Tab = 1 Then
        vTransDate = Format(DTPicker_trans_baby.Value, "yyyy-MM-dd")
    ElseIf SSTab1.Tab = 2 Then
        vTransDate = Format(DTPicker_trans_condolence.Value, "yyyy-MM-dd")
    ElseIf SSTab1.Tab = 3 Then
        vTransDate = Format(DTPicker_trans_death.Value, "yyyy-MM-dd")
    ElseIf SSTab1.Tab = 4 Then
        vTransDate = Format(DTPicker_trans_travel.Value, "yyyy-MM-dd")
    ElseIf SSTab1.Tab = 5 Then
        vTransDate = Format(DTPicker_trans_leave.Value, "yyyy-MM-dd")
    End If
    
    txt_nik.Text = ""
    txt_employee_code.Text = ""
    txt_employee_name.Text = ""
    
    vTitleCode = ""
    If SSTab1.Tab = 3 Then
        txt_death_salary = ""
    ElseIf SSTab1.Tab = 5 Then
        txt_leave_salary = ""
    Else
        vBasicSalary = ""
    End If
    
    TDBCombo_pph21(SSTab1.Tab).Visible = False
    txt_pph21_name(SSTab1.Tab).Visible = False
    Call opt_manual_Click(SSTab1.Tab)
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
    End If
End Sub

Private Sub TDBGrid_Marriage_HeadClick(ByVal ColIndex As Integer)
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Marriage.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Marriage.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Marriage.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Marriage.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_baby_HeadClick(ByVal ColIndex As Integer)
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Baby.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Baby.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Baby.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Baby.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_condolence_HeadClick(ByVal ColIndex As Integer)
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Condolence.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Condolence.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Condolence.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Condolence.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_death_HeadClick(ByVal ColIndex As Integer)
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Death.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Death.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Death.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Death.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_travelling_HeadClick(ByVal ColIndex As Integer)
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Travelling.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Travelling.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Travelling.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Travelling.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub TDBGrid_Leave_HeadClick(ByVal ColIndex As Integer)
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid_Leave.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid_Leave.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid_Leave.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid_Leave.Columns(ColIndex).DataField
    Call load_data

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub txt_leave_pengali_Validate(Cancel As Boolean)
    txt_leave_allowance.Text = DropAllComma(txt_leave_salary) * txt_leave_pengali.Text
End Sub

Private Sub txt_marriage_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_marriage_allowance.Text) = "" Then
        txt_marriage_allowance.Text = FormatNumber(DropAllComma(txt_marriage_allowance.Text))
        txt_marriage_words.Text = UCase(TerbilangInggris(DropAllComma(txt_marriage_allowance.Text)))
    End If
End Sub

Private Sub txt_baby_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_baby_allowance.Text) = "" Then
        txt_baby_allowance.Text = FormatNumber(DropAllComma(txt_baby_allowance.Text))
        txt_baby_words.Text = UCase(TerbilangInggris(DropAllComma(txt_baby_allowance.Text)))
    End If
End Sub

Private Sub txt_condolence_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_condolence_allowance.Text) = "" Then
        txt_condolence_allowance.Text = FormatNumber(DropAllComma(txt_condolence_allowance.Text))
        txt_condolence_words.Text = UCase(TerbilangInggris(DropAllComma(txt_condolence_allowance.Text)))
    End If
End Sub

Private Sub txt_death_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_death_allowance.Text) = "" Then
        txt_death_allowance.Text = FormatNumber(DropAllComma(txt_death_allowance.Text))
        txt_death_words.Text = UCase(TerbilangInggris(DropAllComma(txt_death_allowance.Text)))
    End If
End Sub

Private Sub txt_travel_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_travel_allowance.Text) = "" Then
        txt_travel_allowance.Text = FormatNumber(DropAllComma(txt_travel_allowance.Text))
        txt_travel_words.Text = UCase(TerbilangInggris(DropAllComma(txt_travel_allowance.Text)))
    End If
End Sub

Private Sub txt_leave_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_leave_allowance.Text) = "" Then
        txt_leave_allowance.Text = FormatNumber(DropAllComma(txt_leave_allowance.Text))
        txt_leave_words.Text = UCase(TerbilangInggris(DropAllComma(txt_leave_allowance.Text)))
    End If
End Sub

Private Sub txt_pph21_value_Validate(Index As Integer, Cancel As Boolean)
    If Not Trim(txt_pph21_value(Index).Text) = "" Then
        txt_pph21_value(Index).Text = FormatNumber(DropAllComma(txt_pph21_value(Index).Text))
    End If
End Sub

Private Sub tdbGrid_Marriage_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_Baby_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_Condolence_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_Death_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_Travelling_FilterChange()
    Call filter_change
End Sub

Private Sub tdbGrid_Leave_FilterChange()
    Call filter_change
End Sub


Private Sub CmdNew_Click(Index As Integer)
    Call new_data
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Call edit_data
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Call simpan_data
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    Call delete_data
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Call cancel_data
End Sub


Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "ID EMP. CODE", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMPLOYEE NAME", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .AddColumn "title_code", 2000, , , , , , , , False
      .AddColumn "start_working", 2000, , , "yyyy-MM-dd", , , , , False
      .AddColumn "basic_salary", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
Dim vParam As String

    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code,title_code,start_working, " & _
                       "(select basic_salary from m_salary_standard WHERE date(salary_date) <= '" & vTransDate & "' order by salary_date desc limit 1) basic_salary " & _
                    "from m_employee " & _
                    "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                       "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                           "OR employee_name LIKE '%" & txt_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code,title_code,start_working, " & _
                       "(select basic_salary from m_salary_standard WHERE date(salary_date) <= '" & vTransDate & "' order by salary_date desc limit 1) basic_salary " & _
                    "from m_employee " & _
                    "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                       "AND " & vParam & " " & _
                       "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                           "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                       "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & _
                                    rs!employee_code & vbTab & rs!title_code & vbTab & _
                                    rs!start_working & vbTab & rs!basic_salary
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
                txt_nik.Text = rs!nik
                
                vTitleCode = rs!title_code
                vStartWorking = rs!start_working
                If SSTab1.Tab = 3 Then
                    txt_death_salary = FormatNumber(rs!basic_salary)
                    txt_death_age = FormatNumber((DateDiff("m", vStartWorking, DTPicker_Death.Value) / 12))
                ElseIf SSTab1.Tab = 5 Then
                    txt_leave_salary = FormatNumber(rs!basic_salary)
                Else
                    vBasicSalary = rs!basic_salary
                End If
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
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            
            vTitleCode = LynxGrid2.CellText(LynxGrid2.Row, 3)
            vStartWorking = LynxGrid2.CellText(LynxGrid2.Row, 4)
            If SSTab1.Tab = 3 Then
                txt_death_salary = FormatNumber(LynxGrid2.CellText(LynxGrid2.Row, 5))
                txt_death_age = FormatNumber((DateDiff("m", vStartWorking, DTPicker_Death.Value) / 12))
            ElseIf SSTab1.Tab = 5 Then
                txt_leave_salary = FormatNumber(LynxGrid2.CellText(LynxGrid2.Row, 5))
            Else
                vBasicSalary = LynxGrid2.CellText(LynxGrid2.Row, 5)
            End If
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
        
        vTitleCode = ""
        vStartWorking = ""
        If SSTab1.Tab = 3 Then
            txt_death_salary = ""
        ElseIf SSTab1.Tab = 5 Then
            txt_leave_salary = ""
        Else
            vBasicSalary = ""
        End If
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

'Private Sub hitung_pph(vPPhCode As String)
'Dim vPPh_5, vPPh_15, vPPh_25, vPPh_30 As Double
'Dim vPPh_Total As Double
'
'    SQL = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 1 AND pph21_code = '" & vPPhCode & "'"
'    rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rspph.RecordCount > 0 Then
'        If txt_jml_pesangon.Value > rspph!pph21_upper Then '50000000
'            v_pph_5 = (rspph!pph21_percentage / 100) * rspph!pph21_upper '50000000
'        Else
'            v_pph_5 = (rspph!pph21_percentage / 100) * txt_jml_pesangon.Value
'        End If
'    End If
'    rspph.Close
'
'    SQL = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 2 AND pph21_code = '" & vPPhCode & "'"
'    rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rspph.RecordCount > 0 Then
'        If txt_jml_pesangon.Value <= rspph!pph21_under Then '50000000
'            v_pph_15 = 0
'        ElseIf v_tot_income < rspph!pph21_upper Then '100000000
'            v_pph_15 = (rspph!pph21_percentage / 100) * (txt_jml_pesangon.Value - rspph!pph21_under) '50000000
'        Else
''                v_pph_5 = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '50000000
'            v_pph_15 = (rspph!pph21_percentage / 100) * (txt_jml_pesangon.Value - rspph!pph21_under) '50000000
'        End If
'    End If
'    rspph.Close
'
'    SQL = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 3 AND pph21_code = '" & vPPhCode & "'"
'    rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rspph.RecordCount > 0 Then
'        If txt_jml_pesangon.Value <= rspph!pph21_under Then '500000000
'            v_pph_25 = 0
'        ElseIf txt_jml_pesangon.Value < rspph!pph21_upper Then  '500000000
'            v_pph_25 = (rspph!pph21_percentage / 100) * (txt_jml_pesangon.Value - rspph!pph21_under) '100000000
'        Else
'            v_pph_25 = (rspph!pph21_percentage / 100) * (rspph!pph21_upper - rspph!pph21_under) '100000000
'        End If
'    End If
'    rspph.Close
'
'    SQL = "SELECT pph21_under, pph21_upper, pph21_percentage FROM m_pph21_detail WHERE pph21_number = 4 AND pph21_code = 'CPST'"
'    rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    If rspph.RecordCount > 0 Then
'        If txt_jml_pesangon.Value <= rspph!pph21_under Then '500000000
'            v_pph_30 = 0
'        Else
'            If txt_jml_pesangon.Value > rspph!pph21_under Then '500000000
'                v_pph_30 = (rspph!pph21_percentage / 100) * (v_tot_income - rspph!pph21_under) '500000000
'            Else
'                v_pph_30 = 0
'            End If
'        End If
'    End If
'    rspph.Close
'End Sub
