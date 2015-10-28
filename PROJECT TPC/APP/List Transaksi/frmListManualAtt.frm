VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_list_manual_att 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANUAL ATTENDANCE"
   ClientHeight    =   11385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "frmListManualAtt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11385
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   14100
      TabIndex        =   25
      Top             =   10620
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
      MICON           =   "frmListManualAtt.frx":058A
      PICN            =   "frmListManualAtt.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.LynxGrid LynxGrid3 
      Height          =   2925
      Left            =   6060
      TabIndex        =   32
      Top             =   1650
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9885
      Left            =   180
      TabIndex        =   0
      Top             =   690
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   17436
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "INPUT MANUAL"
      TabPicture(0)   =   "frmListManualAtt.frx":1638
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label15"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label17"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label18"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label19"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label20"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label21"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label22"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label23"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label24"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label25"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label26"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label27"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label28"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label29"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label30"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label31"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label32"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label33"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label34"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label36"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label35"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label37"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label38"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl_l"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lbl_late"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl_sick"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl_pl"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl_n"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl_a"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lbl_trans"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl_meal"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "lbl_hl_40"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "lbl_hl_30"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lbl_hl_20"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lbl_hk_20"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lbl_hk_15"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmd_reproccess"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdNext"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmdRefresh"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TDBCombo_company"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TDBCombo_Group_Shift"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "DTPicker1"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt_company_name"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt_group_shift"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Frame1"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "cmdPrev"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "timer1"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Frame4"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdNew"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cmdEdit"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cmdDelete"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "ProgressBar3"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).ControlCount=   60
      TabCaption(1)   =   "IMPORT DATA ATTENDANCE"
      TabPicture(1)   =   "frmListManualAtt.frx":1654
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "cmdBrowse"
      Tab(1).Control(2)=   "cmdImport"
      Tab(1).Control(3)=   "ProgressBar1"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "IMPORT ATTENDANCE FROM DATA  MACHINE"
      TabPicture(2)   =   "frmListManualAtt.frx":1670
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "cmdBrowseM"
      Tab(2).Control(2)=   "cmdImportM"
      Tab(2).Control(3)=   "ProgressBar2"
      Tab(2).Control(4)=   "Frame3"
      Tab(2).ControlCount=   5
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   135
         Left            =   10020
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prj_tpc.vbButton cmdDelete 
         Height          =   465
         Left            =   12720
         TabIndex        =   12
         Top             =   1380
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":168C
         PICN            =   "frmListManualAtt.frx":16A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdEdit 
         Height          =   465
         Left            =   11370
         TabIndex        =   13
         Top             =   1380
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":273A
         PICN            =   "frmListManualAtt.frx":2756
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   465
         Left            =   10020
         TabIndex        =   14
         Top             =   1380
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":37E8
         PICN            =   "frmListManualAtt.frx":3804
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame4 
         Caption         =   "Advanced Filter"
         Height          =   1485
         Left            =   4920
         TabIndex        =   33
         Top             =   390
         Width           =   5025
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   36
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DragMode        =   1  'Automatic
            Height          =   285
            Left            =   2220
            TabIndex        =   35
            Top             =   270
            Width           =   2715
         End
         Begin prj_tpc.vbButton cmdBrowse_Emp 
            Height          =   285
            Left            =   1860
            TabIndex        =   34
            Top             =   270
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
            MICON           =   "frmListManualAtt.frx":4896
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPicker_from 
            Height          =   315
            Left            =   960
            TabIndex        =   39
            Top             =   960
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   94699523
            CurrentDate     =   40794
         End
         Begin MSComCtl2.DTPicker DTPicker_to 
            Height          =   315
            Left            =   2460
            TabIndex        =   40
            Top             =   960
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   94699523
            CurrentDate     =   40794
         End
         Begin prj_tpc.vbButton cmdSearch 
            Height          =   795
            Left            =   3630
            TabIndex        =   43
            Top             =   600
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1402
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
            MICON           =   "frmListManualAtt.frx":48B2
            PICN            =   "frmListManualAtt.frx":48CE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txt_employee_code 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4710
            TabIndex        =   37
            Top             =   270
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DTPicker_Periode 
            Height          =   315
            Left            =   960
            TabIndex        =   44
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MM-yyyy"
            Format          =   94699523
            CurrentDate     =   40794
         End
         Begin prj_tpc.vbButton cmdPL 
            Height          =   795
            Left            =   4320
            TabIndex        =   113
            Top             =   600
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1402
            BTYPE           =   14
            TX              =   "&PL"
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
            MICON           =   "frmListManualAtt.frx":5960
            PICN            =   "frmListManualAtt.frx":597C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label9 
            Caption         =   "PERIODE"
            Height          =   195
            Left            =   60
            TabIndex        =   45
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "DATE"
            Height          =   195
            Left            =   60
            TabIndex        =   42
            Top             =   990
            Width           =   435
         End
         Begin VB.Label Label16 
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
            Left            =   2160
            TabIndex        =   41
            Top             =   990
            Width           =   285
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "EMPLOYEE"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   300
            Width           =   870
         End
      End
      Begin VB.Frame Frame3 
         Height          =   7980
         Left            =   -74850
         TabIndex        =   26
         Top             =   450
         Width           =   14715
         Begin prj_tpc.LynxGrid LynxGrid2 
            Height          =   7635
            Left            =   120
            TabIndex        =   27
            Top             =   210
            Width           =   14475
            _ExtentX        =   25532
            _ExtentY        =   13467
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
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
      Begin VB.Frame Frame2 
         Height          =   7980
         Left            =   -74850
         TabIndex        =   19
         Top             =   390
         Width           =   14715
         Begin prj_tpc.LynxGrid LynxGrid1 
            Height          =   7635
            Left            =   120
            TabIndex        =   20
            Top             =   210
            Width           =   14475
            _ExtentX        =   25532
            _ExtentY        =   13467
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   0
         Top             =   7530
      End
      Begin prj_tpc.vbButton cmdPrev 
         Height          =   285
         Left            =   2700
         TabIndex        =   16
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "<"
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
         MICON           =   "frmListManualAtt.frx":6A0E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame1 
         Height          =   7905
         Left            =   90
         TabIndex        =   5
         Top             =   1860
         Width           =   14865
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Height          =   135
            Left            =   14790
            TabIndex        =   112
            Top             =   90
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Frame fraOT 
            Height          =   765
            Left            =   60
            TabIndex        =   109
            Top             =   3780
            Width           =   3645
            Begin prj_tpc.vbButton cmdDelOT 
               Height          =   495
               Left            =   1830
               TabIndex        =   110
               Top             =   180
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":6A2A
               PICN            =   "frmListManualAtt.frx":6A46
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdCancelOT 
               Height          =   495
               Left            =   2490
               TabIndex        =   111
               Top             =   180
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":7AD8
               PICN            =   "frmListManualAtt.frx":7AF4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdSyncOT 
               Height          =   495
               Left            =   510
               TabIndex        =   114
               Top             =   180
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":8B86
               PICN            =   "frmListManualAtt.frx":8BA2
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdApprove 
               Height          =   495
               Left            =   1170
               TabIndex        =   115
               Top             =   180
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":9C34
               PICN            =   "frmListManualAtt.frx":9C50
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
         Begin VB.Frame fraPL 
            Height          =   765
            Left            =   60
            TabIndex        =   106
            Top             =   2340
            Width           =   3645
            Begin prj_tpc.vbButton cmdDelPL 
               Height          =   495
               Left            =   1170
               TabIndex        =   107
               Top             =   180
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":ACE2
               PICN            =   "frmListManualAtt.frx":ACFE
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdCancelPL 
               Height          =   495
               Left            =   1830
               TabIndex        =   108
               Top             =   180
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":BD90
               PICN            =   "frmListManualAtt.frx":BDAC
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
         Begin VB.Frame fraInOut 
            Height          =   1245
            Left            =   60
            TabIndex        =   58
            Top             =   6570
            Width           =   3645
            Begin VB.OptionButton optOut 
               Caption         =   "OUT"
               Height          =   255
               Left            =   1260
               TabIndex        =   60
               Top             =   420
               Width           =   825
            End
            Begin VB.OptionButton optIn 
               Caption         =   "IN"
               Height          =   255
               Left            =   600
               TabIndex        =   59
               Top             =   420
               Width           =   825
            End
            Begin prj_tpc.vbButton cmdOK 
               Height          =   495
               Left            =   2250
               TabIndex        =   61
               Top             =   660
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":CE3E
               PICN            =   "frmListManualAtt.frx":CE5A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdBatal 
               Height          =   495
               Left            =   2910
               TabIndex        =   62
               Top             =   660
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   873
               BTYPE           =   14
               TX              =   ""
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
               MICON           =   "frmListManualAtt.frx":DEEC
               PICN            =   "frmListManualAtt.frx":DF08
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
         Begin VB.Frame fraByPass 
            Caption         =   "By Pass Button"
            Height          =   1155
            Left            =   4740
            TabIndex        =   46
            Top             =   6630
            Visible         =   0   'False
            Width           =   9225
            Begin prj_tpc.vbButton cmdEditByPass 
               Height          =   705
               Left            =   120
               TabIndex        =   47
               Top             =   270
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
               MICON           =   "frmListManualAtt.frx":EF9A
               PICN            =   "frmListManualAtt.frx":EFB6
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdAttByPass 
               Height          =   705
               Left            =   1110
               TabIndex        =   48
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               BTYPE           =   14
               TX              =   "&Att. Form"
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
               MICON           =   "frmListManualAtt.frx":10048
               PICN            =   "frmListManualAtt.frx":10064
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdPLByPass 
               Height          =   705
               Left            =   2100
               TabIndex        =   49
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               BTYPE           =   14
               TX              =   "&Pr. Leave"
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
               MICON           =   "frmListManualAtt.frx":110F6
               PICN            =   "frmListManualAtt.frx":11112
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdOTByPass 
               Height          =   705
               Left            =   3090
               TabIndex        =   50
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               BTYPE           =   14
               TX              =   "&Overtime"
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
               MICON           =   "frmListManualAtt.frx":121A4
               PICN            =   "frmListManualAtt.frx":121C0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdLeaveByPass 
               Height          =   705
               Left            =   4080
               TabIndex        =   51
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               BTYPE           =   14
               TX              =   "&Leave"
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
               MICON           =   "frmListManualAtt.frx":13252
               PICN            =   "frmListManualAtt.frx":1326E
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
               Left            =   7860
               TabIndex        =   52
               Top             =   270
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
               MICON           =   "frmListManualAtt.frx":14300
               PICN            =   "frmListManualAtt.frx":1431C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdChShiftByPass 
               Height          =   705
               Left            =   5070
               TabIndex        =   53
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               BTYPE           =   14
               TX              =   "&Ch. Shift"
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
               MICON           =   "frmListManualAtt.frx":153AE
               PICN            =   "frmListManualAtt.frx":153CA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prj_tpc.vbButton cmdTraining 
               Height          =   705
               Left            =   6480
               TabIndex        =   116
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               BTYPE           =   14
               TX              =   "&Training"
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
               MICON           =   "frmListManualAtt.frx":1645C
               PICN            =   "frmListManualAtt.frx":16478
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Shift 
            Height          =   1455
            Left            =   60
            TabIndex        =   15
            Top             =   180
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   2566
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "SHIFT CODE"
            Columns(0).DataField=   "shift_code"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "SHIFT NAME"
            Columns(1).DataField=   "shift_name"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "FLAG DAY OVER"
            Columns(2).DataField=   "flag_day_over"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3545"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3466"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
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
            Caption         =   "LIST OF SHIFT"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Att 
            Height          =   7605
            Left            =   3750
            TabIndex        =   18
            Top             =   180
            Width           =   11025
            _ExtentX        =   19447
            _ExtentY        =   13414
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
            Columns(1).Caption=   "DATE"
            Columns(1).DataField=   "tgl"
            Columns(1).NumberFormat=   "dd/MM/yyyy"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "ATT DATE"
            Columns(2).DataField=   "att_date"
            Columns(2).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "DAY NAME"
            Columns(3).DataField=   "day_name"
            Columns(3).NumberFormat=   "DDDD"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "EMP. CODE"
            Columns(4).DataField=   "nik"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "EMP. NAME"
            Columns(5).DataField=   "employee_name"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "SHIFT NAME"
            Columns(6).DataField=   "shift_name"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "STATUS CODE"
            Columns(7).DataField=   "status"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "TITLE CODE"
            Columns(8).DataField=   "title_code"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "JOB TITLE"
            Columns(9).DataField=   "title_name"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "TIME IN"
            Columns(10).DataField=   "time_in"
            Columns(10).NumberFormat=   "hh:mm"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "TIME OUT"
            Columns(11).DataField=   "time_out"
            Columns(11).NumberFormat=   "hh:mm"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "SCH. SHIFT"
            Columns(12).DataField=   "act_shift"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "ACT SHIFT"
            Columns(13).DataField=   "shift_code"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "STATUS"
            Columns(14).DataField=   "absent_name"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "ENTRY DATE"
            Columns(15).DataField=   "entry_date"
            Columns(15).NumberFormat=   "yyyy-MM-dd"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "DEPT. CODE"
            Columns(16).DataField=   "department_code"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "DEPT. NAME"
            Columns(17).DataField=   "department_name"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "DIV. CODE"
            Columns(18).DataField=   "division_code"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(19)._VlistStyle=   0
            Columns(19)._MaxComboItems=   5
            Columns(19).Caption=   "DIV. NAME"
            Columns(19).DataField=   "division_name"
            Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(20)._VlistStyle=   0
            Columns(20)._MaxComboItems=   5
            Columns(20).Caption=   "DESCRIPTION"
            Columns(20).DataField=   "descriptioon"
            Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(21)._VlistStyle=   0
            Columns(21)._MaxComboItems=   5
            Columns(21).Caption=   "STATUS NAME"
            Columns(21).DataField=   "absent_name"
            Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(22)._VlistStyle=   0
            Columns(22)._MaxComboItems=   5
            Columns(22).Caption=   "FLAG LATE"
            Columns(22).DataField=   "flag_inc_late"
            Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(23)._VlistStyle=   0
            Columns(23)._MaxComboItems=   5
            Columns(23).Caption=   "SHIFT CODE"
            Columns(23).DataField=   "shift"
            Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(24)._VlistStyle=   0
            Columns(24)._MaxComboItems=   5
            Columns(24).Caption=   "FLAG DAY OVER"
            Columns(24).DataField=   "flag_day_over"
            Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(25)._VlistStyle=   0
            Columns(25)._MaxComboItems=   5
            Columns(25).Caption=   "ENROLLNUMBER"
            Columns(25).DataField=   "enrollnumber"
            Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(26)._VlistStyle=   4
            Columns(26)._MaxComboItems=   5
            Columns(26).Caption=   "TRAINING"
            Columns(26).DataField=   "flag_training"
            Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   27
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=27"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2064"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1984"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2355"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2275"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=2249"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2170"
            Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=1746"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1667"
            Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=3069"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2990"
            Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(40)=   "Column(7).Width=1614"
            Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=1535"
            Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(44)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(46)=   "Column(8).Width=185"
            Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=106"
            Splits(0)._ColumnProps(49)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(50)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(52)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(56)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(58)=   "Column(9)._MinWidth=10"
            Splits(0)._ColumnProps(59)=   "Column(10).Width=1535"
            Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=1455"
            Splits(0)._ColumnProps(62)=   "Column(10)._ColStyle=513"
            Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(64)=   "Column(10)._MinWidth=54215968"
            Splits(0)._ColumnProps(65)=   "Column(11).Width=1588"
            Splits(0)._ColumnProps(66)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(67)=   "Column(11)._WidthInPix=1508"
            Splits(0)._ColumnProps(68)=   "Column(11)._ColStyle=513"
            Splits(0)._ColumnProps(69)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(70)=   "Column(11)._MinWidth=54215968"
            Splits(0)._ColumnProps(71)=   "Column(12).Width=1773"
            Splits(0)._ColumnProps(72)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(73)=   "Column(12)._WidthInPix=1693"
            Splits(0)._ColumnProps(74)=   "Column(12)._ColStyle=513"
            Splits(0)._ColumnProps(75)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(76)=   "Column(13).Width=1667"
            Splits(0)._ColumnProps(77)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(78)=   "Column(13)._WidthInPix=1588"
            Splits(0)._ColumnProps(79)=   "Column(13)._ColStyle=513"
            Splits(0)._ColumnProps(80)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(81)=   "Column(14).Width=2937"
            Splits(0)._ColumnProps(82)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(83)=   "Column(14)._WidthInPix=2858"
            Splits(0)._ColumnProps(84)=   "Column(14)._ColStyle=513"
            Splits(0)._ColumnProps(85)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(86)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(87)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(88)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(89)=   "Column(15)._ColStyle=516"
            Splits(0)._ColumnProps(90)=   "Column(15).Visible=0"
            Splits(0)._ColumnProps(91)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(92)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(93)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(94)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(95)=   "Column(16)._ColStyle=516"
            Splits(0)._ColumnProps(96)=   "Column(16).Visible=0"
            Splits(0)._ColumnProps(97)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(98)=   "Column(17).Width=2725"
            Splits(0)._ColumnProps(99)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(100)=   "Column(17)._WidthInPix=2646"
            Splits(0)._ColumnProps(101)=   "Column(17)._ColStyle=516"
            Splits(0)._ColumnProps(102)=   "Column(17).Visible=0"
            Splits(0)._ColumnProps(103)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(104)=   "Column(18).Width=2725"
            Splits(0)._ColumnProps(105)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(106)=   "Column(18)._WidthInPix=2646"
            Splits(0)._ColumnProps(107)=   "Column(18)._ColStyle=516"
            Splits(0)._ColumnProps(108)=   "Column(18).Visible=0"
            Splits(0)._ColumnProps(109)=   "Column(18).Order=19"
            Splits(0)._ColumnProps(110)=   "Column(19).Width=2725"
            Splits(0)._ColumnProps(111)=   "Column(19).DividerColor=0"
            Splits(0)._ColumnProps(112)=   "Column(19)._WidthInPix=2646"
            Splits(0)._ColumnProps(113)=   "Column(19)._ColStyle=516"
            Splits(0)._ColumnProps(114)=   "Column(19).Visible=0"
            Splits(0)._ColumnProps(115)=   "Column(19).Order=20"
            Splits(0)._ColumnProps(116)=   "Column(20).Width=2725"
            Splits(0)._ColumnProps(117)=   "Column(20).DividerColor=0"
            Splits(0)._ColumnProps(118)=   "Column(20)._WidthInPix=2646"
            Splits(0)._ColumnProps(119)=   "Column(20)._ColStyle=516"
            Splits(0)._ColumnProps(120)=   "Column(20).Visible=0"
            Splits(0)._ColumnProps(121)=   "Column(20).Order=21"
            Splits(0)._ColumnProps(122)=   "Column(21).Width=2725"
            Splits(0)._ColumnProps(123)=   "Column(21).DividerColor=0"
            Splits(0)._ColumnProps(124)=   "Column(21)._WidthInPix=2646"
            Splits(0)._ColumnProps(125)=   "Column(21)._ColStyle=516"
            Splits(0)._ColumnProps(126)=   "Column(21).Visible=0"
            Splits(0)._ColumnProps(127)=   "Column(21).Order=22"
            Splits(0)._ColumnProps(128)=   "Column(22).Width=2725"
            Splits(0)._ColumnProps(129)=   "Column(22).DividerColor=0"
            Splits(0)._ColumnProps(130)=   "Column(22)._WidthInPix=2646"
            Splits(0)._ColumnProps(131)=   "Column(22)._ColStyle=516"
            Splits(0)._ColumnProps(132)=   "Column(22).Visible=0"
            Splits(0)._ColumnProps(133)=   "Column(22).Order=23"
            Splits(0)._ColumnProps(134)=   "Column(23).Width=2725"
            Splits(0)._ColumnProps(135)=   "Column(23).DividerColor=0"
            Splits(0)._ColumnProps(136)=   "Column(23)._WidthInPix=2646"
            Splits(0)._ColumnProps(137)=   "Column(23)._ColStyle=516"
            Splits(0)._ColumnProps(138)=   "Column(23).Visible=0"
            Splits(0)._ColumnProps(139)=   "Column(23).Order=24"
            Splits(0)._ColumnProps(140)=   "Column(24).Width=2725"
            Splits(0)._ColumnProps(141)=   "Column(24).DividerColor=0"
            Splits(0)._ColumnProps(142)=   "Column(24)._WidthInPix=2646"
            Splits(0)._ColumnProps(143)=   "Column(24)._ColStyle=516"
            Splits(0)._ColumnProps(144)=   "Column(24).Visible=0"
            Splits(0)._ColumnProps(145)=   "Column(24).Order=25"
            Splits(0)._ColumnProps(146)=   "Column(25).Width=2725"
            Splits(0)._ColumnProps(147)=   "Column(25).DividerColor=0"
            Splits(0)._ColumnProps(148)=   "Column(25)._WidthInPix=2646"
            Splits(0)._ColumnProps(149)=   "Column(25)._ColStyle=516"
            Splits(0)._ColumnProps(150)=   "Column(25).Visible=0"
            Splits(0)._ColumnProps(151)=   "Column(25).Order=26"
            Splits(0)._ColumnProps(152)=   "Column(26).Width=1561"
            Splits(0)._ColumnProps(153)=   "Column(26).DividerColor=0"
            Splits(0)._ColumnProps(154)=   "Column(26)._WidthInPix=1482"
            Splits(0)._ColumnProps(155)=   "Column(26)._ColStyle=513"
            Splits(0)._ColumnProps(156)=   "Column(26).Order=27"
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
            Caption         =   "LIST OF ATTENDANCE"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=106,.parent=13,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=126,.parent=13,.alignment=3"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=123,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=124,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=125,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=98,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=95,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=96,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=97,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=102,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=99,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=100,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=101,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=110,.parent=13,.alignment=2"
            _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=107,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=108,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=109,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=130,.parent=13,.alignment=2"
            _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=127,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=128,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=129,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=134,.parent=13,.alignment=2"
            _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=131,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=132,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=133,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=138,.parent=13,.alignment=2"
            _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=135,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=136,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=137,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=28,.parent=13"
            _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=25,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=26,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=27,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=46,.parent=13"
            _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=43,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=44,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=45,.parent=17"
            _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=58,.parent=13"
            _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=55,.parent=14"
            _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=56,.parent=15"
            _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=57,.parent=17"
            _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=70,.parent=13"
            _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=67,.parent=14"
            _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=68,.parent=15"
            _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=69,.parent=17"
            _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=78,.parent=13"
            _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=75,.parent=14"
            _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=76,.parent=15"
            _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=77,.parent=17"
            _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=82,.parent=13"
            _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=79,.parent=14"
            _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=80,.parent=15"
            _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=81,.parent=17"
            _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=86,.parent=13"
            _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=83,.parent=14"
            _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=84,.parent=15"
            _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=85,.parent=17"
            _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=114,.parent=13"
            _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=111,.parent=14"
            _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=112,.parent=15"
            _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=113,.parent=17"
            _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=94,.parent=13"
            _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=91,.parent=14"
            _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=92,.parent=15"
            _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=93,.parent=17"
            _StyleDefs(130) =   "Splits(0).Columns(24).Style:id=118,.parent=13"
            _StyleDefs(131) =   "Splits(0).Columns(24).HeadingStyle:id=115,.parent=14"
            _StyleDefs(132) =   "Splits(0).Columns(24).FooterStyle:id=116,.parent=15"
            _StyleDefs(133) =   "Splits(0).Columns(24).EditorStyle:id=117,.parent=17"
            _StyleDefs(134) =   "Splits(0).Columns(25).Style:id=122,.parent=13"
            _StyleDefs(135) =   "Splits(0).Columns(25).HeadingStyle:id=119,.parent=14"
            _StyleDefs(136) =   "Splits(0).Columns(25).FooterStyle:id=120,.parent=15"
            _StyleDefs(137) =   "Splits(0).Columns(25).EditorStyle:id=121,.parent=17"
            _StyleDefs(138) =   "Splits(0).Columns(26).Style:id=142,.parent=13,.alignment=2"
            _StyleDefs(139) =   "Splits(0).Columns(26).HeadingStyle:id=139,.parent=14"
            _StyleDefs(140) =   "Splits(0).Columns(26).FooterStyle:id=140,.parent=15"
            _StyleDefs(141) =   "Splits(0).Columns(26).EditorStyle:id=141,.parent=17"
            _StyleDefs(142) =   "Named:id=33:Normal"
            _StyleDefs(143) =   ":id=33,.parent=0"
            _StyleDefs(144) =   "Named:id=34:Heading"
            _StyleDefs(145) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(146) =   ":id=34,.wraptext=-1"
            _StyleDefs(147) =   "Named:id=35:Footing"
            _StyleDefs(148) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(149) =   "Named:id=36:Selected"
            _StyleDefs(150) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(151) =   "Named:id=37:Caption"
            _StyleDefs(152) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(153) =   "Named:id=38:HighlightRow"
            _StyleDefs(154) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(155) =   "Named:id=39:EvenRow"
            _StyleDefs(156) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(157) =   "Named:id=40:OddRow"
            _StyleDefs(158) =   ":id=40,.parent=33"
            _StyleDefs(159) =   "Named:id=41:RecordSelector"
            _StyleDefs(160) =   ":id=41,.parent=34"
            _StyleDefs(161) =   "Named:id=42:FilterBar"
            _StyleDefs(162) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   1785
            Left            =   60
            TabIndex        =   54
            Top             =   6030
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   3149
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
            Columns(1).Caption=   "DATE"
            Columns(1).DataField=   "date_att"
            Columns(1).NumberFormat=   "dd/MM/yyyy"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TIME"
            Columns(2).DataField=   "time_att"
            Columns(2).NumberFormat=   "HH:mm:ss"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "TYPE"
            Columns(3).DataField=   "type"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "ATT DATE"
            Columns(4).DataField=   "att_date"
            Columns(4).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).Size  =   2
            Splits(0).Size.vt=   2
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   3
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2143"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2064"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1958"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1879"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1270"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1191"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=3678"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3598"
            Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos(0).PreviewInitHeight=   -1
            PrintInfos(0).PreviewInitScreenFill=   -1
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "LIST OF LOG ATTENDANCE"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   1425
            Left            =   60
            TabIndex        =   55
            Top             =   1680
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   2514
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "DATE"
            Columns(0).DataField=   "pl_date"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "EMPLOYEE CODE"
            Columns(1).DataField=   "employee_code"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TIME IN"
            Columns(2).DataField=   "time_in"
            Columns(2).NumberFormat=   "HH:mm"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   4
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "BREAK"
            Columns(3).DataField=   "int_break"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "TIMEOUT"
            Columns(4).DataField=   "time_out"
            Columns(4).NumberFormat=   "HH:mm"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "APPR."
            Columns(5).DataField=   "flag_approval"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "TYPE"
            Columns(6).DataField=   "type"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "SEQ"
            Columns(7).DataField=   "seq"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2963"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2884"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1720"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1640"
            Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1376"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1296"
            Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=1667"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1588"
            Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=979"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=900"
            Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=1005"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=926"
            Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(39)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos(0).PreviewInitHeight=   -1
            PrintInfos(0).PreviewInitScreenFill=   -1
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "LIST OF PRIVATE LEAVE"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
            Height          =   1395
            Left            =   60
            TabIndex        =   56
            Top             =   3150
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   2461
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "DATE"
            Columns(0).DataField=   "date"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "EMPLOYEE CODE"
            Columns(1).DataField=   "employee_code"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TIME IN"
            Columns(2).DataField=   "start_time"
            Columns(2).NumberFormat=   "HH:mm"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   4
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "BREAK"
            Columns(3).DataField=   "int_break"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "TIME OUT"
            Columns(4).DataField=   "end_time"
            Columns(4).NumberFormat=   "HH:mm"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "APPROVE"
            Columns(5).DataField=   "flag_approval"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "SEQ"
            Columns(6).DataField=   "seq"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "FLAG CHANGE STATUS"
            Columns(7).DataField=   "flag_change_status"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2963"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2884"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2037"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1958"
            Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1217"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1138"
            Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=1879"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1799"
            Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=1482"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1402"
            Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
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
            PrintInfos(0).PreviewInitHeight=   -1
            PrintInfos(0).PreviewInitScreenFill=   -1
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "LIST OF OVERTIME"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid4 
            Height          =   1395
            Left            =   60
            TabIndex        =   57
            Top             =   4590
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   2461
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
            Columns(1).Caption=   "DATE"
            Columns(1).DataField=   "att_date"
            Columns(1).NumberFormat=   "yyyy-MM-dd"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "DESCRIPTION"
            Columns(2).DataField=   "description"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2170"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2090"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=3228"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3149"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos(0).PreviewInitHeight=   -1
            PrintInfos(0).PreviewInitScreenFill=   -1
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "LIST OF LEAVE"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
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
      End
      Begin VB.TextBox txt_group_shift 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1530
         Width           =   2415
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1170
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   94699523
         CurrentDate     =   40794
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_Group_Shift 
         Height          =   375
         Left            =   1350
         OleObjectBlob   =   "frmListManualAtt.frx":1750A
         TabIndex        =   6
         Top             =   1530
         Width           =   1095
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1350
         OleObjectBlob   =   "frmListManualAtt.frx":19468
         TabIndex        =   7
         Top             =   1170
         Width           =   1095
      End
      Begin prj_tpc.vbButton cmdRefresh 
         Height          =   495
         Left            =   3510
         TabIndex        =   8
         Top             =   390
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "Refresh"
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
         MICON           =   "frmListManualAtt.frx":1B3CE
         PICN            =   "frmListManualAtt.frx":1B3EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdNext 
         Height          =   285
         Left            =   3030
         TabIndex        =   17
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   ">"
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
         MICON           =   "frmListManualAtt.frx":1C47C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   -69030
         TabIndex        =   21
         Top             =   8640
         Visible         =   0   'False
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prj_tpc.vbButton cmdImport 
         Height          =   555
         Left            =   -73020
         TabIndex        =   22
         Top             =   8610
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Import"
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
         MICON           =   "frmListManualAtt.frx":1C498
         PICN            =   "frmListManualAtt.frx":1C4B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdBrowse 
         Height          =   555
         Left            =   -74700
         TabIndex        =   23
         Top             =   8610
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Browse"
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
         MICON           =   "frmListManualAtt.frx":1D546
         PICN            =   "frmListManualAtt.frx":1D562
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   225
         Left            =   -69030
         TabIndex        =   28
         Top             =   8700
         Visible         =   0   'False
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prj_tpc.vbButton cmdImportM 
         Height          =   555
         Left            =   -73020
         TabIndex        =   29
         Top             =   8670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Import"
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
         MICON           =   "frmListManualAtt.frx":1E5F4
         PICN            =   "frmListManualAtt.frx":1E610
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdBrowseM 
         Height          =   555
         Left            =   -74700
         TabIndex        =   30
         Top             =   8670
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Browse"
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
         MICON           =   "frmListManualAtt.frx":1F6A2
         PICN            =   "frmListManualAtt.frx":1F6BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_reproccess 
         Height          =   675
         Left            =   10020
         TabIndex        =   63
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1191
         BTYPE           =   14
         TX              =   "&Reproccess"
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
         MICON           =   "frmListManualAtt.frx":20750
         PICN            =   "frmListManualAtt.frx":2076C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl_hk_15 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   13620
         TabIndex        =   105
         Top             =   570
         Width           =   375
      End
      Begin VB.Label lbl_hk_20 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   13620
         TabIndex        =   104
         Top             =   750
         Width           =   375
      End
      Begin VB.Label lbl_hl_20 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   14430
         TabIndex        =   103
         Top             =   570
         Width           =   405
      End
      Begin VB.Label lbl_hl_30 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   14430
         TabIndex        =   102
         Top             =   750
         Width           =   405
      End
      Begin VB.Label lbl_hl_40 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   14430
         TabIndex        =   101
         Top             =   930
         Width           =   405
      End
      Begin VB.Label lbl_meal 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   12600
         TabIndex        =   100
         Top             =   390
         Width           =   405
      End
      Begin VB.Label lbl_trans 
         Caption         =   "0"
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
         Left            =   12600
         TabIndex        =   99
         Top             =   570
         Width           =   405
      End
      Begin VB.Label lbl_a 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   12600
         TabIndex        =   98
         Top             =   750
         Width           =   405
      End
      Begin VB.Label lbl_n 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12600
         TabIndex        =   97
         Top             =   930
         Width           =   405
      End
      Begin VB.Label lbl_pl 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   11610
         TabIndex        =   96
         Top             =   570
         Width           =   405
      End
      Begin VB.Label lbl_sick 
         Caption         =   "0"
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
         Left            =   11610
         TabIndex        =   95
         Top             =   750
         Width           =   405
      End
      Begin VB.Label lbl_late 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11610
         TabIndex        =   94
         Top             =   930
         Width           =   405
      End
      Begin VB.Label lbl_l 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11610
         TabIndex        =   93
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label38 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14370
         TabIndex        =   92
         Top             =   930
         Width           =   45
      End
      Begin VB.Label Label37 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14370
         TabIndex        =   91
         Top             =   570
         Width           =   45
      End
      Begin VB.Label Label35 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14370
         TabIndex        =   90
         Top             =   750
         Width           =   45
      End
      Begin VB.Label Label36 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13560
         TabIndex        =   89
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label34 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13560
         TabIndex        =   88
         Top             =   540
         Width           =   45
      End
      Begin VB.Label Label33 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12540
         TabIndex        =   87
         Top             =   930
         Width           =   45
      End
      Begin VB.Label Label32 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12540
         TabIndex        =   86
         Top             =   750
         Width           =   45
      End
      Begin VB.Label Label31 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12540
         TabIndex        =   85
         Top             =   390
         Width           =   45
      End
      Begin VB.Label Label30 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12540
         TabIndex        =   84
         Top             =   570
         Width           =   45
      End
      Begin VB.Label Label29 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11550
         TabIndex        =   83
         Top             =   930
         Width           =   45
      End
      Begin VB.Label Label28 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11550
         TabIndex        =   82
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label27 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11550
         TabIndex        =   81
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label26 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11550
         TabIndex        =   80
         Top             =   540
         Width           =   45
      End
      Begin VB.Label Label25 
         Caption         =   "X 3"
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
         Left            =   14070
         TabIndex        =   79
         Top             =   750
         Width           =   315
      End
      Begin VB.Label Label24 
         Caption         =   "X 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14070
         TabIndex        =   78
         Top             =   930
         Width           =   315
      End
      Begin VB.Label Label23 
         Caption         =   "X 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14070
         TabIndex        =   77
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "X 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13110
         TabIndex        =   76
         Top             =   750
         Width           =   405
      End
      Begin VB.Label Label21 
         Caption         =   "X 1.5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13110
         TabIndex        =   75
         Top             =   570
         Width           =   465
      End
      Begin VB.Label Label20 
         Caption         =   "OT HL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14070
         TabIndex        =   74
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "OT HK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13080
         TabIndex        =   73
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   12090
         TabIndex        =   72
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12090
         TabIndex        =   71
         Top             =   930
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "Trans"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12060
         TabIndex        =   70
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Meal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12060
         TabIndex        =   69
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label13 
         Caption         =   "Late"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11160
         TabIndex        =   68
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "SL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11160
         TabIndex        =   67
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "PL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11160
         TabIndex        =   66
         Top             =   570
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11160
         TabIndex        =   65
         Top             =   390
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Label1"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   -69030
         TabIndex        =   31
         Top             =   9000
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label Label5 
         Caption         =   "Label1"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   -69030
         TabIndex        =   24
         Top             =   8940
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "DATE"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GROUP SHIFT"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COMPANY"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1200
         Width           =   795
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MANUAL ATTENDANCE"
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
      Left            =   390
      TabIndex        =   1
      Top             =   150
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frmListManualAtt.frx":217FE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16860
   End
End
Attribute VB_Name = "frm_list_manual_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Dim rs As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsGroupShift As New ADODB.Recordset
Dim rsShift As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset
Dim rsImportAtt As New ADODB.Recordset

Dim rsLogAtt As New ADODB.Recordset
Dim rsSPL As New ADODB.Recordset
Dim rsPL As New ADODB.Recordset
Dim rsEmpLeave As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim ButtonFlag As Integer
Dim Saturday As New TrueOleDBGrid70.Style
Dim Sunday As New TrueOleDBGrid70.Style
Dim Holiday As New TrueOleDBGrid70.Style

Dim abstatus As Integer
Dim v_flag_present As Integer
Dim v_flag_duty As Integer
Dim vFlagTraining As Integer

Dim vMode As Integer

Private Sub cmdDelPL_Click()
Dim i As Integer

    i = MsgBox("Are you sure want to delete this private leave ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    SQL = "DELETE FROM t_private_leave " & _
          "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
            "AND pl_date = '" & Format(TDBGrid2.Columns("pl_date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
            "AND seq = '" & TDBGrid2.Columns("seq").Value & "'"
    CnG.Execute SQL
    
    Call load_data_pl
    
    fraPL.Visible = False
    TDBGrid2.Enabled = True
End Sub

Private Sub cmdDelOT_Click()
Dim i As Integer
    
    i = MsgBox("Are you sure want to delete this overtime ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    SQL = "DELETE FROM t_spl " & _
          "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
            "AND date = '" & Format(TDBGrid3.Columns("date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
            "AND seq = '" & TDBGrid3.Columns("seq").Value & "'"
    CnG.Execute SQL
    
    Call load_data_spl
    
    fraOT.Visible = False
    TDBGrid3.Enabled = True
End Sub

Private Sub cmdSyncOT_Click()
Dim i As Integer
    
    i = MsgBox("Are you sure want to syncronize this overtime ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    SQL = "UPDATE t_spl SET date = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
            "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
            "AND date = '" & Format(TDBGrid3.Columns("date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
            "AND seq = '" & TDBGrid3.Columns("seq").Value & "'"
    CnG.Execute SQL
    
    If TDBGrid3.Columns("flag_approval").Value <> 0 Then
        If TDBGrid3.Columns("flag_change_status").Value <> 0 Then
            SQL = "UPDATE h_attendance SET status = 'OT', flag_manual = 1 " & _
                  "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
                    "and att_date= '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "'"
            CnG.Execute SQL
        End If
    Else
        MsgBox "SPL doesn't approved...", vbExclamation, headerMSG
        fraOT.Visible = False
        TDBGrid3.Enabled = True
        Exit Sub
    End If
    
    Call load_data_spl
    Call load_data_att
    
    fraOT.Visible = False
    TDBGrid3.Enabled = True
End Sub

Private Sub cmdApprove_Click()
Dim i As Integer
    
    i = MsgBox("Are you sure want to approve this overtime ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    SQL = "UPDATE t_spl SET flag_approval = 1," & _
            "user_approval = '" & LOGIN_NAME & "' " & _
          "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
                "and date = '" & Format(TDBGrid3.Columns("date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
                "and company_code= '" & TDBCombo_company.Text & "' " & _
                "and seq = " & TDBGrid3.Columns("seq").Value & ""
    CnG.Execute SQL
    
    If TDBGrid3.Columns("flag_change_status").Value <> 0 Then
        SQL = "UPDATE h_attendance SET status = 'OT', flag_manual = 1 " & _
              "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
                "and att_date= '" & Format(TDBGrid_Att.Columns("date").Value, "yyyy-MM-dd HH:mm:ss") & "'"
        CnG.Execute SQL
    End If
    
    Call load_data_spl
    Call load_data_att
    
    fraOT.Visible = False
    TDBGrid3.Enabled = True
End Sub

Private Sub cmdEditByPass_Click()
    showData
    
    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub cmdAttByPass_Click()
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If
    
    With frm_trans_manual_check
        .int_mode = 1
        vModeLoad = 1
        .btnApprove.Visible = False
        .lblVerify.Caption = ""
        .lblVerify.Visible = False
        .fra_entry.Visible = True
        .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        
        .txt_employee_code.Text = IIf(TDBGrid_Att.Columns("employee_code").Value = "", txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
        .txt_nik.Text = IIf(TDBGrid_Att.Columns("nik").Value = "", txt_nik.Text, TDBGrid_Att.Columns("nik").Value)
        .txt_employee_name.Text = IIf(TDBGrid_Att.Columns("employee_name").Value = "", txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
        .cbo_flag_io.ListIndex = 0
        .DTPicker_check.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        
        .Show 1
    End With
    
    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub cmdOK_Click()
    SQL = "UPDATE h_log_attendance SET flag_io = '" & IIf(optIn.Value, 0, 1) & "' " & _
          "WHERE enrollnumber = '" & Val(TDBGrid_Att.Columns("enrollnumber").Value) & "' " & _
            "AND att_date = '" & Format(TDBGrid1.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "'"
    CnG.Execute SQL
    
    fraInOut.Visible = False
    TDBGrid1.Enabled = True
    
    Call load_data_log_att
End Sub

Private Sub cmdBatal_Click()
    fraInOut.Visible = False
    TDBGrid1.Enabled = True
End Sub

Private Sub cmdCancelPL_Click()
    fraPL.Visible = False
    TDBGrid2.Enabled = True
End Sub

Private Sub cmdCancelOT_Click()
    fraOT.Visible = False
    TDBGrid3.Enabled = True
End Sub

Private Sub cmdPLByPass_Click()
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If
    
    With frm_trans_private_leave
        .int_mode = 1
        vModeLoad = 1
        .btnApprove.Visible = False
        .lblVerify.Caption = ""
        .lblVerify.Visible = False
        .fra_entry.Visible = True
        .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        
        .txt_employee_code.Text = IIf(TDBGrid_Att.Columns("employee_code").Value = "", txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
        .txt_nik.Text = IIf(TDBGrid_Att.Columns("nik").Value = "", txt_nik.Text, TDBGrid_Att.Columns("nik").Value)
        .txt_employee_name.Text = IIf(TDBGrid_Att.Columns("employee_name").Value = "", txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
        .cboType.ListIndex = 0
        .DTPicker_PL.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        .txt_description.Text = "PRIVATE LEAVE"
        .txtBreak = 0
        
        .Show
    End With
    
    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub cmdOTByPass_Click()
Dim vFlagShiftable As Integer
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If
    
    With frm_trans_spl
        .int_mode = 1
        vModeLoad = 1
        .load_data_ot
        .fra_entry.Visible = True
        .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        .txtBreak = 0
        
        .txt_employee_code.Text = IIf(TDBGrid_Att.Columns("employee_code").Value = "", txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
        .txt_nik.Text = IIf(TDBGrid_Att.Columns("nik").Value = "", txt_nik.Text, TDBGrid_Att.Columns("nik").Value)
        .txt_employee_name.Text = IIf(TDBGrid_Att.Columns("employee_name").Value = "", txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
        .DTPicker1.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        
        SQL = "SELECT IFNULL(flag_shiftable,0) flag_shiftable FROM m_employee " & _
              "WHERE employee_code = '" & IIf(TDBGrid_Att.Columns("employee_code").Value = "", txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value) & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

        If rscari.RecordCount > 0 Then
            vFlagShiftable = rscari!flag_shiftable
        End If
        rscari.Close
        
        If vFlagShiftable = 0 Then
            If Format(TDBGrid_Att.Columns("tgl").Value, "DDDD") = "Saturday" Or Format(TDBGrid_Att.Columns("tgl").Value, "DDDD") = "Sunday" Then
                .TDBCombo_ot.Text = "HL"
                .txt_ot_name.Text = "HARI LIBUR"
            Else
                .TDBCombo_ot.Text = "HK"
                .txt_ot_name.Text = "HARI KERJA"
            End If
        Else
            If TDBGrid_Att.Columns("act_shift").Value = "OFF" Then
                .TDBCombo_ot.Text = "HL"
                .txt_ot_name.Text = "HARI LIBUR"
            Else
                .TDBCombo_ot.Text = "HK"
                .txt_ot_name.Text = "HARI KERJA"
            End If
        End If
        
        SQL = "SELECT a.ot_code, b.ot_name FROM t_holiday a JOIN m_ot b ON a.ot_code = b.ot_code " & _
              "WHERE Date(a.holiday_date) = '" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            .TDBCombo_ot.Text = rscari!ot_code
            .txt_ot_name.Text = rscari!ot_name
        End If
        rscari.Close
        
        .cari_jam
        .btnApprove.Visible = False
        .lblVerify.Caption = ""
        .lblVerify.Visible = False
        .Show
    End With
    
    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub cmdLeaveByPass_Click()
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If

    With frm_trans_leave
        .int_mode = 1
        vModeLoad = 1
        .SSTab1.TabVisible(1) = False
        .SSTab1.TabVisible(2) = False
        .fra_entry_Emp.Visible = True
'        .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)

        .txt_employee_code.Text = IIf(TDBGrid_Att.Columns("employee_code").Value = "", txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
        .txt_nik.Text = IIf(TDBGrid_Att.Columns("nik").Value = "", txt_nik.Text, TDBGrid_Att.Columns("nik").Value)
        .txt_employee_name.Text = IIf(TDBGrid_Att.Columns("employee_name").Value = "", txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
        .cbo_date_to.ListIndex = 0
        .DTPicker_date_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        .DTPicker_date_to.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        .txt_description.Text = "LEAVE"
        
        .Show
    End With

    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub cmdChShiftByPass_Click()
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If

    With frm_trans_change_shift
        vModeLoad = 1
        
        .txt_employee_code.Text = IIf(TDBGrid_Att.Columns("employee_code").Value = "", txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
        .txt_nik.Text = IIf(TDBGrid_Att.Columns("nik").Value = "", txt_nik.Text, TDBGrid_Att.Columns("nik").Value)
        .txt_employee_name.Text = IIf(TDBGrid_Att.Columns("employee_name").Value = "", txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
        
        .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
        .load_data_att
        .load_data_shift
        
        .Show
    End With

    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub cmdTraining_Click()
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If
    
    vFlagTraining = TDBGrid_Att.Columns("flag_training").Value
    If vFlagTraining = 0 Then
        SQL = "UPDATE h_attendance SET flag_training = 1 " & _
              "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
                "AND att_date = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "'"
    Else
        SQL = "UPDATE h_attendance SET flag_training = 0 " & _
              "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
                "AND att_date = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "'"
    End If
    
    CnG.Execute SQL

    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
    
    Call load_data_att
    Call check_late(TDBGrid1.Columns("employee_code").Value, DTPicker_from.Value, DTPicker_to.Value)
    Call load_sum_att
End Sub

Private Sub cmdCancel_Click()
    fraByPass.Visible = False
    TDBGrid_Att.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    vMode = 0
    DTPicker1.Value = DTPicker1 + 1

    Call load_data_att
End Sub

Private Sub cmdPrev_Click()
    vMode = 0
    DTPicker1.Value = DTPicker1 - 1

    Call load_data_att
End Sub

Private Sub cmdSearch_Click()
    'validasi group_shift
    If Trim(TDBCombo_Group_Shift.Text) = "" Then
        MsgBox "Group Shift not selected!", vbOKOnly + vbInformation, headerMSG
        TDBCombo_Group_Shift.SetFocus
        Exit Sub
    End If
    
    vMode = 1
    
    SQL = "DELETE FROM t_spl_auto " & _
           "WHERE Date(date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
    CnG.Execute SQL
    
    Call days_func(DTPicker_from.Value, DTPicker_to.Value)
    Call check_late(txt_employee_code, DTPicker_from.Value, DTPicker_to.Value)
'    If flagPLAuto() <> 0 Then Call check_pl(txt_employee_code, DTPicker_from.Value, DTPicker_to.Value)
    Call frm_trans_salary_process.auto_overtime(txt_employee_code, DTPicker_from.Value, DTPicker_to.Value)
    Call load_data_att
    Call load_sum_att
    
'    TDBGrid_Att.FetchRowStyle = True
'    TDBGrid_Att.Refresh
End Sub

Private Sub cmdPL_Click()
Dim i As Integer
    i = MsgBox("Are you sure want to proccess PL between '" _
        & Format(DTPicker_from.Value, "dd-MM-yyyy") & "' to '" _
        & Format(DTPicker_to.Value, "dd-MM-yyyy") & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
        
    If flagPLAuto() <> 0 Then Call check_pl(txt_employee_code, DTPicker_from.Value, DTPicker_to.Value)
    
    If rs.State Then rs.Close
    SQL = "SELECT * FROM m_days WHERE DATE(dt) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    ProgressBar3.Visible = True
    ProgressBar3.Value = 0
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            ProgressBar3.Max = rs.RecordCount
            ProgressBar3.Value = ProgressBar3.Value + 1
        rs.MoveNext
        Wend
    End If
    rs.Close
    ProgressBar3.Visible = False
    
    Screen.MousePointer = vbDefault
    
    Call load_data_att
    Call load_sum_att
End Sub

Private Sub Command1_Click()
Dim rsemp As New ADODB.Recordset

    If rsemp.State Then rsemp.Close
    SQL = "SELECT employee_code FROM m_employee WHERE flag_active <> 0"
    rsemp.Open SQL, CnG, adOpenForwardOnly
    
    If rsemp.RecordCount > 0 Then
        rsemp.MoveFirst
        While Not rsemp.EOF
            If rs.State Then rs.Close
            SQL = "SELECT att_date FROM h_attendance " & _
                  "WHERE DATE(att_date) BETWEEN '2012-12-21' AND '" & Format(Now, "yyyy-MM-dd") & "' " & _
                    "AND employee_code = '" & rsemp.Fields(0).Value & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    If rscari.State Then rscari.Close
                    SQL = "SELECT a.shift_code FROM m_shift_new a JOIN td_emp_group b ON a.group_code = b.emp_group_code " & _
                          "WHERE b.employee_code = '" & rsemp.Fields(0).Value & "' AND b.start_date <= '" & Format(rs.Fields(0).Value, "yyyy-MM-dd") & "' " & _
                            "AND date(a.shift_date) = '" & Format(rs.Fields(0).Value, "yyyy-MM-dd") & "' " & _
                          "ORDER BY a.shift_date DESC LIMIT 1"
                    rscari.Open SQL, CnG, adOpenForwardOnly
                    
                    If rscari.RecordCount > 0 Then
                        SQL = "UPDATE h_attendance SET sch_shift = '" & rscari.Fields(0).Value & "' " & _
                              "WHERE employee_code = '" & rsemp.Fields(0).Value & "' " & _
                                "AND att_date = '" & Format(rs.Fields(0).Value, "yyyy-MM-dd HH:mm:ss") & "'"
                        CnG.Execute SQL
                    End If
                    rscari.Close
                    
                    rs.MoveNext
                Wend
            End If
            rs.Close
            
            rsemp.MoveNext
        Wend
    End If
    rsemp.Close
    
End Sub

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_from, DTPicker_to)
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    vMode = 0
    
    Call load_data_att
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    SSTab1.Tab = 0
    fraInOut.Visible = False
    fraPL.Visible = False
    fraOT.Visible = False
    
    Call createGridKar
    Call load_data_company
    oClause = ""
    vMode = 0
    
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
    DTPicker_Periode.Value = Now
    
    SSTab1.TabVisible(1) = False

    timer1.Enabled = True
    
    cmdNew.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_group_shift()
    If rsGroupShift.State Then rsGroupShift.Close
    SQL = "select * from m_shift_group order by group_code"
    rsGroupShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    Set TDBCombo_Group_Shift.RowSource = rsGroupShift
End Sub

Private Sub load_data_shift()
    If rsShift.State Then rsShift.Close
    SQL = "select shift_code,shift_name,flag_day_over " & _
            "from m_shift where group_code = '" & TDBCombo_Group_Shift.Text & "' " & _
          "order by shift_code"
    rsShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    Set TDBGrid_Shift.DataSource = rsShift
End Sub

Public Sub load_data_att()
Dim vParameter As String
Dim vFlagRollable As Integer
Dim x As Integer

    If rsAtt.State Then rsAtt.Close

    SQL = "SELECT flag_rollable FROM m_shift_group WHERE group_code = '" & TDBCombo_Group_Shift.Text & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vFlagRollable = rscari!flag_rollable
    End If
    rscari.Close
    
    vParameter = IIf(txt_nik.Text <> "", _
                        "a.employee_code = '" & txt_employee_code.Text & "' AND (DATE(b.dt) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "')", _
                        "(DATE(b.dt) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') ")
                        
'    If vFlagRollable = 0 Then
'        vParameter = IIf(txt_nik.Text <> "", _
'                        "a.employee_code = '" & txt_employee_code.Text & "' AND (DATE(b.dt) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "')", _
'                        "(DATE(b.dt) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') ")
'    Else
'        vParameter = IIf(txt_nik.Text <> "", _
'                        "b.employee_code = '" & txt_employee_code.Text & "' AND (DATE(a.shift_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "')", _
'                        "(DATE(a.shift_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') ")
'
'    End If
    
    If vFlagRollable = 0 Then
        SQL = "SELECT  b.dt tgl,c.att_date,b.dt day_name,a.employee_code," & _
                    "a.nik,a.employee_name,c.status,f.absent_name,a.title_code," & _
                    "d.title_name, c.time_in,c.time_out,c.entry_date,a.division_code," & _
                    "e.division_name,c.description, a.department_code, g.department_name," & _
                    "CASE WHEN c.status = 'P' THEN " & _
                        "CASE WHEN c.shift_code = 'S01' THEN 'M' " & _
                            "WHEN c.shift_code = 'S02' THEN 'A' " & _
                            "WHEN c.shift_code = 'S03' THEN 'N' " & _
                            "WHEN c.shift_code = 'ST01' THEN 'DAILY' " & _
                            "WHEN c.shift_code = 'OFF' THEN 'OFF' END ELSE f.absent_name END shift_code," & _
                    "h.shift_name, c.flag_inc_late, f.absent_name," & _
                    "CASE WHEN i.shift_code = 'S01' THEN 'M' " & _
                        "WHEN i.shift_code = 'S02' THEN 'A' " & _
                        "WHEN i.shift_code = 'S03' THEN 'N' " & _
                        "WHEN i.shift_code = 'ST01' THEN 'DAILY' " & _
                        "WHEN i.shift_code = 'OFF' THEN 'OFF' END act_shift, " & _
                    "c.shift_code shift, h.flag_day_over, c.enrollnumber, IFNULL(c.flag_training,0) flag_training " & _
              "FROM m_employee a JOIN m_days b ON 1 = 1 " & _
                    "LEFT JOIN h_attendance c ON DATE(b.dt) = DATE(c.att_date) AND a.employee_code = c.employee_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                    "LEFT JOIN m_division e ON a.division_code = e.division_code AND a.department_code = e.department_code AND a.company_code = e.company_code " & _
                    "LEFT JOIN m_absent_status f ON c.status = f.absent_code " & _
                    "LEFT JOIN m_department g ON a.department_code = g.department_code AND a.company_code = g.company_code " & _
                    "LEFT JOIN m_shift h ON c.shift_code = h.shift_code " & _
                    "LEFT JOIN td_shift2 i ON c.employee_code = i.employee_code "
    Else
        SQL = "SELECT  b.dt tgl,c.att_date,b.dt day_name,a.employee_code," & _
                    "a.nik,a.employee_name,c.status,f.absent_name,a.title_code," & _
                    "d.title_name, c.time_in,c.time_out,c.entry_date,a.division_code," & _
                    "e.division_name,c.description, a.department_code, g.department_name," & _
                    "CASE WHEN c.status = 'P' THEN " & _
                        "CASE WHEN c.shift_code = 'S01' THEN 'M' " & _
                            "WHEN c.shift_code = 'S02' THEN 'A' " & _
                            "WHEN c.shift_code = 'S03' THEN 'N' " & _
                            "WHEN c.shift_code = 'ST01' THEN 'DAILY' " & _
                            "WHEN c.shift_code = 'OFF' THEN 'OFF' END ELSE f.absent_name END shift_code," & _
                    "h.shift_name, c.flag_inc_late, f.absent_name," & _
                    "CASE WHEN (SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) = 'S01' THEN 'M' " & _
                        "WHEN (SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) = 'S02' THEN 'A' " & _
                        "WHEN (SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) = 'S03' THEN 'N' " & _
                        "WHEN (SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) = 'ST01' THEN 'DAILY' " & _
                        "WHEN (SELECT z.shift_code FROM td_emp_group x JOIN m_shift_new z ON x.emp_group_code = z.group_code WHERE employee_code = a.employee_code AND DATE(z.shift_date) = DATE(b.dt) AND DATE(x.start_date) <= DATE(b.dt) ORDER BY x.start_date DESC LIMIT 1) = 'OFF' THEN 'OFF' END act_shift, " & _
                    "c.shift_code shift, h.flag_day_over, c.enrollnumber, IFNULL(c.flag_training,0) flag_training " & _
              "FROM m_employee a JOIN m_days b ON 1 = 1 " & _
                    "LEFT JOIN h_attendance c ON DATE(b.dt) = DATE(c.att_date) AND a.employee_code = c.employee_code " & _
                    "LEFT JOIN m_title d ON a.title_code = d.title_code " & _
                    "LEFT JOIN m_division e ON a.division_code = e.division_code AND a.department_code = e.department_code AND a.company_code = e.company_code " & _
                    "LEFT JOIN m_absent_status f ON c.status = f.absent_code " & _
                    "LEFT JOIN m_department g ON a.department_code = g.department_code AND a.company_code = g.company_code " & _
                    "LEFT JOIN m_shift h ON c.shift_code = h.shift_code "
    
'        SQL = "SELECT a.shift_date tgl,d.att_date,a.shift_date day_name,b.employee_code,c.nik,c.employee_name," & _
'                "d.status,g.absent_name,c.title_code,e.title_name, d.time_in," & _
'                "d.time_out,d.entry_date,c.division_code,f.division_name,d.description,c.department_code, h.department_name, " & _
'                "CASE WHEN d.status = 'P' THEN " & _
'                        "CASE WHEN d.shift_code = 'S01' THEN 'M' " & _
'                            "WHEN d.shift_code = 'S02' THEN 'A' " & _
'                            "WHEN d.shift_code = 'S03' THEN 'N' " & _
'                            "WHEN d.shift_code = 'ST01' THEN 'DAILY' " & _
'                            "WHEN d.shift_code = 'OFF' THEN 'OFF' END ELSE g.absent_name END shift_code," & _
'                "i.shift_name,d.flag_inc_late, g.absent_name, " & _
'                "CASE WHEN a.shift_code = 'S01' THEN 'M' " & _
'                        "WHEN a.shift_code = 'S02' THEN 'A' " & _
'                        "WHEN a.shift_code = 'S03' THEN 'N' " & _
'                        "WHEN a.shift_code = 'ST01' THEN 'DAILY' " & _
'                        "WHEN a.shift_code = 'OFF' THEN 'OFF' END act_shift, " & _
'                "d.shift_code shift, i.flag_day_over " & _
'              "FROM m_shift_new a " & _
'                "JOIN (SELECT employee_code, emp_group_code FROM td_emp_group WHERE DATE(start_date) <= (SELECT DATE(shift_date) FROM m_shift_new ORDER BY emp_group_code DESC LIMIT 1)) b ON a.group_code = b.emp_group_code " & _
'                "LEFT JOIN m_employee c ON b.employee_code = c.employee_code " & _
'                "LEFT JOIN h_attendance d ON b.employee_code = d.employee_code AND DATE(d.att_date) = DATE(a.shift_date) " & _
'                "LEFT JOIN m_title e ON c.title_code = e.title_code " & _
'                "LEFT JOIN m_division f ON c.division_code = f.division_code AND c.department_code = f.department_code AND c.company_code = f.company_code " & _
'                "LEFT JOIN m_absent_status g ON d.status = g.absent_code " & _
'                "LEFT JOIN m_department h ON c.department_code = h.department_code AND c.company_code = h.company_code " & _
'                "LEFT JOIN m_shift i ON d.shift_code = i.shift_code "
    End If
                
    If vMode = 0 Then
        If vFlagRollable = 0 Then
            SQL = SQL & _
                    "where date(c.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                        "AND c.group_code = '" & TDBCombo_Group_Shift.Text & "' " & _
                        "AND c.shift_code = '" & TDBGrid_Shift.Columns("shift_code").Value & "' " & _
                    "order by c.att_date "
        Else
            SQL = SQL & _
                    "where date(c.att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
                        "AND c.group_code = '" & TDBCombo_Group_Shift.Text & "' " & _
                        "AND c.shift_code = '" & TDBGrid_Shift.Columns("shift_code").Value & "' " & _
                    "order by c.att_date "
        End If
    Else
        If vFlagRollable = 0 Then
            SQL = SQL & _
                    "where " & vParameter & " ORDER BY b.dt, c.shift_code"
        Else
            SQL = SQL & _
                    "where " & vParameter & " ORDER BY b.dt, c.shift_code"
        End If
    End If
    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Att.DataSource = rsAtt
    
    cmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)
    cmdEdit.Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    cmdDelete.Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    cmdRefresh.Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    
    Call load_data_user_access(Me)
    cmdNew.Enabled = blnUser_Add
    cmdEdit.Enabled = blnUser_Edit
    cmdDelete.Enabled = blnUser_Delete
End Sub

Private Sub load_data_log_att()
    If rsLogAtt.State Then rsLogAtt.Close
    SQL = "SELECT a.*, DATE(a.att_date) AS date_att, TIME(a.att_date) AS time_att," & _
             "CASE WHEN a.flag_io = 0 THEN 'IN' ELSE 'OUT' END TYPE " & _
          "FROM h_log_attendance a " & _
          "WHERE (a.enrollnumber = '" & Val(TDBGrid_Att.Columns("nik").Value) & "' " & _
                "OR a.employee_code = '" & Val(TDBGrid_Att.Columns("employee_code").Value) & "') " & _
            "AND CASE WHEN '" & Val(TDBGrid_Att.Columns("flag_day_over").Value) & "' = 1 OR WEEKDAY('" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "') = 5 OR WEEKDAY('" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "') = 6 THEN " & _
                    "(date(a.att_date) BETWEEN '" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "' AND ADDDATE('" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "',1)) " & _
                "ELSE date(a.att_date) = '" & Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd") & "' END " & _
          "ORDER BY a.att_date, a.flag_io"
    rsLogAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid1.DataSource = rsLogAtt
End Sub

Private Sub load_data_pl()
Dim vPLDate As String
    If rsPL.State Then rsPL.Close
    
    vPLDate = IIf(TDBGrid_Att.Columns("att_date").Value = "", _
                  Format(TDBGrid_Att.Columns("tgl").Value, "yyyy-MM-dd HH:mm:ss"), _
                  Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss"))
                  
    SQL = "SELECT pl_date, employee_code, time_in, flag_approval, time_out, " & _
            "CASE WHEN flag_type = 0 THEN 'PL' " & _
                "WHEN flag_type = 1 THEN 'SL' " & _
            "ELSE 'AL' END type, seq " & _
          "FROM t_private_leave " & _
          "WHERE employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
            "AND DATE(pl_date) = DATE('" & vPLDate & "')"
    rsPL.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid2.DataSource = rsPL
End Sub

Private Sub load_data_spl()
    If rsSPL.State Then rsSPL.Close
    SQL = "SELECT a.date, a.employee_code, a.start_time, a.flag_approval, a.end_time, seq, IFNULL(a.flag_change_status,0) flag_change_status " & _
          "FROM t_spl a JOIN m_employee b ON a.employee_code = b.employee_code " & _
          "WHERE a.employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
            "AND CASE WHEN b.flag_shiftable = 0 THEN date = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
                "ELSE DATE(date) = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd") & "' END"
    rsSPL.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBGrid3.DataSource = rsSPL
End Sub

Private Sub load_data_leave()
    If rsEmpLeave.State Then rsEmpLeave.Close
    SQL = "SELECT * FROM h_attendance a JOIN t_leave b ON a.employee_code = b.employee_code AND DATE(a.att_date) =  DATE(b.leave_date_from) " & _
            "WHERE a.employee_code = '" & TDBGrid_Att.Columns("employee_code").Value & "' " & _
                "AND att_date = '" & Format(TDBGrid_Att.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss") & "' " & _
                "AND status = 'L'"
    rsEmpLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid4.DataSource = rsEmpLeave
End Sub

Private Sub showData()
Dim s As String
Dim vFlagLate As Integer
Dim vTimeIn, vTimeOut As String
    
    SQL = "SELECT start_time, end_time FROM m_working_day WHERE group_code = '" & TDBCombo_Group_Shift.Text & "' " & _
            "AND shift_code = '" & TDBGrid_Shift.Columns("shift_code").Value & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vTimeIn = Format(rscari!start_time, "HH:mm")
        vTimeOut = Format(rscari!end_time, "HH:mm")
    Else
        vTimeIn = "00:00"
        vTimeOut = "00:00"
    End If
    rscari.Close
    
    If TDBGrid_Att.ApproxCount > 0 Then
        With frm_trans_att_man
            .DTPicker1.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
            .txt_employee_code.Text = IIf(IsNull(TDBGrid_Att.Columns("employee_code").Value), txt_employee_code.Text, TDBGrid_Att.Columns("employee_code").Value)
            .txt_nik.Text = IIf(IsNull(TDBGrid_Att.Columns("nik").Value), txt_nik.Text, TDBGrid_Att.Columns("nik").Value)
            .txt_employee_name.Text = IIf(IsNull(TDBGrid_Att.Columns("employee_name").Value), txt_employee_name.Text, TDBGrid_Att.Columns("employee_name").Value)
            
            .TDBCombo_shift.Text = IIf(TDBGrid_Att.Columns("shift").Value = "", TDBGrid_Shift.Columns("shift_code").Value, TDBGrid_Att.Columns("shift").Value)
            .txt_shift_name.Text = IIf(TDBGrid_Att.Columns("shift_name").Value = "", TDBGrid_Shift.Columns("shift_name").Value, TDBGrid_Att.Columns("shift_name").Value)
            
            .TDBCombo_department.Text = IIf(IsNull(TDBGrid_Att.Columns("department_code").Value), "", TDBGrid_Att.Columns("department_code").Value)
            .txt_department_name.Text = IIf(IsNull(TDBGrid_Att.Columns("department_name").Value), "", TDBGrid_Att.Columns("department_name").Value)
            .txt_division_code.Text = IIf(IsNull(TDBGrid_Att.Columns("division_code").Value), "", TDBGrid_Att.Columns("division_code").Value)
            .txt_division_name.Text = IIf(IsNull(TDBGrid_Att.Columns("division_name").Value), "", TDBGrid_Att.Columns("division_name").Value)
            .txt_title_code.Text = IIf(IsNull(TDBGrid_Att.Columns("title_code").Value), "", TDBGrid_Att.Columns("title_code").Value)
            .txt_title_name.Text = IIf(IsNull(TDBGrid_Att.Columns("title_name").Value), "", TDBGrid_Att.Columns("title_name").Value)
            
            .TDBCombo_Status.Text = IIf(TDBGrid_Att.Columns("status").Value = "", "P", TDBGrid_Att.Columns("status").Value)
            .txt_status_name.Text = IIf(TDBGrid_Att.Columns("absent_name").Value = "", "PRESENT", TDBGrid_Att.Columns("absent_name").Value)
            
            .DTPicker_from.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
            .DTPicker_to.Value = IIf(TDBGrid_Att.Columns("att_date").Value = "", TDBGrid_Att.Columns("tgl").Value, TDBGrid_Att.Columns("att_date").Value)
            
            If TDBGrid_Att.Columns("att_date").Value = "" Then
                .DTPicker_from.Enabled = True
                .DTPicker_to.Enabled = True
                
                .ttin.Value = vTimeIn
                .ttout.Value = vTimeOut
            Else
                .DTPicker_from.Enabled = False
                .DTPicker_to.Enabled = False
                
                If .TDBCombo_Status.Text = "P" Or .TDBCombo_Status.Text = "DT" Or .TDBCombo_Status.Text = "OT" Then
                    .ttin.Value = Format(TDBGrid_Att.Columns("time_in").Value, "hh:mm")
                    .ttout.Value = Format(TDBGrid_Att.Columns("time_out").Value, "hh:mm")
                Else
                    .ttin.Enabled = False
                    .ttout.Enabled = False
                    .ttin.Value = "00:00"
                    .ttout.Value = "00:00"
                End If
            End If
            
            

            .txt_description = IIf(IsNull(TDBGrid_Att.Columns("description").Value), "", TDBGrid_Att.Columns("description").Value)
            
            s = TDBGrid_Att.Columns("flag_inc_late").Value
            vFlagLate = IIf(s = "", 0, s)
            .chkLate.Value = vFlagLate
            
            .chk_all_employee.Enabled = False
            
            If TDBGrid_Att.Columns("att_date").Value = "" Then
                .editTrans = False
            Else
                .editTrans = True
            End If
            
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_att_man
        .DTPicker1.Value = DTPicker1.Value
        .DTPicker_from.Value = DTPicker1.Value
        .DTPicker_to.Value = DTPicker1.Value
        .TDBCombo_Status.Text = "P"
        .txt_status_name.Text = "PRESENT"
        .vGroupCode = TDBCombo_Group_Shift
        .vShiftCode = TDBGrid_Shift.Columns("shift_code").Value
        
        .TDBCombo_shift.Text = TDBGrid_Shift.Columns("shift_code").Value
        .txt_shift_name.Text = TDBGrid_Shift.Columns("shift_name").Value
        
        .chkSat.Value = 0
        .chkSun.Value = 0
        .chkLate.Value = 0
        
        .editTrans = False
        .Show vbModal
    End With
End Sub

'Private Sub Form_Resize()
'    Frame1.Width = Me.Width - 500
'    Frame1.Height = Me.Height - 2500
'    TDBGrid_Shift.Height = Frame1.Height - 400
'    TDBGrid_Shift.Width = Frame1.Width - (TDBGrid_Att.Width + 400)
'    TDBGrid_Att.Height = Frame1.Height - 400
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_list_manual_att = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        LynxGrid1.ClearAll
        createGrid
    ElseIf SSTab1.Tab = 2 Then
        LynxGrid2.ClearAll
        createGridM
    End If
End Sub

Private Sub tdbGrid_Att_DblClick()
    fraByPass.Visible = True
    TDBGrid_Att.Enabled = False
    
    vFlagTraining = TDBGrid_Att.Columns("flag_training").Value
    If vFlagTraining = 0 Then
        cmdTraining.Caption = "&Training"
    Else
        cmdTraining.Caption = "Un&training"
    End If
'    showData
End Sub

Private Sub TDBGrid1_DblClick()
Dim vTglLog As String
    fraInOut.Visible = True
    TDBGrid1.Enabled = False
    
    vTglLog = Format(TDBGrid1.Columns("att_date").Value, "yyyy-MM-dd HH:mm:ss")
    
    SQL = "SELECT flag_io FROM h_log_attendance " & _
          "WHERE enrollnumber = '" & Val(TDBGrid_Att.Columns("nik").Value) & "' " & _
            "AND att_date = '" & vTglLog & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        optIn.Value = IIf(rscari!flag_io = 0, 1, 0)
        optOut.Value = IIf(rscari!flag_io = 0, 0, 1)
    End If
    rscari.Close
End Sub

Private Sub TDBGrid2_DblClick()
    fraPL.Visible = True
    TDBGrid2.Enabled = False
End Sub

Private Sub TDBGrid3_DblClick()
    If TDBGrid_Att.Columns("status").Value <> "OT" Then
        fraOT.Visible = True
        TDBGrid3.Enabled = False
    End If
End Sub

Private Sub TDBGrid_Att_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim vFlagRoll As Integer
Dim i As Integer
    TDBGrid_Att.Bookmark = Bookmark
    
    SQL = "SELECT flag_rollable FROM m_shift_group WHERE group_code = '" & TDBCombo_Group_Shift.Text & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vFlagRoll = rscari!flag_rollable
    End If
    rscari.Close
    
    If vFlagRoll = 0 Then
        If Format(TDBGrid_Att.Columns("day_name").Value, "DDDD") = "Saturday" Then RowStyle.BackColor = vbGreen
        If Format(TDBGrid_Att.Columns("day_name").Value, "DDDD") = "Sunday" Then RowStyle.BackColor = vbGreen
    Else
        If TDBGrid_Att.Columns("act_shift").Value = "OFF" Then RowStyle.BackColor = vbGreen
    End If
End Sub

Private Sub TDBGrid_Att_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        showData
    End If
End Sub

Private Sub TDBGrid_Att_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not IIf(IsNull(TDBGrid_Att.Bookmark), 0, TDBGrid_Att.Bookmark) > 0 Then
        Set TDBGrid1.DataSource = Nothing
        Set TDBGrid2.DataSource = Nothing
        Set TDBGrid3.DataSource = Nothing
        Set TDBGrid4.DataSource = Nothing
        Exit Sub
    End If
    
    Call load_data_log_att
    Call load_data_spl
    Call load_data_pl
    Call load_data_leave
End Sub

Private Sub TDBGrid_Shift_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call load_data_att
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value

        Call load_data_group_shift
    End If
End Sub

Private Sub TDBCombo_Group_Shift_ItemChange()
    If TDBCombo_Group_Shift.ApproxCount > 0 Then
        TDBCombo_Group_Shift.Text = TDBCombo_Group_Shift.Columns("group_code").Value
        txt_group_shift.Text = TDBCombo_Group_Shift.Columns("group_name").Value

        Call load_data_shift
        Call load_data_att
    End If
End Sub

Private Sub CmdNew_Click()
    If TDBGrid_Shift.ApproxCount = 0 Then
        MsgBox "Invalid Shift Code!" & Chr(13) & "Please Check Your Transaction Again...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    newData
End Sub

Private Sub cmdEdit_Click()
    showData
End Sub

Private Sub cmdRefresh_Click()
    vMode = 0
    Call load_data_att
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid_Att.Columns
        Col.FilterText = ""
    Next Col
    rsAtt.Filter = adFilterNone
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

Private Sub TDBGrid_Att_FilterChange()
On Error GoTo Err

    Dim i As Integer
    
    Set Cols = TDBGrid_Att.Columns
    i = TDBGrid_Att.Col
    TDBGrid_Att.HoldFields
    
    rsAtt.Filter = getFilter()
    TDBGrid_Att.Col = i
    TDBGrid_Att.EditActive = True
    
    TDBGrid_Att.SelStart = Len(TDBGrid_Att.Columns(i).FilterText)
    If TDBGrid_Att.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid_Att.Col = i
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim item
    
On Error GoTo Err
    If Not TDBGrid_Att.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid_Att.SelBookmarks
    i = MsgBox("Are you sure want to delete " _
        & SelBks.Count & " attendance's data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        SQL = "DELETE FROM h_attendance where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
            & "and att_date = '" & Format(TDBGrid_Att.Columns("att_date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM t_private_leave where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
                & "AND DATE(pl_date) = '" & Format(TDBGrid_Att.Columns("att_date").CellText(item), "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM t_spl where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
                & "AND DATE(date) = '" & Format(TDBGrid_Att.Columns("att_date").CellText(item), "yyyy-MM-dd") & "'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM t_leave where employee_code = '" & TDBGrid_Att.Columns("employee_code").CellText(item) & "' " _
                & "AND DATE(leave_date_from) = '" & Format(TDBGrid_Att.Columns("att_date").CellText(item), "yyyy-MM-dd") & "'"
        CnG.Execute SQL
                
    Next
    CnG.CommitTrans
    Call load_data_att
    MsgBox i & " attendance's data are successfully deleted", vbInformation, headerMSG
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "ATT. DATE", 1200, lgAlignCenterCenter, lgDate, "yyyy-MM-dd", , , , , , True
      .AddColumn "EMP. CODE", 1300, lgAlignCenterCenter, , , , , , , True
      .AddColumn "EMP. NAME", 2500, , , , , , , , True
      .AddColumn "GROUP CODE", 800, lgAlignCenterCenter, , , , , , , , True
      .AddColumn "SHIFT CODE", 800, , , , , , , , , True
      .AddColumn "ATT. STATUS", 800, lgAlignCenterCenter, , , , , , , True
      .AddColumn "TIME IN", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "TIME OUT", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "DESCRIPTION", 3000, lgAlignCenterCenter, , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub createGridM()
   With LynxGrid2
      .AddColumn "EMP. CODE", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "ATT. DATE", 2000, lgAlignCenterCenter, lgDate, "yyyy-MM-dd hh:mm:ss", , , , , , True
      .AddColumn "CHECK TYPE", 1000, lgAlignCenterCenter, , , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Call fill_grid_excel(CommonDialog1.FileName)
    End If
End Sub

Private Sub cmdBrowseM_Click()
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Call fill_grid_excelM(CommonDialog1.FileName)
    End If
End Sub

Private Sub fill_grid_excel(str_file_name As String)
On Error GoTo Err
    Dim strWorksheet As String
    strWorksheet = "data_attendance"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$] order by 1,2 ASC"
    Adodc1.Refresh
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                LynxGrid1.AddItem .Fields(0) & vbTab & .Fields(1) _
                        & vbTab & .Fields(2) & vbTab & .Fields(3) _
                        & vbTab & .Fields(4) & vbTab & .Fields(5) _
                        & vbTab & .Fields(6) & vbTab & .Fields(7) _
                        & vbTab & .Fields(8)
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid1.Redraw = True
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, "Message Error!"
End Sub

Private Sub fill_grid_excelM(str_file_name As String)
On Error GoTo Err
    Dim strWorksheet As String
    strWorksheet = "data_log"
    
    Adodc2.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc2.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc2.Refresh
    LynxGrid2.Redraw = False
    LynxGrid2.Clear
    With Adodc2.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                LynxGrid2.AddItem .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2)
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid2.Redraw = True
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, "Message Error!"
End Sub

Private Sub cmdImport_Click()
Dim aa As Integer
Dim rsnumber As New ADODB.Recordset
Dim nourut As Long
Dim v_employee_code As String
Dim SQL As String
Dim waktu_in As String, waktu_out As String
Dim time_in As String, time_out As String
Dim start_time As String, end_time As String

'On Error Resume Next
    With LynxGrid1
        If .Rows > 0 Then
            ProgressBar1.Visible = True
            Label5.Visible = True
            ProgressBar1.Max = .Rows - 1
            ProgressBar1.Value = 0
            
            DoEvents
            
            For aa = 0 To .Rows - 1
                
                ProgressBar1.Value = aa
                Label5.Caption = .CellText(aa, 0) & " - " & .CellText(aa, 1)
                             
                SQL = "SELECT employee_code FROM m_employee WHERE nik = '" & .CellText(aa, 1) & "' " _
                            & "AND flag_active <> 0"
                rsnumber.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsnumber.RecordCount > 0 Then
                    v_employee_code = rsnumber!employee_code
                    
                    SQL = "DELETE FROM h_attendance " & _
                          "WHERE att_date = '" & .CellText(aa, 0) & "' " & _
                            "AND employee_code = '" & .CellText(aa, 1) & "'"
                    CnG.Execute SQL
                    
                    SQL = "SELECT start_time, end_time FROM m_shift " & _
                      "WHERE group_code = '" & .CellText(aa, 3) & "' " & _
                        "AND shift_code = '" & .CellText(aa, 4) & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        waktu_in = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(rs!start_time, "hh:mm") & ":00"
                        waktu_out = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(rs!end_time, "hh:mm") & ":00"
                    Else
                        waktu_in = Format(.CellText(aa, 0), "yyyy-MM-dd") & " 07:00:00"
                        waktu_out = Format(.CellText(aa, 0), "yyyy-MM-dd") & " 15:00:00"
                    End If
                    rs.Close
                    
                    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
                    SQL = "Select CAST(concat('" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
                            & "CAST(concat('" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
                            & "now() tglserver " _
                            & "from m_shift where shift_code = '" & .CellText(aa, 4) & "' " & _
                                "AND group_code = '" & .CellText(aa, 3) & "'"
                '        & "from m_shift where shift_code = '" & txtkdshift.Text & "'"
                    rs.Open SQL, CnG, adOpenDynamic, adLockReadOnly
                    
                    start_time = Format(rs!start_time, "yyyy-MM-dd hh:mm:ss")
                    If frm_list_manual_att.TDBGrid_Shift.Columns("flag_day_over").Value = 1 Then
                        end_time = Format(DateAdd("d", 1, rs!end_time), "yyyy-MM-dd hh:mm:ss")
                    Else
                        end_time = Format(rs!end_time, "yyyy-MM-dd hh:mm:ss")
                    End If
                    
                    time_in = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(.CellText(aa, 6), "hh:mm:ss")
                    If Format(.CellText(aa, 7), "hh:mm:ss") <= Format(.CellText(aa, 6), "hh:mm:ss") Then
                        time_out = Format(DateAdd("d", 1, .CellText(aa, 0)), "yyyy-MM-dd") & " " & Format(.CellText(aa, 7), "hh:mm") & ":00"
                    Else
                        time_out = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(.CellText(aa, 7), "hh:mm") & ":00"
                    End If
                    
                    rs.Close
    
                    If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                        Select Case .CellText(aa, 5)
                        Case "P"
                            SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & v_employee_code & "'"
                            CnG.Execute SQL
                            
                            SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                                      "shift_number,start_time,end_time,time_in,time_out,description,entry_date,userinput," & _
                                      "absent_status,flag_present,flag_duty) " & _
                                    "VALUES (" & _
                                      "'" & v_employee_code & "','" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "','" & .CellText(aa, 3) & "'," & _
                                      "'" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "',1,'" & start_time & "','" & end_time & "'," & _
                                      "'" & time_in & "','" & time_out & "'," & _
                                      "'" & .CellText(aa, 8) & "',now(),'" & LOGIN_NAME & "',0,1,0)"
                            CnG.Execute SQL
                        Case "A", "OF", "PR", "S"
                            If .CellText(aa, 5) = "A" Then
                                abstatus = 2
                            ElseIf .CellText(aa, 5) = "OF" Then
                                abstatus = 6
                            ElseIf .CellText(aa, 5) = "PR" Then
                                abstatus = 0
                            ElseIf .CellText(aa, 5) = "S" Then
                                abstatus = 1
                            End If
                            
                            SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & v_employee_code & "'"
                            CnG.Execute SQL
                            
                            SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,flag_present,absent_status," & _
                                    "description,entry_date,userinput,shift_number) " & _
                                  "VALUES (" & _
                                    "'" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "','" & v_employee_code & "','" & .CellText(aa, 3) & "'," & _
                                    "'" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "',0,'" & abstatus & "'," & _
                                    "'" & .CellText(aa, 8) & "',now(),'" & LOGIN_NAME & "',1)"
                            CnG.Execute SQL
                        Case "DT"
                            SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & v_employee_code & "'"
                            CnG.Execute SQL
                                
                            SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,time_in,time_out," & _
                                    "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number) " & _
                                  "VALUES (" & _
                                    "'" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "','" & v_employee_code & "','" & .CellText(aa, 3) & "'," & _
                                    "'" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "'," & _
                                    "'" & waktu_in & "','" & waktu_out & "'," & _
                                    "1,1,0,'" & .CellText(aa, 8) & "',now(),'" & LOGIN_NAME & "',1)"
                            CnG.Execute SQL
                        End Select
                    End If
                End If
                rsnumber.Close
                
                
                DoEvents
            Next
            MsgBox "Import Data Success...!!!", vbInformation, headerMSG
            
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            ProgressBar1.Visible = False
            Label5.Visible = False
        End If
    End With
End Sub

Private Sub cmdImportM_Click()
Dim aa As Integer
Dim rsnumber As New ADODB.Recordset
Dim nourut As Long
Dim v_employee_code As String
Dim SQL As String

'On Error Resume Next
    With LynxGrid2
        If .Rows > 0 Then
            ProgressBar2.Visible = True
            Label6.Visible = True
            ProgressBar2.Max = .Rows - 1
            ProgressBar2.Value = 0
            
            DoEvents
            
            For aa = 0 To .Rows - 1
                
                ProgressBar2.Value = aa
                Label6.Caption = .CellText(aa, 0) & " - " & .CellText(aa, 1)
                             
                SQL = "SELECT a.employee_code FROM m_employee a JOIN m_enroll_link b ON a.employee_code = b.employee_code " & _
                      "WHERE b.enrollnumber = '" & .CellText(aa, 0) & "' AND a.flag_active <> 0"
                rsnumber.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsnumber.RecordCount > 0 Then
                    v_employee_code = rsnumber!employee_code
                    
                    SQL = "DELETE FROM h_log_attendance WHERE enrollnumber = '" & Val(.CellText(aa, 0)) & "' " & _
                            "AND att_date = '" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "'"
                    CnG.Execute SQL
                    
                    SQL = "DELETE FROM h_attendance WHERE enrollnumber = '" & Val(.CellText(aa, 0)) & "' " & _
                            "AND att_date = '" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "'"
                    CnG.Execute SQL
                    
                    SQL = "SELECT * FROM h_log_attendance " & _
                          "WHERE enrollnumber = '" & Val(.CellText(aa, 0)) & "' " & _
                            "AND att_date = '" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "'"
                    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rscari.RecordCount = 0 Then
                        SQL = "INSERT INTO h_log_attendance(att_date,enrollnumber,employee_code,verifymode,flag_io) " & _
                                "VALUES ( " & _
                                "'" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "','" & Val(.CellText(aa, 0)) & "','" & v_employee_code & "',0,'" & Val(.CellText(aa, 2)) & "')"
                        CnG.Execute SQL
                    End If
                    rscari.Close
                End If
                rsnumber.Close
                
                
                DoEvents
            Next
            MsgBox "Import Data Success...!!!", vbInformation, headerMSG
            
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            ProgressBar2.Visible = False
            Label6.Visible = False
        End If
    End With
End Sub

'Private Sub TDBGrid_Att_HeadClick(ByVal ColIndex As Integer)
'
'    x = x + 1
'
'    If x Mod 2 <> 1 And vSubject = TDBGrid_Att.Columns(ColIndex).DataField Then
'        oClause = " ORDER BY " + TDBGrid_Att.Columns(ColIndex).DataField + " DESC"
'    Else
'        oClause = " ORDER BY " + TDBGrid_Att.Columns(ColIndex).DataField + " ASC"
'    End If
'
'    vSubject = TDBGrid_Att.Columns(ColIndex).DataField
'    Call load_data_att
'
'End Sub

Private Sub createGridKar()
   With LynxGrid3
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid3.Clear
        
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
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid3.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid3.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid3.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
                txt_nik.Text = rs!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid3.Visible = True
                LynxGrid3.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid3.Rows > 0 Then
            txt_nik.Text = LynxGrid3.CellText(LynxGrid3.Row, 0)
            txt_employee_name.Text = LynxGrid3.CellText(LynxGrid3.Row, 1)
            txt_employee_code.Text = LynxGrid3.CellText(LynxGrid3.Row, 2)
        End If
        LynxGrid3.Visible = False
    End If
    
    lbl_l.Caption = 0
    lbl_pl.Caption = 0
    lbl_sick.Caption = 0
    lbl_late.Caption = 0
    
    lbl_meal.Caption = 0
    lbl_trans.Caption = 0
    lbl_a.Caption = 0
    lbl_n.Caption = 0
    
    lbl_hk_15.Caption = 0
    lbl_hk_20.Caption = 0
    
    lbl_hl_20.Caption = 0
    lbl_hl_30.Caption = 0
    lbl_hl_40.Caption = 0
End Sub

Private Sub LynxGrid3_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid3.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid3_LostFocus()
    LynxGrid3.Visible = False
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

Private Sub cmdBrowse_Emp_Click()
    isiGridKar (1)
End Sub

Private Sub days_func(start_time As String, end_time As String)
Dim v_tgl_awal, v_tgl_akhir As Date
On Error Resume Next

    v_tgl_awal = Format(start_time, "yyyy-MM-dd")
    v_tgl_akhir = Format(end_time, "yyyy-MM-dd")
    
    v_tgl_awal = DateValue(v_tgl_awal)
    v_tgl_akhir = DateValue(v_tgl_akhir)
    
    SQL = "delete from m_days where dt between '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' and '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "'"
    CnG.Execute SQL
            
        While v_tgl_awal <= v_tgl_akhir
            SQL = "SELECT holiday_date FROM t_holiday WHERE date(holiday_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                SQL = "INSERT INTO m_days (dt,status,description) " & _
                      "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','L','HOLIDAY')"
                CnG.Execute SQL
            Else
                If Format(v_tgl_awal, "dddd") = "Sunday" Then
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','M','SUNDAY')"
                    CnG.Execute SQL
                ElseIf Format(v_tgl_awal, "dddd") = "Saturday" Then
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','S','SATURDAY')"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','W','WORK DAY')"
                    CnG.Execute SQL
                End If
            End If
            rs.Close
            
            
            v_tgl_awal = v_tgl_awal + 1
        Wend
End Sub

'Private Sub cari_jam()
'Dim rs2 As New ADODB.Recordset
'Dim rs3 As New ADODB.Recordset
'Dim rsJam As New ADODB.Recordset
'
'Dim vStatus As String
'Dim v_shift As String
'Dim vFlagHoliday As Integer
'Dim vFlagShiftable
'
'    With frm_trans_spl
'        SQL = "SELECT IFNULL(flag_shiftable,0) flag_shiftable FROM m_employee " & _
'              "WHERE employee_code = '" & .txt_employee_code.Text & "'"
'        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rscari.RecordCount > 0 Then
'            vFlagShiftable = rscari!flag_shiftable
'        End If
'        rscari.Close
'
'        If vFlagShiftable = 0 Then
'            SQL = "SELECT att_date,status FROM h_attendance " & _
'                  "WHERE Date(att_date) = '" & Format(.DTPicker1.Value, "yyyy-MM-dd") & "' " & _
'                    "AND employee_code = '" & .txt_employee_code.Text & "' " & _
'                    "AND (WEEKDAY('" & Format(.DTPicker1.Value, "yyyy-MM-dd") & "') = 5 OR WEEKDAY('" & Format(.DTPicker1.Value, "yyyy-MM-dd") & "') = 6)"
'        Else
'            SQL = "SELECT att_date,status FROM h_attendance " & _
'                  "WHERE Date(att_date) = '" & Format(.DTPicker1.Value, "yyyy-MM-dd") & "' " & _
'                    "AND employee_code = '" & .txt_employee_code.Text & "' " & _
'                    "AND status = 'OF'"
'        End If
'        rsJam.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rsJam.RecordCount > 0 Then
'
'            SQL = "SELECT time_in, time_out FROM h_attendance " & _
'                  "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
'                    "AND employee_code = '" & txt_employee_code.Text & "'"
'            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'            If rs2.RecordCount > 0 Then
'                .TDBTime1.Text = Format(rs2!time_in, "hh:mm")
'                .TDBTime2.Text = Format(rs2!time_out, "hh:mm")
'            Else
'                .TDBTime1.Text = Format(Now, "hh:mm")
'                .TDBTime2.Text = Format(Now, "hh:mm")
'
'                .txt15.Text = 0
'                .txt2.Text = 0
'                .txt3.Text = 0
'                .txt4.Text = 0
'                .txt6.Text = 0
'            End If
'            rs2.Close
'        Else
'            SQL = "SELECT time_in, end_time, time_out, status FROM h_attendance " & _
'                  "WHERE Date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
'                    "AND employee_code = '" & txt_employee_code.Text & "'"
'            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'            If rs2.RecordCount > 0 Then
'
'                If Format(DTPicker1.Value, "dddd") = "Sunday" Then
'                    .TDBTime1.Text = Format(rs2!time_in, "hh:mm")
'                    .TDBTime2.Text = Format(rs2!time_out, "hh:mm")
'                Else
'                    .TDBTime1.Text = Format(rs2!end_time, "hh:mm")
'                    .TDBTime2.Text = Format(rs2!time_out, "hh:mm")
'                End If
'            Else
'                .TDBTime1.Text = Format(Now, "hh:mm")
'                .TDBTime2.Text = Format(Now, "hh:mm")
'
'                .txt15.Text = 0
'                .txt2.Text = 0
'                .txt3.Text = 0
'                .txt4.Text = 0
'                .txt6.Text = 0
'            End If
'            rs2.Close
'        End If
'
'        SQL = "SELECT a.holiday_date, a.ot_code, b.ot_name FROM t_holiday a join m_ot b on a.ot_code = b.ot_code " & _
'              "WHERE Date(holiday_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
'        rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'        If rs2.RecordCount > 0 Then
'            TDBCombo_ot.Text = rs2!ot_code
'            txt_ot_name.Text = rs2!ot_name
'        Else
'            TDBCombo_ot.Text = ""
'            txt_ot_name.Text = ""
'        End If
'        rs2.Close
'
'        rsJam.Close
'    End With
'End Sub

Private Sub cmd_reproccess_Click()
Dim strsql As String
On Error Resume Next

    If txt_employee_code = "" Then
        MsgBox "Employee is empty...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    '+++++++++++++++++++++++++++++ Reproccess Attendance ++++++++++++++++++++++++++++++++
    SQL = "DELETE FROM h_attendance_reproccess"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_attendance_reproccess " & _
          "SELECT * FROM h_attendance " & _
          "WHERE IFNULL(flag_manual,0) = 1 " & _
            "AND (DATE(att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " & _
                "AND '" & Format(DateAdd("d", 1, DTPicker_to.Value), "yyyy-MM-dd") & "') " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
    CnG.Execute SQL

    SQL = "DELETE FROM h_log_attendance_recover"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_log_attendance_reproccess"
    CnG.Execute SQL
    
    SQL = "SELECT enrollnumber FROM m_enroll_link WHERE employee_code = '" & txt_employee_code.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            SQL = "INSERT INTO h_log_attendance_recover(att_date, ip_address, enrollnumber, employee_code, verifymode, " _
                            & "flag_io, flag_attendance, ref_date, entry_date) " _
                    & "SELECT att_date, ip_address, enrollnumber, employee_code, verifymode, flag_io, " _
                            & "flag_attendance , ref_date, entry_date " _
                    & "FROM h_log_attendance " _
                    & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " _
                        & "AND '" & Format(DateAdd("d", 1, DTPicker_to.Value), "yyyy-MM-dd") & "') " _
                        & "AND enrollnumber = '" & Val(rs!enrollnumber) & "' " _
                    & "ORDER BY att_date, flag_io"
            CnG.Execute SQL
        rs.MoveNext
        Wend
    End If
    rs.Close
    
    SQL = "DELETE FROM h_attendance " _
            & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " _
                & "AND '" & Format(DateAdd("d", 1, DTPicker_to.Value), "yyyy-MM-dd") & "') " _
                & "AND employee_code = '" & txt_employee_code.Text & "'"
    CnG.Execute SQL
    
'    SQL = "DELETE FROM h_log_attendance " _
'            & "WHERE (date(att_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' " _
'                & "AND '" & Format(DateAdd("d", 1, DTPicker_to.Value), "yyyy-MM-dd") & "') " _
'                & "AND enrollnumber = '" & Val(txt_nik.Text) & "'"
'    CnG.Execute SQL
    
    If rs.State Then rs.Close
    
    SQL = "SELECT * from h_log_attendance_recover ORDER BY att_date, flag_io"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    ProgressBar3.Visible = True
    ProgressBar3.Value = 0
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            ProgressBar3.Max = rs.RecordCount
            ProgressBar3.Value = ProgressBar3.Value + 1
            
            SQL = "INSERT INTO h_log_attendance_reproccess(att_date, ip_address, enrollnumber, employee_code, verifymode, " _
                            & "flag_io, flag_attendance, ref_date, entry_date) VALUES (" _
                    & "'" & Format(rs!att_date, "yyyy-MM-dd hh:nn:ss") & "', '" & rs!ip_address & "', '" & IIf(IsNull(rs!enrollnumber), 0, rs!enrollnumber) & "', " _
                    & "'" & rs!employee_code & "', '" & IIf(IsNull(rs!verifymode), 0, rs!verifymode) & "', '" & rs!flag_io & "', " _
                    & "'" & IIf(IsNull(rs!flag_attendance), 0, rs!flag_attendance) & "' , '" & Format(IIf(IsNull(rs!ref_date), 0, rs!ref_date), "yyyy-MM-dd hh:nn:ss") & "', '" & Format(IIf(IsNull(rs!entry_date), 0, rs!entry_date), "yyyy-MM-dd hh:nn:ss") & "')"
            CnG.Execute SQL
        rs.MoveNext
        Wend
    End If
    rs.Close
    
    SQL = "DELETE a FROM h_attendance a JOIN h_attendance_reproccess b ON a.employee_code = b.employee_code AND a.att_date = b.att_date"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_attendance " & _
          "SELECT * FROM h_attendance_reproccess"
    CnG.Execute SQL
    
    Screen.MousePointer = vbDefault
    
    ProgressBar3.Visible = False
    '++++++++++++++++++++++++++++++++ END Of Reproccess Attendance ++++++++++++++++++++++++
    
    '++++++++++++++++++++++++++++++++ Proses Salary +++++++++++++++++++++++++++++++++++++++
    Dim d1 As String
    Dim rsEmployee As New ADODB.Recordset
    Dim bulan As String
    Dim tgl As String
    Dim v_tgl_akhir As Date
    Dim v_tgl_mc As Date
    Dim v_end_mc As Date
    Dim int_month As Integer
    Dim int_year As Integer
    Dim bln_awal, bln_akhir, thn_awal, thn_akhir As String
    Dim v_pph21_type As String
    Dim v_jstk_type As String
    
    d1 = Format(DTPicker_to.Value, "yyyy-MM-01")
    
    SQL = "select LAST_DAY('" & d1 & "') tgl_akhir, a.employee_code, a.employee_name," & _
            "a.no_jamsostek,a.npwp,a.end_working,a.flag_active,a.nik," & _
            "(SELECT pph21_type FROM m_salary_standard WHERE employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) pph21_type," & _
            "(SELECT jstk_type FROM m_salary_standard WHERE employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) jstk_type " & _
            "from m_employee a " & _
            "JOIN (SELECT employee_code FROM h_attendance WHERE DATE(att_date) BETWEEN '" & Format(DTPicker_from, "yyyy-MM-dd") & "'  AND '" & Format(DTPicker_to, "yyyy-MM-dd") & "' " & _
                 "GROUP BY employee_code) d ON a.employee_code = d.employee_code " & _
            "where a.company_code = '" & TDBCombo_company.Text & "' " & _
                "AND a.employee_code = '" & txt_employee_code.Text & "'"
    rsEmployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    If rsEmployee.RecordCount > 0 Then
        v_tgl_akhir = rsEmployee!tgl_akhir
        v_pph21_type = IIf(IsNull(rsEmployee!pph21_type), "STD", rsEmployee!pph21_type)
        v_jstk_type = IIf(IsNull(rsEmployee!jstk_type), "STD", rsEmployee!jstk_type)
    End If
    
    If rsEmployee.RecordCount > 0 Then
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
                        
            str1 = "DELETE a FROM h_salary a JOIN m_employee b on a.employee_code = b.employee_code " & _
                "WHERE LEFT(a.month,7) = '" & Format(d1, "yyyy-MM") & "' " & _
                    "AND a.company_code = '" & TDBCombo_company.Text & "' " & _
                    "AND a.employee_code = '" & rsEmployee!employee_code & "'"
            CnG.Execute str1
            
            str1 = "DELETE FROM t_spl_auto " & _
                   "WHERE Date(date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-mm-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-mm-dd") & "' " & _
                        "AND employee_code = '" & rsEmployee!employee_code & "'"
            CnG.Execute str1
            
            rsEmployee.MoveNext
        Wend
    End If
    
    If rsEmployee.RecordCount > 0 Then
        DoEvents
    
        ProgressBar3.Max = rsEmployee.RecordCount
        ProgressBar3.Value = 0
                
        ProgressBar3.Visible = True
        
        '------------------------------- Hitung Overtime -----------------------------
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
            Call frm_trans_salary_process.auto_overtime(rsEmployee!employee_code, DTPicker_from.Value, DTPicker_to.Value)
            
            rsEmployee.MoveNext
            DoEvents
        Wend
        '-----------------------------------------------------------------------------
        
        If flagPLAuto() <> 0 Then
            '--------------------------------- Hitung Late -------------------------------
            rsEmployee.MoveFirst
            While Not rsEmployee.EOF
                Call check_late(rsEmployee!employee_code, DTPicker_from.Value, DTPicker_to.Value)
                
                rsEmployee.MoveNext
                DoEvents
            Wend
            '-----------------------------------------------------------------------------
        End If
        
'        '--------------------------------- Hitung PL -------------------------------
'        rsEmployee.MoveFirst
'        While Not rsEmployee.EOF
'            Call check_pl(rsEmployee!employee_code, DTPicker_from.Value, DTPicker_to.Value)
'
'            rsEmployee.MoveNext
'            DoEvents
'        Wend
'        '-----------------------------------------------------------------------------
        
        '------------------------------- Hitung Salary -------------------------------
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
            ProgressBar3.Value = ProgressBar3.Value + 1
            
            int_month = month(v_tgl_akhir)
            int_year = year(v_tgl_akhir)
                        
            Call frm_trans_salary_process.proses_su(rsEmployee!employee_code, _
                Format(DTPicker_from.Value, "yyyy-MM-dd"), Format(DTPicker_to.Value, "yyyy-MM-dd"), _
                IIf(IsNull(rsEmployee!no_jamsostek), "", rsEmployee!no_jamsostek), _
                IIf(IsNull(rsEmployee!npwp), 0, rsEmployee!npwp), TDBCombo_company.Text, _
                IIf(IsNull(Format(rsEmployee!end_working, "yyyyMM")), "0", Format(rsEmployee!end_working, "yyyyMM")), _
                rsEmployee!flag_active, Format(DTPicker_to.Value, "yyyy-MM"), v_pph21_type, v_jstk_type)

            '++++++++++++++++++++++++++++++ Update Data Loan +++++++++++++++++++++++++++
            
            SQL = "UPDATE td_loan SET flag_paid = 1 " _
                & "Where employee_code = '" & rsEmployee!employee_code & "' " _
                & "AND Month(installment_month) = '" & month(DTPicker_to.Value) & "' " _
                & "AND Year(installment_month) = '" & year(DTPicker_to.Value) & "'"
            CnG.Execute (SQL)
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
            rsEmployee.MoveNext
            DoEvents
        Wend
        '-------------------------------------------------------------------------------
        
        SQL = "UPDATE h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                "Set a.salary_value = 0 " & _
                "WHERE a.month = LEFT('" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "', 7) AND a.salary_code = 'SU-052'"
        CnG.Execute SQL
                
        'Update Temp Salary Proses ++++++++++++++++++++++++++++++++
        SQL = "UPDATE temp_sal_proses set salary_proses = 1 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute (SQL)
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    End If
    
    ProgressBar3.Visible = False
    '++++++++++++++++++++++++++++++ END Proses Salary +++++++++++++++++++++++++++++++++++++
    
    MsgBox "Proccess Successfully!", vbInformation, headerMSG
End Sub

Private Sub load_sum_att()
Dim rsSalary As New ADODB.Recordset
    
    If rsSalary.State Then rsSalary.Close
    SQL = "CALL spr_sum_detail_att('" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "'," & _
            "'" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "'," & _
            "'" & txt_employee_code.Text & "')"
    rsSalary.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsSalary.RecordCount > 0 Then
        With rsSalary
            lbl_l.Caption = .Fields("days_leave").Value
            lbl_pl.Caption = .Fields("private_leave").Value
            lbl_sick.Caption = .Fields("sick_leave").Value
            lbl_late.Caption = .Fields("sum_late").Value
            
            lbl_meal.Caption = .Fields("meal_days").Value
            lbl_trans.Caption = .Fields("transport_days").Value
            lbl_a.Caption = .Fields("afternoon_days").Value
            lbl_n.Caption = .Fields("night_days").Value
            
            lbl_hk_15.Caption = .Fields("ot15").Value
            lbl_hk_20.Caption = .Fields("ot20").Value
            
            lbl_hl_20.Caption = .Fields("ot20_hol").Value
            lbl_hl_30.Caption = .Fields("ot30_hol").Value
            lbl_hl_40.Caption = .Fields("ot40_hol").Value
        End With
    Else
        lbl_l.Caption = 0
        lbl_pl.Caption = 0
        lbl_sick.Caption = 0
        lbl_late.Caption = 0
        
        lbl_meal.Caption = 0
        lbl_trans.Caption = 0
        lbl_a.Caption = 0
        lbl_n.Caption = 0
        
        lbl_hk_15.Caption = 0
        lbl_hk_20.Caption = 0
        
        lbl_hl_20.Caption = 0
        lbl_hl_30.Caption = 0
        lbl_hl_40.Caption = 0
    End If
    rsSalary.Close
End Sub
