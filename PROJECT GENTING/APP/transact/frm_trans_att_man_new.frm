VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_trans_att_man 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INPUT KEHADIRAN MANUAL"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_all_employee 
      Caption         =   "All Employee"
      Height          =   255
      Left            =   1290
      TabIndex        =   49
      Top             =   1350
      Width           =   1455
   End
   Begin VB.TextBox txtkdkar 
      Height          =   285
      Left            =   4950
      TabIndex        =   48
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin prj_genting.LynxGrid LynxGrid2 
      Height          =   4515
      Left            =   1290
      TabIndex        =   34
      Top             =   1980
      Visible         =   0   'False
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7964
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
   Begin VB.CommandButton vbbutton5 
      Caption         =   "EXIT"
      Height          =   585
      Left            =   2790
      TabIndex        =   33
      Top             =   5160
      Width           =   1905
   End
   Begin VB.CommandButton vbbutton2 
      Caption         =   "..."
      Height          =   285
      Left            =   2430
      TabIndex        =   32
      Top             =   1680
      Width           =   315
   End
   Begin VB.CommandButton vbbutton1 
      Caption         =   "SAVE"
      Height          =   585
      Left            =   780
      TabIndex        =   4
      Top             =   5160
      Width           =   1905
   End
   Begin VB.TextBox txtjmljam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   885
      Left            =   5250
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   1065
   End
   Begin TDBTime6Ctl.TDBTime ttin 
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Top             =   3000
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":006C
      Spin            =   "frm_trans_att_man_new.frx":00BC
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   0
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
      Text            =   "00:00"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1320
      TabIndex        =   29
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   94502915
      CurrentDate     =   40823
   End
   Begin VB.TextBox txtnmDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   28
      Top             =   930
      Width           =   2805
   End
   Begin VB.TextBox txtkdshift 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6660
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
      Top             =   930
      Width           =   855
   End
   Begin VB.TextBox txtnmshift 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   7530
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   26
      Top             =   930
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   150
      TabIndex        =   24
      Top             =   4620
      Width           =   10695
   End
   Begin VB.TextBox txtkdtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   1290
      TabIndex        =   23
      Top             =   2010
      Width           =   795
   End
   Begin VB.TextBox txtkddiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   7590
      TabIndex        =   22
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   1680
      Width           =   1125
   End
   Begin VB.TextBox txtnmkar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2790
      TabIndex        =   14
      Top             =   1680
      Width           =   4005
   End
   Begin VB.TextBox txtdivision 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8340
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   2100
      TabIndex        =   12
      Top             =   2010
      Width           =   3105
   End
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   5340
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3030
      Width           =   5565
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   10
      Top             =   480
      Width           =   2025
   End
   Begin VB.TextBox txtentry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   8760
      TabIndex        =   8
      Top             =   150
      Width           =   2025
   End
   Begin VB.ComboBox cmbdep 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1290
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   930
      Width           =   1095
   End
   Begin TDBTime6Ctl.TDBTime ttout 
      Height          =   285
      Left            =   1290
      TabIndex        =   2
      Top             =   3330
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frm_trans_att_man_new.frx":00E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_att_man_new.frx":0150
      Spin            =   "frm_trans_att_man_new.frx":01A0
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   0
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
      Text            =   "00:00"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Hari Libur"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1290
      TabIndex        =   35
      Top             =   2730
      Width           =   1005
   End
   Begin VB.TextBox txt_jml_lembur 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3540
      TabIndex        =   36
      Top             =   3030
      Width           =   1485
   End
   Begin VB.TextBox txt_jml_jam 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2370
      TabIndex        =   38
      Top             =   3030
      Width           =   1065
   End
   Begin VB.TextBox Combo12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4230
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtBreak 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      TabIndex        =   40
      Top             =   3630
      Width           =   855
   End
   Begin VB.TextBox txt_absensi 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   42
      Top             =   2340
      Width           =   2805
   End
   Begin VB.ComboBox cmbAbsensi 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frm_trans_att_man_new.frx":01C8
      Left            =   1290
      List            =   "frm_trans_att_man_new.frx":01E1
      TabIndex        =   43
      Top             =   2340
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   6930
      TabIndex        =   45
      Top             =   2370
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94502915
      CurrentDate     =   40823
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   315
      Left            =   9300
      TabIndex        =   46
      Top             =   2370
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   94502915
      CurrentDate     =   40823
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frm_trans_att_man_new.frx":01FE
      Left            =   8580
      List            =   "frm_trans_att_man_new.frx":0200
      TabIndex        =   44
      Text            =   "..."
      Top             =   2370
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "LAIN LAIN"
      Height          =   195
      Left            =   8970
      TabIndex        =   72
      Top             =   5550
      Width           =   765
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      Caption         =   "CUTI"
      Height          =   195
      Left            =   8970
      TabIndex        =   71
      Top             =   5310
      Width           =   375
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "TUGAS DINAS"
      Height          =   195
      Left            =   8970
      TabIndex        =   70
      Top             =   5070
      Width           =   1095
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "MANGKIR"
      Height          =   195
      Left            =   5910
      TabIndex        =   69
      Top             =   5760
      Width           =   750
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "IJIN"
      Height          =   195
      Left            =   5910
      TabIndex        =   68
      Top             =   5520
      Width           =   285
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "SAKIT"
      Height          =   195
      Left            =   5910
      TabIndex        =   67
      Top             =   5280
      Width           =   465
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "HADIR"
      Height          =   195
      Left            =   5910
      TabIndex        =   66
      Top             =   5040
      Width           =   510
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   8850
      TabIndex        =   65
      Top             =   5550
      Width           =   45
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   8850
      TabIndex        =   64
      Top             =   5310
      Width           =   45
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   8850
      TabIndex        =   63
      Top             =   5070
      Width           =   45
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   5790
      TabIndex        =   62
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   5790
      TabIndex        =   61
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   5790
      TabIndex        =   60
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   5790
      TabIndex        =   59
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "TH"
      Height          =   195
      Left            =   8580
      TabIndex        =   58
      Top             =   5550
      Width           =   225
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "L"
      Height          =   195
      Left            =   8580
      TabIndex        =   57
      Top             =   5310
      Width           =   90
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "DT"
      Height          =   195
      Left            =   8580
      TabIndex        =   56
      Top             =   5070
      Width           =   225
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "SK"
      Height          =   195
      Left            =   5520
      TabIndex        =   55
      Top             =   5760
      Width           =   210
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "PR"
      Height          =   195
      Left            =   5520
      TabIndex        =   54
      Top             =   5520
      Width           =   225
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "S"
      Height          =   195
      Left            =   5520
      TabIndex        =   53
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "P"
      Height          =   195
      Left            =   5520
      TabIndex        =   52
      Top             =   5040
      Width           =   105
   End
   Begin VB.Label Label18 
      Caption         =   "LEGEND :"
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
      Left            =   5520
      TabIndex        =   51
      Top             =   4770
      Width           =   945
   End
   Begin VB.Label Label17 
      Caption         =   "* yyyy-MM-dd"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2970
      TabIndex        =   50
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label16 
      Caption         =   "Dari :"
      Height          =   195
      Left            =   6390
      TabIndex        =   47
      Top             =   2430
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Status Absensi :"
      Height          =   210
      Left            =   60
      TabIndex        =   41
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Lembur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3540
      TabIndex        =   37
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "Jam Kerja :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2370
      TabIndex        =   31
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shift :"
      Height          =   195
      Left            =   6210
      TabIndex        =   25
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Label5 
      Caption         =   "Karyawan :"
      Height          =   210
      Left            =   420
      TabIndex        =   21
      Top             =   1740
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Divisi : "
      Height          =   210
      Left            =   7050
      TabIndex        =   20
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Jabatan :"
      Height          =   210
      Left            =   540
      TabIndex        =   19
      Top             =   2070
      Width           =   705
   End
   Begin VB.Label Label8 
      Caption         =   "Jam Masuk :"
      Height          =   195
      Left            =   330
      TabIndex        =   18
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label Label9 
      Caption         =   "Remark :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5370
      TabIndex        =   17
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Jam Istirahat :"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   3660
      Width           =   1005
   End
   Begin VB.Label Label14 
      Caption         =   "Jam Pulang :"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label Label11 
      Caption         =   "Nama User :"
      Height          =   210
      Left            =   7830
      TabIndex        =   11
      Top             =   540
      Width           =   885
   End
   Begin VB.Label Label10 
      Caption         =   "Tanggal Input :"
      Height          =   210
      Left            =   7620
      TabIndex        =   9
      Top             =   210
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Dept / Area :"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   960
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal :"
      Height          =   195
      Left            =   570
      TabIndex        =   5
      Top             =   240
      Width           =   705
   End
End
Attribute VB_Name = "frm_trans_att_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As New ADODB.Recordset
Dim strsql As String
Public editTrans As Boolean
Public v_tt_in, v_tt_out, v_absen_status As String
Dim v_15, v_2, v_3, v_4 As Double
Dim v_tot_work, v_tot_ot As Double
Public v_interval As Double
Dim v_start_time, v_end_time As Date
Dim v_flag_ot As String
Dim v_meal, v_trans As Double
Dim time_in, time_out As String


Private Sub createKar()
   With LynxGrid2
      .AddColumn "NIK", 700, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "Division", 2000, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "Title", 2000, , , , , , , , , True
      .AddColumn "level", 2000, , , , , , , , False
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        strsql = "select nik,employee_name,division_code,division_name," _
                    & "title_code,title_name,level_code,employee_code " _
                & "from m_employee " _
                & "WHERE flag_active <> 0 AND department_code = '" & cmbdep.Text & "' AND " _
                & "(nik LIKE '%" & txt_nik.Text & "%' " _
                & "OR employee_name LIKE '%" & txt_nik.Text & "%') " _
                & "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
                
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs2.MoveFirst
            While Not rs2.EOF
                LynxGrid2.AddItem rs2!nik & vbTab & rs2!EMPLOYEE_NAME & vbTab & _
                rs2!division_code & vbTab & rs2!division_name & vbTab & rs2!title_code & vbTab & _
                rs2!title_name & vbTab & rs2!employee_code
                rs2.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs2.RecordCount = 1 Then
                rs2.MoveFirst
                txtkdkar.Text = rs2!employee_code
                txtnmkar.Text = rs2!EMPLOYEE_NAME
                txtkddiv.Text = rs2!division_code
                txtdivision.Text = IIf(IsNull(rs2!division_name) = True, "", rs2!division_name)
                txtkdtitle.Text = rs2!title_code
                txttitle.Text = rs2!title_name
                txt_nik.Text = rs2!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs2.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txtnmkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txtkddiv.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
            txtdivision.Text = LynxGrid2.CellText(LynxGrid2.Row, 3)
            txtkdtitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 4)
            txttitle.Text = LynxGrid2.CellText(LynxGrid2.Row, 5)
            txtkdkar.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
'            ttin.SetFocus
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub Check1_Click()
    txtjmljam_Change
    
    Call totalJam
End Sub
Private Sub chk_all_employee_Click()
    If chk_all_employee.Value = 0 Then
        txt_nik.Enabled = True
        vbButton2.Enabled = True
    Else
        txt_nik.Enabled = False
        vbButton2.Enabled = False
    End If
End Sub

Public Sub cmbAbsensi_Click()
Select Case cmbAbsensi
    Case "P"
        txt_absensi.Text = "HADIR"
    Case "S"
        txt_absensi.Text = "SAKIT"
    Case "PR"
        txt_absensi.Text = "IJIN"
    Case "SK"
        txt_absensi.Text = "MANGKIR"
    Case "DT"
        txt_absensi.Text = "TUGAS DINAS"
    Case "L"
        txt_absensi.Text = "CUTI"
    Case "TH"
        txt_absensi.Text = "LAIN-LAIN"
End Select

If cmbAbsensi.Text <> "P" Then
    DTPicker2.Visible = True
    Label16.Visible = True
    Combo3.Visible = True
    ttin.Enabled = False
    txtBreak.Enabled = False
    ttout.Enabled = False
    DTPicker2.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
    DTPicker3.Value = Format(DTPicker1.Value, "yyyy-MM-dd")
    ttin.Text = "00:00"
    ttout.Text = "00:00"
'    txt_jml_jam.Text = "0"
    txt_jml_lembur.Text = "0"
Else
    Label16.Visible = False
    DTPicker2.Visible = False
    DTPicker3.Visible = False
    Combo3.Visible = False
    ttin.Enabled = True
    ttout.Enabled = True
    txtBreak.Enabled = True
    ttin.Text = Format(v_start_time, "HH:mm")
    ttout.Text = Format(v_end_time, "HH:mm")
End If
End Sub

Private Sub txtBreak_Change()
    Call totalJam
End Sub

Private Sub totalJam()
    Dim tglin As Date, tglout As Date
    Dim jambreak As Integer
    Dim tot_menit As Double
    Dim tgl_masuk As Date
    Dim tgl_keluar As Date
    Dim menit As Double
    Dim jam As Double
    Dim v_ot As Double
    Dim hour As Double
    
    Dim e As Long
    
    tglin = Date & " " & Format(ttin.Value, "hh:nn:ss")
    
    If Format(ttin.Value, "hh:nn:ss") > Format(ttout.Value, "hh:nn:s") Then
        tglout = Date + 1 & " " & ttout.Value
    Else
        tglout = Date & " " & ttout.Value
    End If
    
    jambreak = Val(txtBreak.Text)
    
    tgl_masuk = Format(tglin, "yyyy-MM-dd hh:nn:ss")
    tgl_keluar = Format(tglout, "yyyy-MM-dd hh:nn:ss")
  
    hour = DateDiff("h", tgl_masuk, tgl_keluar) - jambreak
    menit = (DateDiff("n", tgl_masuk, tgl_keluar)) / 60
    jam = roundDown(menit)
    
    e = (menit - jam) * 60

If Check1.Value = 0 Then
    txt_jml_jam.Text = DateDiff("h", time_in, time_out) - v_interval
    
    If (hour - txt_jml_jam) < 1 Then
        If e <= 30 Then
            tot_menit = 0.5
        Else
            tot_menit = 1
        End If
        
        txtjmljam.Text = jam - jambreak + tot_menit
    Else
        txtjmljam.Text = jam - jambreak + (e / 60)
    End If
    
    If txtjmljam.Text < 0 Then
        txtjmljam.Text = 0
    Else
        txtjmljam.Text = txtjmljam.Text
    End If
    
    v_ot = Val(txtjmljam.Text) - txt_jml_jam.Text
    
    If v_ot = 0.5 Or v_ot = 0 Then
        txt_jml_lembur.Text = 0
    Else
        txt_jml_lembur.Text = Val(txtjmljam.Text) - txt_jml_jam.Text
    End If
Else
    txt_jml_jam.Text = 0
    
    If hour < 1 Then
        If e <= 30 Then
            tot_menit = 0.5
        Else
            tot_menit = 1
        End If
        
        txtjmljam.Text = jam - jambreak + tot_menit
    Else
        txtjmljam.Text = jam - jambreak + (e / 60)
    End If
    
    If txtjmljam.Text < 0 Then
        txtjmljam.Text = 0
    Else
        txtjmljam.Text = txtjmljam.Text
    End If
    
    v_ot = Val(txtjmljam.Text) - txt_jml_jam.Text
    
    If v_ot = 0.5 Or v_ot = 0 Then
        txt_jml_lembur.Text = 0
    Else
        If v_ot > 15 Then
            txt_jml_lembur.Text = 15
        Else
            txt_jml_lembur.Text = Val(txtjmljam.Text)
        End If
    End If
End If

If cmbAbsensi.Text <> "P" Then
    txt_jml_jam.Text = "0"
    txt_jml_lembur.Text = "0"
End If

End Sub

Private Sub Combo3_Click()
    If Combo3.Text = "To" Then
        DTPicker3.Value = DTPicker2.Value
        DTPicker3.Visible = True
    Else
        DTPicker3.Visible = False
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        vbButton1_Click
    End If
End Sub

Private Sub Form_Load()
    createKar
    
    
    If editTrans = False Then
        isijamkerja
    End If
        
        
    Combo3.AddItem "..."
    Combo3.AddItem "To"
    Combo3.Text = "..."
    
End Sub

Public Sub isijamkerja()
Dim rsshift As New ADODB.Recordset

    strsql = "SELECT start_time, end_time, break_interval_minute from m_shift " & _
            "where shift_code = '" & frm_list_manual_att.TDBCombo1.Text & "' " & _
            "AND company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
    rsshift.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsshift.RecordCount > 0 Then
        v_start_time = rsshift!start_time
        v_end_time = rsshift!end_time
        v_interval = rsshift!break_interval_minute
    End If
    rsshift.Close
    
    time_in = v_start_time
    If Format(v_end_time, "Hh:mm") < Format(v_start_time, "HH:mm") Then
        time_out = v_end_time + 1
    Else
        time_out = v_end_time
    End If
    
    txtBreak.Text = v_interval
    ttin.Text = Format(v_start_time, "HH:mm")
    ttout.Text = Format(v_end_time, "Hh:mm")
    
    txt_jml_jam.Text = DateDiff("h", time_in, time_out) - txtBreak.Text
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

Private Sub TDBCombo1_ItemChange()
    If TDBCombo1.ApproxCount > 0 Then
        TDBCombo1.Text = TDBCombo1.Columns("absent_code").Value
        txtnmabsentstatus.Text = TDBCombo1.Columns("absent_name").Value
        
        If TDBCombo1.Text <> "0" Then
            DTPicker2.Visible = True
            Label16.Visible = True
            Combo3.Visible = True
            ttin.Enabled = False
            txtBreak.Enabled = False
            ttout.Enabled = False
        Else
            Label16.Visible = False
            DTPicker2.Visible = False
            DTPicker3.Visible = False
            Combo3.Visible = False
            ttin.Enabled = True
            ttout.Enabled = True
            txtBreak.Enabled = True
        End If
    End If
End Sub

Private Sub ttin_Change()
    Call totalJam
End Sub

Private Sub ttin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBreak.SetFocus
    End If
End Sub

Private Sub ttout_Change()
    Call totalJam
End Sub

Private Sub ttout_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtket.SetFocus
    End If
End Sub

Private Sub txtjmljam_Change()
On Error Resume Next

'v_tot_work = DateDiff("h", v_start_time, v_end_time) - v_interval
'v_tot_ot = Val(txtjmljam.Text) - v_tot_work
'v_tot_ot = txt_jml_lembur.Text

If Check1.Value = 1 Then
    If v_flag_ot = 1 Then
        v_tot_work = 0
        If v_tot_ot > 15 Then
            v_tot_ot = 15
        Else
            v_tot_ot = Val(txtjmljam.Text)
        End If
            Select Case Val(v_tot_ot)
                Case Is <= 7
                    v_15 = 0
                    v_2 = Val(v_tot_ot)
                    v_3 = 0
                    v_4 = 0
                Case Is = 8
                    v_15 = 0
                    v_2 = 7
                    v_3 = 1
                    v_4 = 0
                Case Else
                    v_15 = 0
                    v_2 = 7
                    If (v_tot_ot - v_2) < 1 Then
                        v_3 = v_tot_ot - v_2
                    Else
                        v_3 = 1
                    End If
                    v_4 = Val(v_tot_ot) - Val(v_2) - Val(v_3)
            End Select
        
            v_trans = 2
        
            If v_tot_ot >= 3 And v_tot_ot < 11 Then
                v_meal = 1
            ElseIf v_tot_ot >= 11 Then
                v_meal = 2
            Else
                v_meal = 0
            End If
        Else
            v_tot_ot = 0
            v_15 = 0
            v_2 = 0
            v_3 = 0
            v_4 = 0
            v_meal = 0
            v_trans = 0
        End If
Else
    If v_flag_ot = 1 Then
        v_tot_work = DateDiff("h", time_in, time_out) - v_interval
        v_tot_ot = Val(txtjmljam.Text) - v_tot_work
        Select Case Val(v_tot_ot)
            Case Is = 1
                v_15 = 1
                v_2 = 0
                v_3 = 0
                v_4 = 0
            Case Is > 1
                v_15 = 1
                v_2 = Val(v_tot_ot) - Val(v_15)
                v_3 = 0
                v_4 = 0
            Case Else
                v_15 = 0
                v_2 = 0
                v_3 = 0
                v_4 = 0
        End Select
        
        If ttin.Text <= "06:00" Or ttout.Text >= "22:00" Then
            v_trans = 1
        Else
            v_trans = 0
        End If
        
        If v_tot_ot >= 3 And v_tot_ot < 9 Then
            v_meal = 1
        ElseIf v_tot_ot >= 9 Then
            v_meal = 2
        Else
            v_meal = 0
        End If
    Else
        v_tot_ot = 0
        v_15 = 0
        v_2 = 0
        v_3 = 0
        v_4 = 0
        v_meal = 0
        v_trans = 0
    End If
End If

'If v_flag_ot = 1 Then
'    If ttin.Text < "06:00" Or ttout.Text > "22:00" Then
'        v_trans = 1
'    Else
'        v_trans = 0
'    End If
'
'    If txtjmljam.Text >= 3 And txtjmljam.Text < 6 Then
'        v_meal = 1
'    ElseIf txtjmljam.Text > 6 Then
'        v_meal = 2
'    Else
'        v_meal = 0
'    End If
'Else
'    v_meal = 0
'    v_trans = 0
'End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        vbButton1.SetFocus
    End If
End Sub

Private Sub vbButton1_Click()
Dim rsemp As New ADODB.Recordset
Dim aa As Integer
Dim a, b, c, d As String

    If cmbAbsensi = "P" Then
'        If ttin.Value = ttout.Value Then
'            MsgBox "Invalid Check In Time...!!!", vbExclamation, "Error"
'            Exit Sub
'        End If
        
'        If Format(ttout.Value, "hh:mm") = "00:00" Then
'            MsgBox "Invalid Format Check Out Time....!!!", vbExclamation, "Error"
'            Exit Sub
'        End If
        
        If Format(ttout.Value, "hh:mm") = "" Or Format(ttout.Value, "hh:mm") = "" Then
            MsgBox "Invalid Format Check In or Check Out Time....!!!", vbExclamation, "Error"
            Exit Sub
        End If
    End If
    
    '+++++++++++++++++++APAKAH KODE KARYAWAN SUDAH BENAR++++++++++++++++++++
    If chk_all_employee.Value = 0 Then
        strsql = "SELECT 1 FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount = 0 Then
            MsgBox "Invalid Employee Code...!!", vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False And chk_all_employee.Value = 0 Then
        strsql = "SELECT 1 FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
            & "AND DATE(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If rs2.RecordCount > 0 Then
            MsgBox "This Employee Is Already Exist...!!", vbCritical
            rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    If editTrans = False And chk_all_employee.Value = 0 Then
        If cmbAbsensi.Text = "0" Then 'Present
            strsql = "SELECT att_date,shift_code FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND DATE(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
            rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        ElseIf Combo3.ListIndex = 1 Then 'Lebih dr 1 hari
            strsql = "SELECT att_date,shift_code FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND (DATE(att_date) BETWEEN '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "' " _
                & "AND '" & Format(DTPicker3.Value, "yyyy-MM-dd") & "')"
            rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        Else
            strsql = "SELECT att_date,shift_code FROM h_attendance WHERE employee_code = '" & txtkdkar.Text & "' " _
                & "AND DATE(att_date) = '" & Format(DTPicker2.Value, "yyyy-MM-dd") & "'"
            rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        End If
                
        If rs2.RecordCount > 0 Then
            Dim v_shift As String
            v_shift = IIf(rs2!shift_code = "P", "HADIR", IIf(rs2!shift_code = "S", "SAKIT", IIf(rs2!shift_code = "PR", "IJIN", _
                    IIf(rs2!shift_code = "SK", "MANGKIR", IIf(rs2!shift_code = "DT", "TUGAS DINAS", IIf(rs2!shift_code = "L", "CUTI", _
                    "LAIN-LAIN"))))))
            
                MsgBox "This Employee Is Already Exist With Status " & v_shift, vbCritical
                rs2.Close
            Exit Sub
        End If
        rs2.Close
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Dim tglawal As Date, tglAkhir As Date
    Dim v_flag_active As Integer
    
    tglawal = DTPicker2.Value
    If Combo3.Text = "Hingga" Then
        tglAkhir = DTPicker3.Value
    Else
        tglAkhir = DTPicker2.Value
    End If
    
    strsql = "SELECT employee_code, flag_active FROM m_employee WHERE employee_code = '" & txtkdkar.Text & "'"
    rs2.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        If Not rs2.EOF Then
            v_flag_active = rs2!flag_active
        End If
    rs2.Close
    
    If v_flag_active = 2 Then
        MsgBox "Employees in the MC condition, Update Attendance does not Allowed!"
        Exit Sub
    End If
    
        Select Case cmbAbsensi.Text
        Case "P"
            InsertPresent
        Case "SK", "PR", "S", "TH"
            Dim abstatus As Integer
            If cmbAbsensi.Text = "SK" Then
                abstatus = 2
            ElseIf cmbAbsensi.Text = "TH" Then
                abstatus = 6
            ElseIf cmbAbsensi.Text = "PR" Then
                abstatus = 0
            ElseIf cmbAbsensi.Text = "S" Then
                abstatus = 1
'            ElseIf TDBCombo1.Text = "PR" Then
'                abstatus = 3
            End If
            
            CnG.BeginTrans
            While tglawal <= tglAkhir
                
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                    strsql = "SELECT employee_code FROM m_employee " _
                            & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " _
                            & "AND department_code = '" & cmbdep.Text & "' " _
                            & "AND flag_active <> 0"
                    rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    For aa = 0 To rsemp.RecordCount - 1
                        strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute strsql

                        strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,flag_present,absent_status,description,entry_date,userinput,shift_number) " _
                            & "VALUES " _
                            & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & cmbAbsensi.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & txtkdshift.Text & "')"
                        CnG.Execute strsql
                        rsemp.MoveNext
                    Next
                    rsemp.Close
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> cmbAbsensi.Text Then
                        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                        CnG.Execute strsql
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & txtkdkar.Text & "'"
                    CnG.Execute strsql
                    
                    strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_number,shift_code,flag_present,absent_status,description,entry_date,userinput) " _
                        & "VALUES " _
                        & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & txtkdkar.Text & "','" & txtkdshift.Text & "','" & cmbAbsensi.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "')"
                    CnG.Execute strsql
                End If
                
                tglawal = tglawal + 1
                
                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
            Wend
            CnG.CommitTrans
        Case "DT"
            Dim waktu_in As String, waktu_out As String, waktu_break_out As String, waktu_break_in As String
            
            CnG.BeginTrans
            While tglawal <= tglAkhir
                
                    waktu_in = Format(tglawal, "yyyy-MM-dd") & " 08:00:00"
                    waktu_out = Format(tglawal, "yyyy-MM-dd") & " 17:00:00"
                    waktu_break_out = Format(tglawal, "yyyy-MM-dd") & " 13:00:00"
                    waktu_break_in = Format(tglawal, "yyyy-MM-dd") & " 12:00:00"
                
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    strsql = "SELECT employee_code from m_employee " _
                            & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " _
                            & "AND department_code = '" & cmbdep.Text & "' " _
                            & "AND flag_active <> 0"
                    rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

                    For aa = 0 To rsemp.RecordCount - 1
                        strsql = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute strsql

                        strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,time_in,time_out,time_out_break,time_in_break," _
                                & "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number) " _
                            & "VALUES " _
                            & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & cmbAbsensi.Text & "','" & waktu_in & "','" & waktu_out & "','" & waktu_break_out & "','" & waktu_break_in & "'," _
                                & "1,1,0,'" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & txtkdshift.Text & "')"
                        CnG.Execute strsql
                        rsemp.MoveNext
                    Next
                    rsemp.Close
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> cmbAbsensi.Text Then
                        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                        CnG.Execute strsql
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "DELETE FROM h_attendance WHERE att_date = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & txtkdkar.Text & "'"
                    CnG.Execute strsql
                        
                    strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_number,shift_code,time_in,time_out,time_out_break,time_in_break," _
                            & "flag_present,flag_duty,absent_status,description,entry_date,userinput) " _
                        & "VALUES " _
                        & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & txtkdkar.Text & "','" & cmbAbsensi.Text & "','" & cmbAbsensi.Text & "','" & waktu_in & "','" & waktu_out & "','" & waktu_break_out & "','" & waktu_break_in & "'," _
                            & "1,1,0,'" & txtket.Text & "',now(),'" & LOGIN_CODE & "')"
                    CnG.Execute strsql
                End If
                
                tglawal = tglawal + 1
                
                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
            Wend
            CnG.CommitTrans
            
        Case "L"
            Dim rsleave As New ADODB.Recordset
            abstatus = 3
            
            CnG.BeginTrans
                While tglawal <= tglAkhir
                If chk_all_employee.Value = 1 Then
                    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                    strsql = "SELECT employee_code from m_employee " _
                            & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " _
                            & "AND department_code = '" & cmbdep.Text & "' " _
                            & "AND flag_active <> 0"
                    rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

                    For aa = 0 To rsemp.RecordCount - 1
                        strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                        CnG.Execute strsql

                        strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_code,flag_present,absent_status,description,entry_date,userinput,shift_number) " _
                            & "VALUES " _
                            & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & rsemp!employee_code & "','" & cmbAbsensi.Text & "',0,'" & abstatus & "','" & txtket.Text & "',now(),'" & LOGIN_CODE & "','" & txtkdshift.Text & "')"
                        CnG.Execute strsql
                        rsemp.MoveNext
                    Next
                Else
                    '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                    If v_absen_status <> cmbAbsensi.Text Then
                        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                        CnG.Execute strsql
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                    strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(tglawal, "yyyy-MM-dd") & "' AND employee_code = '" & txtkdkar.Text & "'"
                    CnG.Execute strsql
                    
                    strsql = "INSERT INTO h_attendance (att_date,employee_code,shift_number,shift_code,flag_present,absent_status,description,entry_date,userinput) " _
                        & "VALUES " _
                        & "('" & Format(tglawal, "yyyy-MM-dd") & "','" & txtkdkar.Text & "','" & cmbAbsensi.Text & "','" & cmbAbsensi.Text & "',0,3,'" & txtket.Text & "',now(),'" & LOGIN_CODE & "')"
                    CnG.Execute strsql
                End If
                
                tglawal = tglawal + 1
                
                CnG.Execute "call spg_leave_periode2 ('" & Format(tglawal, "yyyy-MM-dd") & "')"
            Wend
            CnG.CommitTrans
        Case Else
            MsgBox "Please Check Absent Status....!!!"
            Exit Sub
    End Select
    
    MsgBox "Save Succesfully..", vbInformation, headerMSG
    txtkdkar.Text = ""
    txt_nik.Text = ""
    txtnmkar.Text = ""
    txtkddiv.Text = ""
    txtdivision.Text = ""
    txtkdtitle.Text = ""
    txttitle.Text = ""
    ttin.Value = Format(v_start_time, "HH:mm")
    ttout.Value = Format(v_end_time, "HH:mm")
    txtket.Text = ""
    txtBreak.Text = v_interval
    cmbAbsensi.Text = "P"
    cmbAbsensi_Click
    editTrans = False
    
    frm_list_manual_att.isiGridAbsen
End Sub

Private Sub vbbutton2_Click()
    isiGridKar (1)
End Sub

Private Sub vbButton5_Click()
    Unload Me
End Sub

Private Sub InsertPresent()
Dim start_time As String, end_time As String, max_break_out As String, min_break_in As String
Dim time_in As String, time_out_break As String, time_in_break As String, time_out As String
Dim tot_overtime As Double
Dim rsemp As New ADODB.Recordset

    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
    strsql = "Select CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
            & "CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
            & "CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(max_break_out)) as datetime) max_break_out," _
            & "CAST(concat('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',' ', time(min_break_in)) as datetime) min_break_in," _
            & "curdate() tglserver " _
        & "from m_shift where shift_code= '" & txtkdshift.Text & "'"
    rs2.Open strsql, CnG, adOpenDynamic, adLockReadOnly
    
    start_time = Format(rs2!start_time, "yyyy-MM-dd hh:mm:ss")
    end_time = Format(rs2!end_time, "yyyy-MM-dd hh:mm:ss")
    
    max_break_out = Format(DateAdd("h", Val(txtBreak.Text), rs2!min_break_in), "yyyy-MM-dd hh:mm:ss")
    
    min_break_in = Format(rs2!min_break_in, "yyyy-MM-dd hh:mm:ss")
        
    time_in = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttin.Value, "hh:mm") & ":00"
        
    time_out_break = max_break_out
    time_in_break = min_break_in
    
    If Format(ttout.Value, "hh:mm:ss") <= Format(ttin.Value, "hh:mm:ss") Then
        time_out = Format(DTPicker1.Value + 1, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    Else
        time_out = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    End If
    
'    time_out = Format(DTPicker1.Value, "yyyy-MM-dd") & " " & Format(ttout.Value, "hh:mm") & ":00"
    
    rs2.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If chk_all_employee.Value = 1 Then
'        strsql = "SELECT a.employee_code, b.flag_ot from m_employee a join m_salary_standard b on a.employee_code = b.employee_code " _
'                & "where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' AND " _
'                & "department_code = '" & cmbdep.Text & "' order by b.number desc limit 1"
'        rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

        strsql = "SELECT (SELECT employee_code FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1) employee_code, " & _
                    "(SELECT flag_ot FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1) flag_ot, " & _
                    "(SELECT salary_date FROM m_salary_standard WHERE employee_code = a.employee_code ORDER BY number DESC LIMIT 1) salary_date " & _
                "FROM m_employee a " & _
                "WHERE company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "' " & _
                    "AND department_code = '" & cmbdep.Text & "' " & _
                    "AND flag_active <> 0 " & _
                "ORDER BY a.employee_code"
        rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        For aa = 0 To rsemp.RecordCount - 1
            v_flag_ot = IIf(IsNull(rsemp!flag_ot), 0, rsemp!flag_ot)
        
            txtjmljam_Change
            
            tot_overtime = ((1.5 * Val(v_15)) + (2 * Val(v_2)) + (3 * Val(v_3)) + (4 * Val(v_4)))
            
            If editTrans = False Then
                '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
                strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                CnG.Execute strsql
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    
                strsql = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' AND employee_code = '" & rsemp!employee_code & "'"
                CnG.Execute strsql
                        
                strsql = "INSERT INTO h_attendance (employee_code,att_date,shift_code,shift_number,start_time,end_time," _
                    & "time_in,time_in_diff," _
                    & "time_out_break,time_out_break_diff," _
                    & "time_in_break,time_in_break_diff," _
                    & "time_out,time_out_diff,description,entry_date,userinput,absent_status,flag_present," _
                    & "total_ot,15jam,2jam,3jam,4jam,holiday,tot_overtime,flag_meal,flag_transport) VALUES " _
                    & "('" & rsemp!employee_code & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & cmbAbsensi.Text & "','" & txtkdshift.Text & "','" & start_time & "','" & end_time & "'," _
                    & "'" & time_in & "',CASE WHEN TIMEDIFF('" & time_in & "','" & start_time & "') < 0 THEN TIMEDIFF('" & start_time & "','" & time_in & "') ELSE TIMEDIFF('" & time_in & "','" & start_time & "') END," _
                    & "'" & max_break_out & "',CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                    & "'" & min_break_in & "',CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                    & "'" & time_out & "',CASE WHEN TIMEDIFF('" & time_out & "','" & end_time & "') < 0 THEN TIMEDIFF('" & end_time & "','" & time_out & "') ELSE TIMEDIFF('" & time_out & "','" & end_time & "') END," _
                    & "'" & txtket.Text & "',now(),'" & LOGIN_CODE & "',0,1," _
                    & "'" & txt_jml_lembur & "','" & v_15 & "','" & v_2 & "','" & v_3 & "'," _
                    & "'" & v_4 & "','" & Check1.Value & "','" & tot_overtime & "','" & v_meal & "','" & v_trans & "')"
            Else
                '+++++++++++++++++++++++++++++++++ Check Edit +++++++++++++++++++++++++++++
                If v_absen_status <> TDBCombo1.Text Then
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                ElseIf v_tt_in <> ttin.Value Or v_tt_out <> ttout.Value Then
                    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                    CnG.Execute strsql
                End If
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                strsql = "UPDATE h_attendance set employee_code = '" & rsemp!employee_code & "'," _
                    & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & cmbAbsensi.Text & "'," _
                    & "shift_number = '" & txtkdshift.Text & "',start_time = '" & start_time & "',end_time = '" & end_time & "'," _
                    & "time_in = '" & time_in & "',flag_present = 1," _
                    & "time_in_diff = CASE WHEN TIMEDIFF('" & time_in & "','" & start_time & "') < 0 THEN TIMEDIFF('" & start_time & "','" & time_in & "') ELSE TIMEDIFF('" & time_in & "','" & start_time & "') END," _
                    & "time_out_break = '" & max_break_out & "'," _
                    & "time_out_break_diff = CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                    & "time_in_break = '" & min_break_in & "'," _
                    & "time_in_break_diff = CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                    & "time_out = '" & time_out & "'," _
                    & "time_out_diff = CASE WHEN TIMEDIFF('" & time_out & "','" & end_time & "') < 0 THEN TIMEDIFF('" & end_time & "','" & time_out & "') ELSE TIMEDIFF('" & time_out & "','" & end_time & "') END," _
                    & "description = '" & txtket.Text & "',editdate = now(),useredit = '" & LOGIN_CODE & "',absent_status = 0," _
                    & "total_ot = '" & txt_jml_lembur & "',15jam = '" & v_15 & "',2jam = '" & v_2 & "'," _
                    & "3jam = '" & v_3 & "',4jam = '" & v_4 & "',holiday='" & Check1.Value & "'," _
                    & "tot_overtime = '" & tot_overtime & "',flag_meal = '" & v_meal & "',flag_transport = '" & v_trans & "' " _
                    & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                    & "AND att_date = '" & Format(frm_list_manual_att.DTPicker1.Value, "yyyy-MM-dd") & "' " _
                    & "AND shift_number = '" & txtkdshift.Text & "'"
            End If
            CnG.Execute strsql
            rsemp.MoveNext
        Next
        rsemp.Close
    Else
        strsql = "SELECT flag_ot from m_salary_standard " _
                & "where employee_code = '" & txtkdkar & "' order by number desc limit 1"
        rsemp.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rsemp.RecordCount > 0 Then
            v_flag_ot = IIf(IsNull(rsemp!flag_ot), 0, rsemp!flag_ot)
        End If
        rsemp.Close
        
        txtjmljam_Change
        
        tot_overtime = ((1.5 * Val(v_15)) + (2 * Val(v_2)) + (3 * Val(v_3)) + (4 * Val(v_4)))
        
        If editTrans = False Then
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
            CnG.Execute strsql
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                  
            strsql = "INSERT INTO h_attendance (employee_code,att_date,shift_code,shift_number,start_time,end_time," _
                    & "time_in,time_in_diff," _
                    & "time_out_break,time_out_break_diff," _
                    & "time_in_break,time_in_break_diff," _
                    & "time_out,time_out_diff,description,entry_date,userinput,absent_status,flag_present," _
                    & "total_ot,15jam,2jam,3jam,4jam,holiday,tot_overtime,flag_meal,flag_transport) VALUES " _
                    & "('" & txtkdkar.Text & "','" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & cmbAbsensi.Text & "','" & txtkdshift.Text & "','" & start_time & "','" & end_time & "'," _
                    & "'" & time_in & "',CASE WHEN TIMEDIFF('" & time_in & "','" & start_time & "') < 0 THEN TIMEDIFF('" & start_time & "','" & time_in & "') ELSE TIMEDIFF('" & time_in & "','" & start_time & "') END," _
                    & "'" & max_break_out & "',CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                    & "'" & min_break_in & "',CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                    & "'" & time_out & "',CASE WHEN TIMEDIFF('" & time_out & "','" & end_time & "') < 0 THEN TIMEDIFF('" & end_time & "','" & time_out & "') ELSE TIMEDIFF('" & time_out & "','" & end_time & "') END," _
                    & "'" & txtket.Text & "',now(),'" & LOGIN_CODE & "',0,1," _
                    & "'" & txt_jml_lembur & "','" & v_15 & "','" & v_2 & "','" & v_3 & "'," _
                    & "'" & v_4 & "','" & Check1.Value & "','" & tot_overtime & "','" & v_meal & "','" & v_trans & "')"
        Else
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            If v_tt_in <> ttin.Value Or v_tt_out <> ttout.Value Then
                strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & frm_list_manual_att.TDBCombo_company.Text & "'"
                CnG.Execute strsql
            End If
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            strsql = "UPDATE h_attendance set employee_code = '" & txtkdkar.Text & "'," _
                    & "att_date = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "',shift_code = '" & cmbAbsensi.Text & "'," _
                    & "shift_number = '" & txtkdshift.Text & "',start_time = '" & start_time & "',end_time = '" & end_time & "'," _
                    & "time_in = '" & time_in & "',flag_present = 1," _
                    & "time_in_diff = CASE WHEN TIMEDIFF('" & time_in & "','" & start_time & "') < 0 THEN TIMEDIFF('" & start_time & "','" & time_in & "') ELSE TIMEDIFF('" & time_in & "','" & start_time & "') END," _
                    & "time_out_break = '" & max_break_out & "'," _
                    & "time_out_break_diff = CASE WHEN TIMEDIFF('" & time_out_break & "','" & max_break_out & "') < 0 THEN TIMEDIFF('" & max_break_out & "','" & time_out_break & "') ELSE TIMEDIFF('" & time_out_break & "','" & max_break_out & "') END," _
                    & "time_in_break = '" & min_break_in & "'," _
                    & "time_in_break_diff = CASE WHEN TIMEDIFF('" & time_in_break & "','" & min_break_in & "') < 0 THEN TIMEDIFF('" & min_break_in & "','" & time_in_break & "') ELSE TIMEDIFF('" & time_in_break & "','" & min_break_in & "') END," _
                    & "time_out = '" & time_out & "'," _
                    & "time_out_diff = CASE WHEN TIMEDIFF('" & time_out & "','" & end_time & "') < 0 THEN TIMEDIFF('" & end_time & "','" & time_out & "') ELSE TIMEDIFF('" & time_out & "','" & end_time & "') END," _
                    & "description = '" & txtket.Text & "',editdate = now(),useredit = '" & LOGIN_CODE & "',absent_status = 0," _
                    & "total_ot = '" & txt_jml_lembur & "',15jam = '" & v_15 & "',2jam = '" & v_2 & "'," _
                    & "3jam = '" & v_3 & "',4jam = '" & v_4 & "',holiday='" & Check1.Value & "'," _
                    & "tot_overtime = '" & tot_overtime & "',flag_meal = '" & v_meal & "',flag_transport = '" & v_trans & "' " _
                    & "WHERE employee_code = '" & txtkdkar.Text & "' " _
                    & "AND att_date = '" & Format(frm_list_manual_att.DTPicker1.Value, "yyyy-MM-dd") & "' " _
                    & "AND shift_number = '" & txtkdshift.Text & "'"
        End If
        CnG.Execute strsql
    End If

End Sub
