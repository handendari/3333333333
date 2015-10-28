VERSION 5.00
Begin VB.Form frm_trans_salary_dtl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALARY PAYMENT"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11745
   Icon            =   "frm_trans_salary_dtl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_angsuran_koperasi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9210
      TabIndex        =   105
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txt_iuran_koperasi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9210
      TabIndex        =   103
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox txt_tax_correction 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9210
      TabIndex        =   101
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txt_description 
      Appearance      =   0  'Flat
      Height          =   1065
      Left            =   4260
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Top             =   5910
      Width           =   4125
   End
   Begin VB.TextBox txt_night_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7170
      TabIndex        =   89
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txt_night_days 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6150
      TabIndex        =   87
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txt_afternoon_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7170
      TabIndex        =   85
      Top             =   4860
      Width           =   1215
   End
   Begin VB.TextBox txt_afternoon_days 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6150
      TabIndex        =   83
      Top             =   4860
      Width           =   495
   End
   Begin VB.TextBox txt_transport_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      TabIndex        =   80
      Top             =   4110
      Width           =   1215
   End
   Begin VB.TextBox txt_transport_days 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6180
      TabIndex        =   78
      Top             =   4110
      Width           =   495
   End
   Begin VB.TextBox txt_meal_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      TabIndex        =   76
      Top             =   3810
      Width           =   1215
   End
   Begin VB.TextBox txt_meal_days 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6180
      TabIndex        =   74
      Top             =   3810
      Width           =   495
   End
   Begin VB.TextBox txt_other_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6180
      TabIndex        =   72
      Top             =   3480
      Width           =   2235
   End
   Begin VB.TextBox txt_position_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6180
      TabIndex        =   70
      Top             =   3180
      Width           =   2235
   End
   Begin VB.TextBox txt_attendance_allowance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6180
      TabIndex        =   68
      Top             =   2880
      Width           =   2235
   End
   Begin VB.TextBox txt_6_sh_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   58
      Top             =   7590
      Width           =   765
   End
   Begin VB.TextBox txt_6_sh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   56
      Top             =   7590
      Width           =   735
   End
   Begin VB.TextBox txt_4_sh_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   54
      Top             =   7290
      Width           =   765
   End
   Begin VB.TextBox txt_4_sh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   52
      Top             =   7290
      Width           =   735
   End
   Begin VB.TextBox txt_4_h_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   49
      Top             =   6660
      Width           =   765
   End
   Begin VB.TextBox txt_4_h 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   47
      Top             =   6660
      Width           =   735
   End
   Begin VB.TextBox txt_3_h_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   45
      Top             =   6360
      Width           =   765
   End
   Begin VB.TextBox txt_3_h 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   43
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txt_2_h_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   41
      Top             =   6060
      Width           =   765
   End
   Begin VB.TextBox txt_2_h 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   39
      Top             =   6060
      Width           =   735
   End
   Begin VB.TextBox txt_2_wd_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   36
      Top             =   5430
      Width           =   765
   End
   Begin VB.TextBox txt_2_wd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   34
      Top             =   5430
      Width           =   735
   End
   Begin VB.TextBox txt_15_wd_value 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3060
      TabIndex        =   32
      Top             =   5130
      Width           =   765
   End
   Begin VB.TextBox txt_15_wd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   30
      Top             =   5130
      Width           =   735
   End
   Begin VB.TextBox txt_late_frequency 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   25
      Top             =   4110
      Width           =   795
   End
   Begin VB.TextBox txt_late 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   22
      Top             =   3810
      Width           =   795
   End
   Begin VB.TextBox txt_absent_leave 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   19
      Top             =   3510
      Width           =   795
   End
   Begin VB.TextBox txt_sick_leave 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   16
      Top             =   3210
      Width           =   795
   End
   Begin VB.TextBox txt_private_leave 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   13
      Top             =   2910
      Width           =   795
   End
   Begin VB.TextBox txt_annual_leave 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   10
      Top             =   2610
      Width           =   795
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10140
      TabIndex        =   106
      Top             =   7260
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
      MICON           =   "frm_trans_salary_dtl.frx":058A
      PICN            =   "frm_trans_salary_dtl.frx":05A6
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
      Left            =   9150
      TabIndex        =   107
      Top             =   7260
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
      MICON           =   "frm_trans_salary_dtl.frx":1638
      PICN            =   "frm_trans_salary_dtl.frx":1654
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbl_month 
      Height          =   255
      Left            =   10350
      TabIndex        =   109
      Top             =   750
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lbl_employee_code 
      Height          =   255
      Left            =   9510
      TabIndex        =   108
      Top             =   750
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label61 
      Caption         =   "Angsuran Koperasi"
      Height          =   225
      Left            =   8760
      TabIndex        =   104
      Top             =   5940
      Width           =   1635
   End
   Begin VB.Label Label60 
      Caption         =   "Iuran Koperasi"
      Height          =   225
      Left            =   8760
      TabIndex        =   102
      Top             =   5220
      Width           =   1635
   End
   Begin VB.Label Label59 
      Caption         =   "Tax Correction"
      Height          =   225
      Left            =   8760
      TabIndex        =   100
      Top             =   4500
      Width           =   1635
   End
   Begin VB.Label lbl_actual_received 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   9240
      TabIndex        =   99
      Top             =   3930
      Width           =   2235
   End
   Begin VB.Label Label58 
      Caption         =   "Actual Received"
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
      Left            =   8760
      TabIndex        =   98
      Top             =   3630
      Width           =   2085
   End
   Begin VB.Label lbl_income_tax 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   9240
      TabIndex        =   97
      Top             =   3270
      Width           =   2235
   End
   Begin VB.Label Label57 
      Caption         =   "Income Tax"
      Height          =   225
      Left            =   8760
      TabIndex        =   96
      Top             =   2970
      Width           =   1635
   End
   Begin VB.Label lbl_jamsostek 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   9240
      TabIndex        =   95
      Top             =   2610
      Width           =   2235
   End
   Begin VB.Label Label56 
      Caption         =   "Jamsostek"
      Height          =   225
      Left            =   8760
      TabIndex        =   94
      Top             =   2310
      Width           =   1635
   End
   Begin VB.Label lbl_total_received 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   9240
      TabIndex        =   93
      Top             =   1950
      Width           =   2235
   End
   Begin VB.Label Label55 
      Caption         =   "Total Received"
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
      Left            =   8760
      TabIndex        =   92
      Top             =   1650
      Width           =   2085
   End
   Begin VB.Line Line2 
      X1              =   8580
      X2              =   8580
      Y1              =   1410
      Y2              =   7920
   End
   Begin VB.Label Label54 
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
      Height          =   225
      Left            =   4170
      TabIndex        =   90
      Top             =   5610
      Width           =   2085
   End
   Begin VB.Label Label53 
      Caption         =   "Days"
      Height          =   225
      Left            =   6690
      TabIndex        =   88
      Top             =   5190
      Width           =   435
   End
   Begin VB.Label Label52 
      Caption         =   "Night Time"
      Height          =   225
      Left            =   4290
      TabIndex        =   86
      Top             =   5190
      Width           =   1635
   End
   Begin VB.Label Label51 
      Caption         =   "Days"
      Height          =   225
      Left            =   6690
      TabIndex        =   84
      Top             =   4890
      Width           =   435
   End
   Begin VB.Label Label50 
      Caption         =   "Afternoon"
      Height          =   225
      Left            =   4290
      TabIndex        =   82
      Top             =   4890
      Width           =   1635
   End
   Begin VB.Label Label49 
      Caption         =   "Shift Allowance"
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
      Left            =   4200
      TabIndex        =   81
      Top             =   4560
      Width           =   2085
   End
   Begin VB.Label Label48 
      Caption         =   "Days"
      Height          =   225
      Left            =   6720
      TabIndex        =   79
      Top             =   4140
      Width           =   435
   End
   Begin VB.Label Label47 
      Caption         =   "Transport Allowance"
      Height          =   225
      Left            =   4320
      TabIndex        =   77
      Top             =   4140
      Width           =   1635
   End
   Begin VB.Label Label46 
      Caption         =   "Days"
      Height          =   225
      Left            =   6720
      TabIndex        =   75
      Top             =   3840
      Width           =   435
   End
   Begin VB.Label Label45 
      Caption         =   "Meal Allowance"
      Height          =   225
      Left            =   4320
      TabIndex        =   73
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Label Label44 
      Caption         =   "Other Allowance"
      Height          =   225
      Left            =   4320
      TabIndex        =   71
      Top             =   3510
      Width           =   1635
   End
   Begin VB.Label Label43 
      Caption         =   "Position Allowance"
      Height          =   225
      Left            =   4320
      TabIndex        =   69
      Top             =   3210
      Width           =   1635
   End
   Begin VB.Label Label41 
      Caption         =   "Attendance Allowance"
      Height          =   225
      Left            =   4320
      TabIndex        =   67
      Top             =   2910
      Width           =   1635
   End
   Begin VB.Label Label40 
      Caption         =   "Allowances :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4200
      TabIndex        =   66
      Top             =   2580
      Width           =   2835
   End
   Begin VB.Label lbl_bonus 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   6180
      TabIndex        =   65
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Label Label42 
      Caption         =   "Bonus/Gratification"
      Height          =   225
      Left            =   4320
      TabIndex        =   64
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label lbl_incentive 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   6180
      TabIndex        =   63
      Top             =   1860
      Width           =   2235
   End
   Begin VB.Label lbl_adjusted_basic 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   6180
      TabIndex        =   62
      Top             =   1560
      Width           =   2235
   End
   Begin VB.Label Label39 
      Caption         =   "Incentive"
      Height          =   225
      Left            =   4320
      TabIndex        =   61
      Top             =   1860
      Width           =   1635
   End
   Begin VB.Label Label38 
      Caption         =   "Adjusted Basic"
      Height          =   225
      Left            =   4320
      TabIndex        =   60
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label Label37 
      Caption         =   "Actual Received Income :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4170
      TabIndex        =   59
      Top             =   1230
      Width           =   2835
   End
   Begin VB.Line Line1 
      X1              =   4020
      X2              =   4020
      Y1              =   1440
      Y2              =   7950
   End
   Begin VB.Label Label36 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   57
      Top             =   7620
      Width           =   765
   End
   Begin VB.Label Label35 
      Caption         =   "X 6"
      Height          =   225
      Left            =   480
      TabIndex        =   55
      Top             =   7620
      Width           =   585
   End
   Begin VB.Label Label34 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   53
      Top             =   7320
      Width           =   765
   End
   Begin VB.Label Label33 
      Caption         =   "X 4"
      Height          =   225
      Left            =   480
      TabIndex        =   51
      Top             =   7320
      Width           =   585
   End
   Begin VB.Label Label32 
      Caption         =   "Special Holiday"
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
      Left            =   300
      TabIndex        =   50
      Top             =   7020
      Width           =   2085
   End
   Begin VB.Label Label31 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   48
      Top             =   6690
      Width           =   765
   End
   Begin VB.Label Label30 
      Caption         =   "X 4"
      Height          =   225
      Left            =   480
      TabIndex        =   46
      Top             =   6690
      Width           =   585
   End
   Begin VB.Label Label29 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   44
      Top             =   6390
      Width           =   765
   End
   Begin VB.Label Label28 
      Caption         =   "X 3"
      Height          =   225
      Left            =   480
      TabIndex        =   42
      Top             =   6390
      Width           =   585
   End
   Begin VB.Label Label27 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   40
      Top             =   6090
      Width           =   765
   End
   Begin VB.Label Label26 
      Caption         =   "X 2"
      Height          =   225
      Left            =   480
      TabIndex        =   38
      Top             =   6090
      Width           =   585
   End
   Begin VB.Label Label25 
      Caption         =   "Holiday"
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
      Left            =   300
      TabIndex        =   37
      Top             =   5790
      Width           =   2085
   End
   Begin VB.Label Label24 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   35
      Top             =   5460
      Width           =   765
   End
   Begin VB.Label Label23 
      Caption         =   "X 2"
      Height          =   225
      Left            =   480
      TabIndex        =   33
      Top             =   5460
      Width           =   585
   End
   Begin VB.Label Label22 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   2040
      TabIndex        =   31
      Top             =   5160
      Width           =   765
   End
   Begin VB.Label Label21 
      Caption         =   "X 1.5"
      Height          =   225
      Left            =   480
      TabIndex        =   29
      Top             =   5160
      Width           =   585
   End
   Begin VB.Label Label20 
      Caption         =   "Normal Working Day"
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
      Left            =   300
      TabIndex        =   28
      Top             =   4860
      Width           =   2085
   End
   Begin VB.Label Label19 
      Caption         =   "OVERTIME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   27
      Top             =   4560
      Width           =   2085
   End
   Begin VB.Label Label18 
      Caption         =   "Times"
      Height          =   225
      Left            =   3060
      TabIndex        =   26
      Top             =   4140
      Width           =   765
   End
   Begin VB.Label Label17 
      Caption         =   "Late Frequency"
      Height          =   225
      Left            =   330
      TabIndex        =   24
      Top             =   4140
      Width           =   1635
   End
   Begin VB.Label Label16 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   3060
      TabIndex        =   23
      Top             =   3840
      Width           =   765
   End
   Begin VB.Label Label15 
      Caption         =   "Late"
      Height          =   225
      Left            =   330
      TabIndex        =   21
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Label Label14 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   3060
      TabIndex        =   20
      Top             =   3540
      Width           =   765
   End
   Begin VB.Label Label13 
      Caption         =   "Absent Leave"
      Height          =   225
      Left            =   330
      TabIndex        =   18
      Top             =   3540
      Width           =   1635
   End
   Begin VB.Label Label11 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   3060
      TabIndex        =   17
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label Label10 
      Caption         =   "Sick Leave"
      Height          =   225
      Left            =   330
      TabIndex        =   15
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label Label9 
      Caption         =   "Hour(s)"
      Height          =   225
      Left            =   3060
      TabIndex        =   14
      Top             =   2940
      Width           =   765
   End
   Begin VB.Label Label8 
      Caption         =   "Private Leave"
      Height          =   225
      Left            =   330
      TabIndex        =   12
      Top             =   2940
      Width           =   1635
   End
   Begin VB.Label Label7 
      Caption         =   "Day(s)"
      Height          =   225
      Left            =   3060
      TabIndex        =   11
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label6 
      Caption         =   "Annual Leave"
      Height          =   225
      Left            =   330
      TabIndex        =   9
      Top             =   2640
      Width           =   1635
   End
   Begin VB.Label Label5 
      Caption         =   "Leave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   2310
      Width           =   2085
   End
   Begin VB.Label lbl_jk_jkk 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   2220
      TabIndex        =   7
      Top             =   1890
      Width           =   1635
   End
   Begin VB.Label lbl_basic_salary 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   225
      Left            =   2220
      TabIndex        =   6
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "JK and JKK"
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   1890
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "Monthly Basic Salary"
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Basic Information :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   3
      Top             =   1230
      Width           =   2085
   End
   Begin VB.Label lbl_employee_name 
      Enabled         =   0   'False
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
      Left            =   1560
      TabIndex        =   2
      Top             =   750
      Width           =   7725
   End
   Begin VB.Label Label1 
      Caption         =   "Detail Salary for"
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY PAYMENT"
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
      TabIndex        =   0
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_salary_dtl.frx":26E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14760
   End
End
Attribute VB_Name = "frm_trans_salary_dtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSalary As New ADODB.Recordset
Dim vTgl As String
Dim vBasicSalary As Double
Dim vIntWH As Double
Dim vPPh21 As String
Dim vPTKP As String
Dim vBiayaJabatan As Double
Dim vBruto As Double
Dim vNetto As Double

Dim vPrivateLeave As Double
Dim vAbsentLeave As Double
Dim vLate As Double
Dim vTotGajiBersih As Double

Dim vMarital As Integer
Dim vSex As Integer
Dim vChildren As Integer

Dim vOT_15 As Double
Dim vOT_2 As Double
Dim vOT_3 As Double
Dim vOT_4 As Double
Dim vOT_6 As Double
Dim vTotOT As Double

Dim vStartWorking As String
Dim vNettoSetahun As Double
Dim vPTKP_Value As Double
Dim vPKP As Double

Dim vPPh5 As Double
Dim vPPh15 As Double
Dim vPPh25 As Double
Dim vPPh30 As Double
Dim vPPh21Setahun As Double
Dim vPPh21_Value As Double

Dim vTotPenaltiHours As Double
Dim vProsenPenalty As Double
Dim vProsenPresenceAllow As Double

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err
    CnG.BeginTrans
    SQL = "UPDATE h_salary SET flag_manual = 1 " & _
            "WHERE employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND month = '" & lbl_month.Caption & "'"
    CnG.Execute SQL
    
    SQL = "DELETE FROM h_salary_manual WHERE employee_code = '" & lbl_employee_code.Caption & "' " & _
            "AND month = '" & lbl_month.Caption & "'"
    CnG.Execute SQL
    
    SQL = "INSERT INTO h_salary_manual " & _
          "VALUES ( " & _
            "'" & lbl_month.Caption & "','" & lbl_employee_code.Caption & "'," & _
            "'" & txt_annual_leave.Text & "','" & txt_private_leave.Text & "'," & _
            "'" & txt_sick_leave.Text & "','" & txt_absent_leave.Text & "'," & _
            "'" & txt_late.Text & "','" & txt_late_frequency & "'," & _
            "'" & txt_15_wd.Text & "','" & txt_2_wd.Text & "'," & _
            "'" & txt_2_h.Text & "','" & txt_3_h.Text & "','" & txt_4_h.Text & "'," & _
            "'" & txt_4_sh.Text & "','" & txt_6_sh.Text & "'," & _
            "'" & DropAllComma(txt_attendance_allowance.Text) & "','" & DropAllComma(txt_other_allowance.Text) & "'," & _
            "'" & txt_meal_days.Text & "','" & txt_transport_days.Text & "'," & _
            "'" & txt_afternoon_days.Text & "','" & txt_night_days & "'," & _
            "'" & DropAllComma(txt_tax_correction.Text) & "','" & DropAllComma(txt_iuran_koperasi.Text) & "'," & _
            "'" & DropAllComma(txt_angsuran_koperasi.Text) & "','" & Replace(txt_description.Text, "'", "''") & "'," & _
            "Now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL
    
    vLate = txt_late.Text * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vAbsentLeave = txt_absent_leave.Text * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vPrivateLeave = txt_private_leave.Text * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vPrivateLeave = vPrivateLeave + vAbsentLeave
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    
    vTotGajiBersih = Val(DropAllComma(lbl_total_received.Caption)) + Val(DropAllComma(lbl_jk_jkk.Caption))
    
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_attendance_allowance.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-020'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_meal_allowance.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-023'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_transport_allowance.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-024'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(lbl_incentive.Caption)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-025'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vOT_15)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-0251'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vOT_2)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-0252'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vOT_3)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-0253'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vOT_4)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-0254'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vOT_6)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-0255'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_afternoon_allowance.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-026'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_night_allowance.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-027'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_other_allowance.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-028'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPrivateLeave)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-0711'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_angsuran_koperasi.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-075'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_iuran_koperasi.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-076'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vLate)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-077'"
    
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vTotGajiBersih)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-28'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vBiayaJabatan)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-286'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vNetto)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-2891'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vNettoSetahun)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-29'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPTKP)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-30'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPKP)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-31'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPPh5)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-32'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPPh15)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-33'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPPh25)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-34'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPPh30)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-35'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPPh21Setahun)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-35A'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(vPPh21_Value)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-36'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(lbl_actual_received.Caption)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-37'"
    CnG.Execute "UPDATE h_salary SET salary_value = '" & Val(DropAllComma(txt_tax_correction.Text)) & "' " & _
                "WHERE employee_code = '" & lbl_employee_code.Caption & "' AND month = '" & lbl_month.Caption & "' AND salary_code = 'SU-037'"
    CnG.CommitTrans
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub Form_Load()
    
    vPrivateLeave = 0
    vAbsentLeave = 0
    vLate = 0
    
    If rsSalary.State Then rsSalary.Close
    SQL = "CALL spr_list_salary_detail('" & Format(frm_list_salary.TDBGrid1.Columns("date_from").Value, "yyyy-MM-dd") & "'," & _
            "'" & Format(frm_list_salary.TDBGrid1.Columns("date_to").Value, "yyyy-MM-dd") & "'," & _
            "'" & frm_list_salary.TDBGrid1.Columns("month").Value & "'," & _
            "'" & frm_list_salary.TDBGrid1.Columns("employee_code").Value & "')"
    rsSalary.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rsSalary.RecordCount > 0 Then
        With rsSalary
            lbl_month.Caption = UCase(.Fields("month").Value)
            lbl_employee_code.Caption = UCase(.Fields("employee_code").Value)
            
            lbl_employee_name.Caption = UCase(.Fields("employee_name").Value)
            lbl_basic_salary.Caption = FormatNumber(.Fields("basic_salary").Value)
            lbl_jk_jkk.Caption = FormatNumber(.Fields("jk_jkk").Value)
            
            txt_annual_leave.Text = FormatNumber(.Fields("days_leave").Value)
            txt_private_leave.Text = FormatNumber(.Fields("private_leave").Value)
            txt_sick_leave.Text = FormatNumber(.Fields("sick_leave").Value)
            txt_absent_leave.Text = FormatNumber(.Fields("absent_leave").Value)
            txt_late.Text = FormatNumber(.Fields("sum_late").Value)
            txt_late_frequency = FormatNumber(.Fields("count_late").Value)
            
            txt_15_wd.Text = FormatNumber(.Fields("ot15").Value)
            txt_15_wd_value.Text = FormatNumber(.Fields("ot15").Value * 1.5)
            txt_2_wd.Text = FormatNumber(.Fields("ot20").Value)
            txt_2_wd_value.Text = FormatNumber(.Fields("ot20").Value * 2)
            
            txt_2_h.Text = FormatNumber(.Fields("ot20_hol").Value)
            txt_2_h_value.Text = FormatNumber(.Fields("ot20_hol").Value * 2)
            txt_3_h.Text = FormatNumber(.Fields("ot30_hol").Value)
            txt_3_h_value.Text = FormatNumber(.Fields("ot30_hol").Value * 3)
            txt_4_h.Text = FormatNumber(.Fields("ot40_hol").Value)
            txt_4_h_value.Text = FormatNumber(.Fields("ot40_hol").Value * 4)
            
            txt_4_sh.Text = FormatNumber(.Fields("ot40_hr").Value)
            txt_4_sh_value.Text = FormatNumber(.Fields("ot40_hr").Value * 4)
            txt_6_sh.Text = FormatNumber(.Fields("ot60_hr").Value)
            txt_6_sh_value.Text = FormatNumber(.Fields("ot60_hr").Value * 6)
            
            lbl_adjusted_basic.Caption = FormatNumber(.Fields("received_salary").Value)
            lbl_incentive.Caption = FormatNumber(.Fields("overtime").Value)
            lbl_bonus.Caption = FormatNumber(.Fields("bonus").Value)
            
            txt_attendance_allowance.Text = FormatNumber(.Fields("attendant").Value)
            txt_position_allowance.Text = FormatNumber(.Fields("position_allow").Value)
            txt_other_allowance.Text = FormatNumber(.Fields("other_allow").Value)
            txt_meal_days.Text = FormatNumber(.Fields("meal_days").Value)
            txt_meal_allowance.Text = FormatNumber(.Fields("meal_allow").Value)
            txt_transport_days.Text = FormatNumber(.Fields("transport_days").Value)
            txt_transport_allowance.Text = FormatNumber(.Fields("transport_allow").Value)
            
            txt_afternoon_days.Text = FormatNumber(.Fields("afternoon_days").Value)
            txt_afternoon_allowance.Text = FormatNumber(.Fields("afternoon_allow").Value)
            txt_night_days.Text = FormatNumber(.Fields("night_days").Value)
            txt_night_allowance.Text = FormatNumber(.Fields("night_allow").Value)
            
            txt_description.Text = ""
            
            lbl_total_received.Caption = FormatNumber(.Fields("gross_income").Value)
            lbl_jamsostek.Caption = FormatNumber(.Fields("jms").Value)
            lbl_income_tax.Caption = IIf(.Fields("income_tax").Value < 0, 0, FormatNumber(.Fields("income_tax").Value))
            lbl_actual_received.Caption = FormatNumber(.Fields("actual").Value)
            
            txt_tax_correction.Text = FormatNumber(.Fields("tax_correction").Value)
            txt_iuran_koperasi.Text = FormatNumber(.Fields("coop_contr").Value)
            txt_angsuran_koperasi.Text = FormatNumber(.Fields("coop_install").Value)
            
            txt_description.Text = .Fields("description").Value
            
            vMarital = .Fields("marital_status").Value
            vSex = .Fields("sex").Value
            vChildren = .Fields("no_of_children").Value
        End With
    End If
    rsSalary.Close
    
    vTgl = lbl_month.Caption & "-20"
    SQL = "(SELECT IFNULL(basic_salary,0) basic_salary, pph21_type, ptkp_type " & _
                    "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                   "Where a.employee_code = '" & lbl_employee_code & "' And a.salary_date <= '" & vTgl & "' " & _
                   "ORDER BY a.salary_date DESC LIMIT 1)"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vBasicSalary = rscari!basic_salary
        vPPh21 = rscari!pph21_type
        vPTKP = rscari!ptkp_type
    End If
    rscari.Close
    
    SQL = "(SELECT IFNULL(wh_value,0) wh_value FROM m_pref_att);"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vIntWH = rscari!wh_value
    End If
    rscari.Close
End Sub

Private Sub lbl_adjusted_basic_Change()
On Error Resume Next
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub lbl_incentive_Change()
On Error Resume Next
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub lbl_jamsostek_Change()
On Error Resume Next
    lbl_actual_received.Caption = Val(DropAllComma(lbl_total_received.Caption)) - Val(DropAllComma(lbl_jamsostek.Caption)) _
                                  - Val(DropAllComma(lbl_income_tax.Caption)) - Val(DropAllComma(txt_tax_correction.Text)) _
                                  - Val(DropAllComma(txt_iuran_koperasi.Text)) - Val(DropAllComma(txt_angsuran_koperasi))
    lbl_actual_received.Caption = FormatNumber(lbl_actual_received.Caption)
End Sub

Private Sub lbl_income_tax_Change()
On Error Resume Next
    lbl_actual_received.Caption = Val(DropAllComma(lbl_total_received.Caption)) - Val(DropAllComma(lbl_jamsostek.Caption)) _
                                  - Val(DropAllComma(lbl_income_tax.Caption)) - Val(DropAllComma(txt_tax_correction.Text)) _
                                  - Val(DropAllComma(txt_iuran_koperasi.Text)) - Val(DropAllComma(txt_angsuran_koperasi))
    lbl_actual_received.Caption = FormatNumber(lbl_actual_received.Caption)
End Sub

Private Sub lbl_total_received_Change()
On Error Resume Next
    lbl_actual_received.Caption = Val(DropAllComma(lbl_total_received.Caption)) - Val(DropAllComma(lbl_jamsostek.Caption)) _
                                  - Val(DropAllComma(lbl_income_tax.Caption)) - Val(DropAllComma(txt_tax_correction.Text)) _
                                  - Val(DropAllComma(txt_iuran_koperasi.Text)) - Val(DropAllComma(txt_angsuran_koperasi))
    lbl_actual_received.Caption = FormatNumber(lbl_actual_received.Caption)
End Sub

Private Sub txt_tax_correction_LostFocus()
On Error Resume Next
    txt_tax_correction.Text = FormatNumber(txt_tax_correction.Text)
    lbl_actual_received.Caption = Val(DropAllComma(lbl_total_received.Caption)) - Val(DropAllComma(lbl_jamsostek.Caption)) _
                                  - Val(DropAllComma(lbl_income_tax.Caption)) - Val(DropAllComma(txt_tax_correction.Text)) _
                                  - Val(DropAllComma(txt_iuran_koperasi.Text)) - Val(DropAllComma(txt_angsuran_koperasi))
    lbl_actual_received.Caption = FormatNumber(lbl_actual_received.Caption)
End Sub

Private Sub txt_iuran_koperasi_LostFocus()
On Error Resume Next
    txt_iuran_koperasi.Text = FormatNumber(txt_iuran_koperasi.Text)
    lbl_actual_received.Caption = Val(DropAllComma(lbl_total_received.Caption)) - Val(DropAllComma(lbl_jamsostek.Caption)) _
                                  - Val(DropAllComma(lbl_income_tax.Caption)) - Val(DropAllComma(txt_tax_correction.Text)) _
                                  - Val(DropAllComma(txt_iuran_koperasi.Text)) - Val(DropAllComma(txt_angsuran_koperasi))
    lbl_actual_received.Caption = FormatNumber(lbl_actual_received.Caption)
End Sub

Private Sub txt_angsuran_koperasi_LostFocus()
On Error Resume Next
    txt_angsuran_koperasi.Text = FormatNumber(txt_angsuran_koperasi.Text)
    lbl_actual_received.Caption = Val(DropAllComma(lbl_total_received.Caption)) - Val(DropAllComma(lbl_jamsostek.Caption)) _
                                  - Val(DropAllComma(lbl_income_tax.Caption)) - Val(DropAllComma(txt_tax_correction.Text)) _
                                  - Val(DropAllComma(txt_iuran_koperasi.Text)) - Val(DropAllComma(txt_angsuran_koperasi))
    lbl_actual_received.Caption = FormatNumber(lbl_actual_received.Caption)
End Sub

Private Sub txt_annual_leave_LostFocus()
    txt_annual_leave.Text = FormatNumber(txt_annual_leave.Text)
End Sub

Private Sub txt_sick_leave_LostFocus()
    txt_sick_leave.Text = txt_sick_leave.Text
End Sub

Private Sub txt_late_frequency_LostFocus()
    txt_late_frequency.Text = txt_late_frequency.Text
End Sub

Private Sub txt_attendance_allowance_LostFocus()
On Error Resume Next
    txt_attendance_allowance.Text = FormatNumber(txt_attendance_allowance.Text)
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_other_allowance_LostFocus()
On Error Resume Next
    txt_other_allowance.Text = FormatNumber(txt_other_allowance.Text)
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_meal_allowance_LostFocus()
On Error Resume Next
    txt_meal_allowance.Text = FormatNumber(txt_meal_allowance.Text)
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_transport_allowance_LostFocus()
On Error Resume Next
    txt_transport_allowance = FormatNumber(txt_transport_allowance.Text)
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_afternoon_allowance_LostFocus()
On Error Resume Next
    txt_afternoon_allowance.Text = FormatNumber(txt_afternoon_allowance.Text)
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_night_allowance_LostFocus()
On Error Resume Next
    txt_night_allowance.Text = FormatNumber(txt_night_allowance.Text)
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_15_wd_LostFocus()
On Error Resume Next
    txt_15_wd.Text = FormatNumber(txt_15_wd.Text)
    txt_15_wd_value.Text = FormatNumber(txt_15_wd.Text * 1.5)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_2_wd_LostFocus()
On Error Resume Next
    txt_2_wd.Text = FormatNumber(txt_2_wd.Text)
    txt_2_wd_value.Text = FormatNumber(txt_2_wd.Text * 2)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_2_h_LostFocus()
On Error Resume Next
    txt_2_h.Text = FormatNumber(txt_2_h.Text)
    txt_2_h_value.Text = FormatNumber(txt_2_h.Text * 2)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_3_h_LostFocus()
On Error Resume Next
    txt_3_h.Text = FormatNumber(txt_3_h.Text)
    txt_3_h_value.Text = FormatNumber(txt_3_h.Text * 3)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_4_h_LostFocus()
On Error Resume Next
    txt_4_h.Text = FormatNumber(txt_4_h.Text)
    txt_4_h_value.Text = FormatNumber(txt_4_h.Text * 4)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_4_sh_LostFocus()
On Error Resume Next
    txt_4_sh.Text = FormatNumber(txt_4_sh.Text)
    txt_4_sh_value.Text = FormatNumber(txt_4_sh.Text * 4)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_6_sh_LostFocus()
On Error Resume Next
    txt_6_sh.Text = FormatNumber(txt_6_sh.Text)
    txt_6_sh_value.Text = FormatNumber(txt_6_sh.Text * 6)
    
    vOT_15 = Val(DropAllComma(txt_15_wd_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_2 = (Val(DropAllComma(txt_2_wd_value.Text)) + Val(DropAllComma(txt_2_h_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_3 = Val(DropAllComma(txt_3_h_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vOT_4 = (Val(DropAllComma(txt_4_h_value.Text)) + Val(DropAllComma(txt_4_sh_value.Text))) * Round(vBasicSalary / vIntWH, 0)
    vOT_6 = Val(DropAllComma(txt_6_sh_value.Text)) * Round(vBasicSalary / vIntWH, 0)
    vTotOT = vOT_15 + vOT_2 + vOT_3 + vOT_4 + vOT_6
    lbl_incentive.Caption = FormatNumber(vTotOT)
    
    Call hitung_pph
End Sub

Private Sub txt_meal_days_LostFocus()
    vTgl = lbl_month.Caption & "-20"
    txt_meal_days.Text = FormatNumber(txt_meal_days.Text)
    
    SQL = "SELECT meal_allowance FROM m_salary_standard " & _
            "WHERE employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND date(salary_date) <= '" & vTgl & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        txt_meal_allowance.Text = FormatNumber(txt_meal_days.Text * rscari!meal_allowance)
    End If
    rscari.Close
    
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)


    Call hitung_pph
End Sub

Private Sub txt_private_leave_LostFocus()
On Error Resume Next
    txt_private_leave.Text = txt_private_leave.Text
    
    vLate = Val(txt_late.Text) * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vAbsentLeave = Val(txt_absent_leave.Text) * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vPrivateLeave = Val(txt_private_leave.Text) * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    lbl_adjusted_basic.Caption = FormatNumber(Round(Val(DropAllComma(lbl_basic_salary.Caption)) - vPrivateLeave - vAbsentLeave - vLate), 0)
    
    SQL = "(SELECT presence_allowance " & _
            "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
            "WHERE a.employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND a.salary_date <= '" & Format(lbl_month.Caption, "yyyy-MM-20") & "' " & _
            "ORDER BY a.salary_date DESC LIMIT 1)"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vProsenPresenceAllow = rs!presence_allowance
    End If
    rs.Close
    
    vTotPenaltiHours = Val(txt_absent_leave.Text) + Val(txt_private_leave.Text)
    
    SQL = "SELECT percentage FROM m_pref_preall " & _
          "WHERE limit_hours <= '" & vTotPenaltiHours & "' " & _
          "ORDER BY id_number DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vProsenPenalty = rs!percentage
    Else
        vProsenPenalty = 0
    End If
    rs.Close
    
    txt_attendance_allowance.Text = ((vProsenPresenceAllow / 100) * Val(DropAllComma(lbl_basic_salary.Caption))) - _
                                    ((vProsenPenalty / 100) * ((vProsenPresenceAllow / 100) * Val(DropAllComma(lbl_basic_salary.Caption))))
    txt_attendance_allowance.Text = FormatNumber(txt_attendance_allowance.Text)
    
    Call hitung_pph
End Sub

Private Sub txt_absent_leave_LostFocus()
On Error Resume Next
    txt_absent_leave.Text = txt_absent_leave.Text
    
    vLate = Val(txt_late.Text) * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vPrivateLeave = Val(txt_private_leave.Text) * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vAbsentLeave = Val(txt_absent_leave.Text) * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    lbl_adjusted_basic.Caption = FormatNumber(Round(Val(DropAllComma(lbl_basic_salary.Caption)) - vPrivateLeave - vAbsentLeave - vLate), 0)
    
    SQL = "(SELECT presence_allowance " & _
            "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
            "WHERE a.employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND a.salary_date <= '" & Format(lbl_month.Caption, "yyyy-MM-20") & "' " & _
            "ORDER BY a.salary_date DESC LIMIT 1)"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vProsenPresenceAllow = rs!presence_allowance
    End If
    rs.Close
    
    vTotPenaltiHours = Val(txt_absent_leave.Text) + Val(txt_private_leave.Text)
    
    SQL = "SELECT percentage FROM m_pref_preall " & _
          "WHERE limit_hours <= '" & vTotPenaltiHours & "' " & _
          "ORDER BY id_number DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vProsenPenalty = rs!percentage
    Else
        vProsenPenalty = 0
    End If
    rs.Close
    
    txt_attendance_allowance.Text = ((vProsenPresenceAllow / 100) * Val(DropAllComma(lbl_basic_salary.Caption))) - _
                                    ((vProsenPenalty / 100) * ((vProsenPresenceAllow / 100) * Val(DropAllComma(lbl_basic_salary.Caption))))
    txt_attendance_allowance.Text = FormatNumber(txt_attendance_allowance.Text)
    
    Call hitung_pph
End Sub

Private Sub txt_late_LostFocus()
On Error Resume Next
    txt_late.Text = txt_late.Text
    
    vPrivateLeave = txt_private_leave.Text * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vAbsentLeave = txt_absent_leave.Text * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    vLate = txt_late.Text * Round(vBasicSalary * Round(1 / vIntWH, 6), 0)
    lbl_adjusted_basic.Caption = FormatNumber(Round(Val(DropAllComma(lbl_basic_salary.Caption)) - vPrivateLeave - vAbsentLeave - vLate), 0)
    
    Call hitung_pph
End Sub

Private Sub txt_transport_days_LostFocus()
    vTgl = lbl_month.Caption & "-20"
    txt_transport_days.Text = FormatNumber(txt_transport_days.Text)
    
    SQL = "SELECT transport_allowance FROM m_salary_standard " & _
            "WHERE employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND date(salary_date) <= '" & vTgl & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        txt_transport_allowance.Text = FormatNumber(txt_transport_days.Text * rscari!transport_allowance)
    End If
    rscari.Close
    
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)

    Call hitung_pph
End Sub

Private Sub txt_afternoon_days_LostFocus()
    vTgl = lbl_month.Caption & "-20"
    txt_afternoon_days.Text = FormatNumber(txt_afternoon_days.Text)
    
    SQL = "SELECT shift2_allowance FROM m_salary_standard " & _
            "WHERE employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND date(salary_date) <= '" & vTgl & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        txt_afternoon_allowance.Text = FormatNumber(txt_afternoon_days.Text * rscari!shift2_allowance)
    End If
    rscari.Close
    
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub txt_night_days_LostFocus()
    vTgl = lbl_month.Caption & "-20"
    txt_night_days.Text = FormatNumber(txt_night_days.Text)
    
    SQL = "SELECT shift3_allowance FROM m_salary_standard " & _
            "WHERE employee_code = '" & lbl_employee_code.Caption & "' " & _
                "AND date(salary_date) <= '" & vTgl & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        txt_night_allowance.Text = FormatNumber(txt_night_days.Text * rscari!shift3_allowance)
    End If
    rscari.Close
    
    lbl_total_received.Caption = Val(DropAllComma(lbl_adjusted_basic.Caption)) + Val(DropAllComma(lbl_incentive.Caption)) + Val(DropAllComma(txt_attendance_allowance)) _
                                 + Val(DropAllComma(txt_position_allowance.Text)) + Val(DropAllComma(txt_other_allowance.Text)) + Val(DropAllComma(txt_meal_allowance.Text)) _
                                 + Val(DropAllComma(txt_transport_allowance.Text)) + Val(DropAllComma(txt_afternoon_allowance.Text) + Val(DropAllComma(txt_night_allowance.Text)))
    lbl_total_received.Caption = FormatNumber(lbl_total_received.Caption)
    
    Call hitung_pph
End Sub

Private Sub hitung_pph()
Dim rsPPh As New ADODB.Recordset
Dim a As Integer
    
    vBruto = Val(DropAllComma(lbl_total_received.Caption)) + Val(DropAllComma(lbl_jk_jkk.Caption))
    vBiayaJabatan = IIf((0.05 * vBruto) > 500000, 500000, (0.05 * vBruto))
    
    vNetto = vBruto - Val(DropAllComma(lbl_jamsostek.Caption)) _
             - vBiayaJabatan
    
    SQL = "SELECT start_working  " & _
            "FROM m_employee " & _
            "WHERE employee_code = '" & lbl_employee_code.Caption & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly
    
    If rscari.RecordCount > 0 Then
        vStartWorking = Format(rscari!start_working, "yyyy-MM-dd")
    End If
    rscari.Close
    
    a = (DateDiff("m", vStartWorking, year(Now) & "-12-31")) + 1
    If a < 12 And Format(vStartWorking, "yyyy") = Left(lbl_month.Caption, 4) Then
        vNettoSetahun = vNetto * ((DateDiff("m", vStartWorking, year(Now) & "-12-31")) + 1)
    Else
        vNettoSetahun = vNetto * 12
    End If
    
    SQL = "SELECT f_get_ptkp(" & vMarital & ", " & vChildren & "," & vSex & ", 1,'" & vPTKP & "') ptkp_value"
    rscari.Open SQL, CnG, adOpenForwardOnly
    
    If rscari.RecordCount > 0 Then
        vPTKP_Value = rscari!ptkp_value
    End If
    rscari.Close
    
    vPKP = vNettoSetahun - vPTKP_Value
    vPKP = Int(vPKP / 1000) * 1000
    
    If vPKP < 50000000 Then
        vPPh5 = 0.05 * vPKP
        vPPh15 = 0
        vPPh25 = 0
        vPPh30 = 0
    ElseIf vPKP > 50000000 And vPKP < 250000000 Then
        vPPh5 = 0.05 * 50000000
        vPPh15 = 0.15 * (vPKP - 50000000)
        vPPh25 = 0
        vPPh30 = 0
'        vPPh21Setahun = (0.05 * 50000000) + (0.15 * (vPKP - 50000000))
    ElseIf vPKP > 250000000 And vPKP < 500000000 Then
        vPPh5 = 0.05 * 50000000
        vPPh15 = 0.15 * 200000000
        vPPh25 = 0.25 * (vPKP - 250000000)
        vPPh30 = 0
'        vPPh21Setahun = (0.05 * 50000000) + (0.15 * 200000000) + (0.25 * (vPKP - 50000000))
    Else
        vPPh5 = 0.05 * 50000000
        vPPh15 = 0.15 * 200000000
        vPPh25 = 0.25 * 250000000
        vPPh30 = 0.35 * (vPKP - 500000000)
'        vPPh21Setahun = (0.05 * 50000000) + (0.15 * 200000000) + (0.25 * 250000000) + (0.35 * (vPKP - 500000000))
    End If
    vPPh21Setahun = vPPh5 + vPPh15 + vPPh25 + vPPh30
    
    If a < 12 And Format(vStartWorking, "yyyy") = Left(lbl_month.Caption, 4) Then
        vPPh21_Value = vPPh21Setahun / ((DateDiff("m", vStartWorking, year(Now) & "-12-31")) + 1)
    Else
        vPPh21_Value = vPPh21Setahun / 12
    End If
    
    If vPPh21_Value <= 0 Then
        lbl_income_tax.Caption = FormatNumber(0)
    Else
        lbl_income_tax.Caption = FormatNumber(Round(vPPh21_Value, 0))
    End If
End Sub
