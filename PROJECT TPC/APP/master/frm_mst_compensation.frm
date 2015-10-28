VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Begin VB.Form frm_mst_compensation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER COMPENSATION"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11760
   Icon            =   "frm_mst_compensation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   90
      Top             =   7440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "RESIGN"
      TabPicture(0)   =   "frm_mst_compensation.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PENSION"
      TabPicture(1)   =   "frm_mst_compensation.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "P H K"
      TabPicture(2)   =   "frm_mst_compensation.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -73140
         TabIndex        =   57
         Top             =   4170
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave 
            Height          =   705
            Index           =   2
            Left            =   5820
            TabIndex        =   24
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
            MICON           =   "frm_mst_compensation.frx":05DE
            PICN            =   "frm_mst_compensation.frx":05FA
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
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   -73140
         TabIndex        =   45
         Top             =   810
         Width           =   7965
         Begin VB.TextBox txt_masa_kerja_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   2
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   60
            Top             =   600
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox txt_pesangon_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   2
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   58
            Top             =   270
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.CheckBox chk_masa_kerja 
            Caption         =   "YES"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   16
            Top             =   600
            Width           =   705
         End
         Begin VB.CheckBox chk_pph21 
            Caption         =   "YES"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   22
            Top             =   2460
            Width           =   945
         End
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   2
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            Top             =   2760
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox chk_pesangon 
            Caption         =   "YES"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   15
            Top             =   300
            Width           =   735
         End
         Begin VB.CheckBox chk_leave 
            Caption         =   "YES"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   17
            Top             =   900
            Width           =   765
         End
         Begin VB.CheckBox chk_ganti 
            Caption         =   "YES"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   18
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox txt_ganti 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   3000
            TabIndex        =   19
            Top             =   1530
            Width           =   435
         End
         Begin VB.TextBox txt_kali_pesangon 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   3000
            TabIndex        =   20
            Top             =   1920
            Width           =   435
         End
         Begin VB.TextBox txt_kali_masa_kerja 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   4980
            TabIndex        =   21
            Top             =   1920
            Width           =   435
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
            Height          =   375
            Index           =   2
            Left            =   3000
            OleObjectBlob   =   "frm_mst_compensation.frx":168C
            TabIndex        =   23
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pesangon 
            Height          =   375
            Index           =   2
            Left            =   3780
            OleObjectBlob   =   "frm_mst_compensation.frx":35E9
            TabIndex        =   59
            Top             =   270
            Visible         =   0   'False
            Width           =   1305
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_masa_kerja 
            Height          =   375
            Index           =   2
            Left            =   3780
            OleObjectBlob   =   "frm_mst_compensation.frx":5557
            TabIndex        =   61
            Top             =   600
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "GET PENGHARGAAN MASA KERJA"
            Height          =   255
            Left            =   150
            TabIndex        =   56
            Top             =   630
            Width           =   2775
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "USE PPH 21"
            Height          =   225
            Left            =   1350
            TabIndex        =   55
            Top             =   2490
            Width           =   1575
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "GET PESANGON"
            Height          =   255
            Left            =   150
            TabIndex        =   54
            Top             =   330
            Width           =   2775
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "CONVERT AVAILABLE LEAVE"
            Height          =   255
            Left            =   600
            TabIndex        =   53
            Top             =   930
            Width           =   2325
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "UANG PENGGANTI HAK"
            Height          =   255
            Left            =   600
            TabIndex        =   52
            Top             =   1230
            Width           =   2325
         End
         Begin VB.Label lblGanti 
            Caption         =   "% X (Uang Pesangom + Uang Penghargaan Masa Kerja)"
            Height          =   255
            Index           =   2
            Left            =   3450
            TabIndex        =   51
            Top             =   1560
            Width           =   4365
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "FORMULA"
            Height          =   255
            Left            =   570
            TabIndex        =   50
            Top             =   1980
            Width           =   2325
         End
         Begin VB.Label Label16 
            Caption         =   "X Uang Pesangon +     "
            Height          =   255
            Left            =   3450
            TabIndex        =   49
            Top             =   1950
            Width           =   1845
         End
         Begin VB.Label Label15 
            Caption         =   "X Uang Penghargaan Masa Kerja"
            Height          =   255
            Left            =   5460
            TabIndex        =   48
            Top             =   1950
            Width           =   2445
         End
         Begin VB.Label Label14 
            Caption         =   "+ Uang Penggantian Hak"
            Height          =   255
            Left            =   3030
            TabIndex        =   47
            Top             =   2220
            Width           =   2445
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   -73140
         TabIndex        =   33
         Top             =   810
         Width           =   7965
         Begin VB.TextBox txt_masa_kerja_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   1
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   64
            Top             =   600
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox txt_pesangon_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   1
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   62
            Top             =   270
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox txt_kali_masa_kerja 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4980
            TabIndex        =   11
            Top             =   1920
            Width           =   435
         End
         Begin VB.TextBox txt_kali_pesangon 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   10
            Top             =   1920
            Width           =   435
         End
         Begin VB.TextBox txt_ganti 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   9
            Top             =   1530
            Width           =   435
         End
         Begin VB.CheckBox chk_ganti 
            Caption         =   "YES"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   8
            Top             =   1200
            Width           =   945
         End
         Begin VB.CheckBox chk_leave 
            Caption         =   "YES"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   7
            Top             =   900
            Width           =   945
         End
         Begin VB.CheckBox chk_pesangon 
            Caption         =   "YES"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   5
            Top             =   300
            Width           =   675
         End
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   1
            Left            =   4620
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            Top             =   2760
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox chk_pph21 
            Caption         =   "YES"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   12
            Top             =   2460
            Width           =   945
         End
         Begin VB.CheckBox chk_masa_kerja 
            Caption         =   "YES"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   6
            Top             =   600
            Width           =   675
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
            Height          =   375
            Index           =   1
            Left            =   3000
            OleObjectBlob   =   "frm_mst_compensation.frx":74CB
            TabIndex        =   13
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pesangon 
            Height          =   375
            Index           =   1
            Left            =   3780
            OleObjectBlob   =   "frm_mst_compensation.frx":9428
            TabIndex        =   63
            Top             =   270
            Visible         =   0   'False
            Width           =   1305
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_masa_kerja 
            Height          =   375
            Index           =   1
            Left            =   3780
            OleObjectBlob   =   "frm_mst_compensation.frx":B396
            TabIndex        =   65
            Top             =   600
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label13 
            Caption         =   "+ Uang Penggantian Hak"
            Height          =   255
            Left            =   3030
            TabIndex        =   44
            Top             =   2220
            Width           =   2445
         End
         Begin VB.Label Label12 
            Caption         =   "X Uang Penghargaan Masa Kerja"
            Height          =   255
            Left            =   5460
            TabIndex        =   43
            Top             =   1950
            Width           =   2445
         End
         Begin VB.Label Label11 
            Caption         =   "X Uang Pesangon +     "
            Height          =   255
            Left            =   3450
            TabIndex        =   42
            Top             =   1950
            Width           =   1845
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "FORMULA"
            Height          =   255
            Left            =   570
            TabIndex        =   41
            Top             =   1980
            Width           =   2325
         End
         Begin VB.Label lblGanti 
            Caption         =   "% X (Uang Pesangom + Uang Penghargaan Masa Kerja)"
            Height          =   255
            Index           =   1
            Left            =   3450
            TabIndex        =   40
            Top             =   1560
            Width           =   4365
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "UANG PENGGANTI HAK"
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   1230
            Width           =   2325
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CONVERT AVAILABLE LEAVE"
            Height          =   255
            Left            =   600
            TabIndex        =   38
            Top             =   930
            Width           =   2325
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "GET PESANGON"
            Height          =   255
            Left            =   150
            TabIndex        =   37
            Top             =   330
            Width           =   2775
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "USE PPH 21"
            Height          =   225
            Left            =   1350
            TabIndex        =   36
            Top             =   2490
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "GET PENGHARGAAN MASA KERJA"
            Height          =   255
            Left            =   150
            TabIndex        =   35
            Top             =   630
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -73140
         TabIndex        =   32
         Top             =   4170
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave 
            Height          =   705
            Index           =   1
            Left            =   5820
            TabIndex        =   14
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
            MICON           =   "frm_mst_compensation.frx":D30A
            PICN            =   "frm_mst_compensation.frx":D326
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
      Begin VB.Frame Frame4 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   1860
         TabIndex        =   27
         Top             =   4170
         Width           =   7965
         Begin prj_tpc.vbButton cmdSave 
            Height          =   705
            Index           =   0
            Left            =   5820
            TabIndex        =   3
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
            MICON           =   "frm_mst_compensation.frx":E3B8
            PICN            =   "frm_mst_compensation.frx":E3D4
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
      Begin VB.Frame Frame8 
         Height          =   3255
         Left            =   1860
         TabIndex        =   25
         Top             =   810
         Width           =   7965
         Begin VB.TextBox txt_pesangon_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   0
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   66
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chk_resign_leave 
            Caption         =   "YES"
            Height          =   255
            Left            =   2700
            TabIndex        =   0
            Top             =   1110
            Width           =   945
         End
         Begin VB.CheckBox chk_resign_pph 
            Caption         =   "YES"
            Height          =   255
            Left            =   2700
            TabIndex        =   1
            Top             =   1410
            Width           =   945
         End
         Begin VB.TextBox txt_pph21_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Index           =   0
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1740
            Visible         =   0   'False
            Width           =   2535
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
            Height          =   375
            Index           =   0
            Left            =   2700
            OleObjectBlob   =   "frm_mst_compensation.frx":F466
            TabIndex        =   2
            Top             =   1740
            Visible         =   0   'False
            Width           =   1575
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_pesangon 
            Height          =   375
            Index           =   0
            Left            =   2700
            OleObjectBlob   =   "frm_mst_compensation.frx":113C3
            TabIndex        =   67
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "CONVERT AVAILABLE LEAVE"
            Height          =   255
            Left            =   150
            TabIndex        =   31
            Top             =   1140
            Width           =   2325
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "USE PPH 21"
            Height          =   225
            Left            =   1350
            TabIndex        =   30
            Top             =   1440
            Width           =   1125
         End
      End
   End
   Begin prj_tpc.vbButton cmdExit 
      Height          =   705
      Left            =   10560
      TabIndex        =   28
      Top             =   7230
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
      MICON           =   "frm_mst_compensation.frx":13331
      PICN            =   "frm_mst_compensation.frx":1334D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER COMPENSATION"
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
      TabIndex        =   29
      Top             =   150
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_compensation.frx":143DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frm_mst_compensation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsResign As New ADODB.Recordset
Dim rsPension As New ADODB.Recordset
Dim rsPHK As New ADODB.Recordset

Dim rspph As New ADODB.Recordset
Dim rsAllowRange As New ADODB.Recordset

Dim vIdNumber As Integer
            
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer

Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False
End Function

Private Sub check_invalid()

End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
End Function

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
'On Error GoTo Err

Dim vFlagOTMethod As Integer
Dim vFlagCalcStart As Integer
Dim v_flag_presence As Integer
Dim v_flag_meal As Integer
Dim v_flag_transport As Integer

    vSetData = 1
    
    If SSTab1.Tab = 0 Then
        If rsResign.RecordCount > 0 Then
            With rsResign
                TDBCombo_pesangon(SSTab1.Tab).Text = IIf(IsNull(.Fields("pesangon_code").Value), "", .Fields("pesangon_code").Value)
                txt_pesangon_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("allow_name").Value), "", .Fields("allow_name").Value)
                
                chk_resign_leave.Value = IIf(IsNull(.Fields("flag_leave").Value), 0, .Fields("flag_leave").Value)
                chk_resign_pph.Value = IIf(IsNull(.Fields("flag_pph21").Value), 0, .Fields("flag_pph21").Value)
                
                TDBCombo_pph(SSTab1.Tab).Text = IIf(IsNull(.Fields("pph21_code").Value), "", .Fields("pph21_code").Value)
                txt_pph21_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("pph21_name").Value), "", .Fields("pph21_name").Value)
            End With
        End If
    ElseIf SSTab1.Tab = 1 Then
        vSetData = 0
        If rsPension.RecordCount > 0 Then
            With rsPension
                chk_pesangon(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_pesangon").Value), 0, .Fields("flag_pesangon").Value)
                TDBCombo_pesangon(SSTab1.Tab).Text = IIf(IsNull(.Fields("pesangon_code").Value), "", .Fields("pesangon_code").Value)
                txt_pesangon_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("allow_name").Value), "", .Fields("allow_name").Value)
                
                chk_masa_kerja(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_masa_kerja").Value), 0, .Fields("flag_masa_kerja").Value)
                TDBCombo_masa_kerja(SSTab1.Tab).Text = IIf(IsNull(.Fields("masa_kerja_code").Value), "", .Fields("masa_kerja_code").Value)
                txt_masa_kerja_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("allow_name").Value), "", .Fields("allow_name").Value)
                
                chk_leave(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_leave").Value), 0, .Fields("flag_leave").Value)
                chk_ganti(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_hak").Value), 0, .Fields("flag_hak").Value)
                
                txt_ganti(SSTab1.Tab).Text = IIf(IsNull(.Fields("prosentase_value").Value), 0, .Fields("prosentase_value").Value)
                txt_kali_pesangon(SSTab1.Tab).Text = IIf(IsNull(.Fields("pengali_pesangon").Value), 0, .Fields("pengali_pesangon").Value)
                txt_kali_masa_kerja(SSTab1.Tab).Text = IIf(IsNull(.Fields("pengali_masa_kerja").Value), 0, .Fields("pengali_masa_kerja").Value)
                
                chk_pph21(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_pph21").Value), "", .Fields("flag_pph21").Value)
                TDBCombo_pph(SSTab1.Tab).Text = IIf(IsNull(.Fields("pph21_code").Value), "", .Fields("pph21_code").Value)
                txt_pph21_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("pph21_name").Value), "", .Fields("pph21_name").Value)
            End With
        End If
    ElseIf SSTab1.Tab = 2 Then
        vSetData = 0
        If rsPHK.RecordCount > 0 Then
            With rsPHK
                chk_pesangon(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_pesangon").Value), 0, .Fields("flag_pesangon").Value)
                TDBCombo_pesangon(SSTab1.Tab).Text = IIf(IsNull(.Fields("pesangon_code").Value), "", .Fields("pesangon_code").Value)
                txt_pesangon_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("allow_name").Value), "", .Fields("allow_name").Value)
                
                chk_masa_kerja(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_masa_kerja").Value), 0, .Fields("flag_masa_kerja").Value)
                TDBCombo_masa_kerja(SSTab1.Tab).Text = IIf(IsNull(.Fields("masa_kerja_code").Value), "", .Fields("masa_kerja_code").Value)
                txt_masa_kerja_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("allow_name").Value), "", .Fields("allow_name").Value)
                
                chk_leave(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_leave").Value), 0, .Fields("flag_leave").Value)
                chk_ganti(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_hak").Value), 0, .Fields("flag_hak").Value)
                
                txt_ganti(SSTab1.Tab).Text = IIf(IsNull(.Fields("prosentase_value").Value), 0, .Fields("prosentase_value").Value)
                txt_kali_pesangon(SSTab1.Tab).Text = IIf(IsNull(.Fields("pengali_pesangon").Value), 0, .Fields("pengali_pesangon").Value)
                txt_kali_masa_kerja(SSTab1.Tab).Text = IIf(IsNull(.Fields("pengali_masa_kerja").Value), 0, .Fields("pengali_masa_kerja").Value)
                
                chk_pph21(SSTab1.Tab).Value = IIf(IsNull(.Fields("flag_pph21").Value), 0, .Fields("flag_pph21").Value)
                TDBCombo_pph(SSTab1.Tab).Text = IIf(IsNull(.Fields("pph21_code").Value), "", .Fields("pph21_code").Value)
                txt_pph21_name(SSTab1.Tab).Text = IIf(IsNull(.Fields("pph21_name").Value), "", .Fields("pph21_name").Value)
            End With
        End If
    End If
    
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub chk_ganti_Click(Index As Integer)
    If chk_ganti(SSTab1.Tab).Value = 0 Then
        txt_ganti(SSTab1.Tab).Visible = False
        lblGanti(SSTab1.Tab).Visible = False
        
        txt_ganti(SSTab1.Tab).Text = 0
    Else
        txt_ganti(SSTab1.Tab).Visible = True
        lblGanti(SSTab1.Tab).Visible = True
    End If
End Sub

Private Sub chk_masa_kerja_Click(Index As Integer)
    If chk_masa_kerja(SSTab1.Tab).Value = 0 Then
        TDBCombo_masa_kerja(SSTab1.Tab).Visible = False
        txt_masa_kerja_name(SSTab1.Tab).Visible = False
        
        TDBCombo_masa_kerja(SSTab1.Tab).Text = ""
        txt_masa_kerja_name(SSTab1.Tab).Text = ""
    Else
        TDBCombo_masa_kerja(SSTab1.Tab).Visible = True
        txt_masa_kerja_name(SSTab1.Tab).Visible = True
    End If
End Sub

Private Sub chk_pesangon_Click(Index As Integer)
    If chk_pesangon(SSTab1.Tab).Value = 0 Then
        TDBCombo_pesangon(SSTab1.Tab).Visible = False
        txt_pesangon_name(SSTab1.Tab).Visible = False
        
        TDBCombo_pesangon(SSTab1.Tab).Text = ""
        txt_pesangon_name(SSTab1.Tab).Text = ""
    Else
        TDBCombo_pesangon(SSTab1.Tab).Visible = True
        txt_pesangon_name(SSTab1.Tab).Visible = True
    End If
End Sub

Private Sub chk_pph21_Click(Index As Integer)
    If chk_pph21(SSTab1.Tab).Value = 0 Then
        TDBCombo_pph(SSTab1.Tab).Visible = False
        txt_pph21_name(SSTab1.Tab).Visible = False
        
        TDBCombo_pph(SSTab1.Tab).Text = ""
        txt_pph21_name(SSTab1.Tab).Text = ""
    Else
        TDBCombo_pph(SSTab1.Tab).Visible = True
        txt_pph21_name(SSTab1.Tab).Visible = True
    End If
End Sub

Private Sub chk_resign_pph_Click()
    If chk_resign_pph.Value = 0 Then
        TDBCombo_pph(SSTab1.Tab).Visible = False
        txt_pph21_name(SSTab1.Tab).Visible = False
        
        TDBCombo_pph(SSTab1.Tab).Text = ""
        txt_pph21_name(SSTab1.Tab).Text = ""
    Else
        TDBCombo_pph(SSTab1.Tab).Visible = True
        txt_pph21_name(SSTab1.Tab).Visible = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub new_data()
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()
'On Error GoTo Err

Dim vNo As Integer

    CnG.BeginTrans
    
    If SSTab1.Tab = 0 Then
        SQL = "DELETE FROM m_comp_resign"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_comp_resign(pesangon_code,flag_leave,flag_pph21," & _
                "pph21_code,entry_date,entry_user) " & _
              "VALUES " & _
                "('" & TDBCombo_pesangon(SSTab1.Tab).Text & "','" & chk_resign_leave.Value & "'," & _
                "'" & chk_resign_pph.Value & "','" & TDBCombo_pph(SSTab1.Tab).Text & "'," & _
                "now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "DELETE FROM m_comp_pension"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_comp_pension(flag_pesangon,pesangon_code,flag_masa_kerja," & _
                "masa_kerja_code,flag_leave,flag_hak,prosentase_value,pengali_pesangon," & _
                "pengali_masa_kerja,flag_pph21,pph21_code,entry_date,entry_user) " & _
              "VALUES " & _
                "('" & chk_pesangon(SSTab1.Tab).Value & "','" & TDBCombo_pesangon(SSTab1.Tab).Text & "'," & _
                "'" & chk_masa_kerja(SSTab1.Tab).Value & "','" & TDBCombo_masa_kerja(SSTab1.Tab).Text & "'," & _
                "'" & chk_leave(SSTab1.Tab).Value & "','" & chk_ganti(SSTab1.Tab).Value & "'," & _
                "'" & txt_ganti(SSTab1.Tab).Text & "','" & txt_kali_pesangon(SSTab1.Tab).Text & "'," & _
                "'" & txt_kali_masa_kerja(SSTab1.Tab).Text & "','" & chk_pph21(SSTab1.Tab) & "'," & _
                "'" & TDBCombo_pph(SSTab1.Tab).Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "DELETE FROM m_comp_phk"
        CnG.Execute SQL
        
        SQL = "INSERT INTO m_comp_phk(flag_pesangon,pesangon_code,flag_masa_kerja," & _
                "masa_kerja_code,flag_leave,flag_hak,prosentase_value,pengali_pesangon," & _
                "pengali_masa_kerja,flag_pph21,pph21_code,entry_date,entry_user) " & _
              "VALUES " & _
                "('" & chk_pesangon(SSTab1.Tab).Value & "','" & TDBCombo_pesangon(SSTab1.Tab).Text & "'," & _
                "'" & chk_masa_kerja(SSTab1.Tab).Value & "','" & TDBCombo_masa_kerja(SSTab1.Tab).Text & "'," & _
                "'" & chk_leave(SSTab1.Tab).Value & "','" & chk_ganti(SSTab1.Tab).Value & "'," & _
                "'" & txt_ganti(SSTab1.Tab).Text & "','" & txt_kali_pesangon(SSTab1.Tab).Text & "'," & _
                "'" & txt_kali_masa_kerja(SSTab1.Tab).Text & "','" & chk_pph21(SSTab1.Tab) & "'," & _
                "'" & TDBCombo_pph(SSTab1.Tab).Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

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

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal C As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
    cmdSave(SSTab1.Tab).Enabled = a And blnUser_Add
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            Ctr.Value = Now
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
    
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
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

Private Sub CmdSave_Click(Index As Integer)
    Call insert_new_data
    
    MsgBox "Save Succesfully...", vbInformation, headerMSG
End Sub

Private Sub Form_Load()
    Call load_data_user_access(Me)
    
    SSTab1.Tab = 0
    
    Call load_data_pesangon(SSTab1.Tab)
    Call load_data
End Sub

Private Sub clear_filter()

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
    Set frm_mst_compensation = Nothing
End Sub

Private Sub filter_change()
On Error GoTo Err

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsResign.State Then rsResign.Close
        SQL = "select a.*, b.pph21_name, c.allow_name " & _
                "from m_comp_resign a left join m_pph21 b on a.pph21_code = b.pph21_code " & _
                    "left join m_allow_gen c on a.pesangon_code = c.allow_code"
        rsResign.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 1 Then
        If rsPension.State Then rsPension.Close
        SQL = "select a.*, b.pph21_name, c.allow_name, d.allow_name masa_kerja_name " & _
                "from m_comp_pension a left join m_pph21 b on a.pph21_code = b.pph21_code " & _
                    "left join m_allow_gen c on a.pesangon_code = c.allow_code " & _
                    "left join m_allow_gen d on a.masa_kerja_code = d.allow_code"
        rsPension.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    ElseIf SSTab1.Tab = 2 Then
        If rsPHK.State Then rsPHK.Close
        SQL = "select a.*, b.pph21_name, c.allow_name, d.allow_name masa_kerja_name " & _
                "from m_comp_phk a left join m_pph21 b on a.pph21_code = b.pph21_code " & _
                    "left join m_allow_gen c on a.pesangon_code = c.allow_code " & _
                    "left join m_allow_gen d on a.masa_kerja_code = d.allow_code"
        rsPHK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        Call set_edit_data
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Call load_data_user_access(Me)
    
    Call load_data_pesangon(SSTab1.Tab)
    Call load_data_pph21(SSTab1.Tab)
    Call load_data
End Sub

Private Sub TDBCombo_pesangon_ItemChange(Index As Integer)
    If TDBCombo_pesangon(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_pesangon(SSTab1.Tab).Text = TDBCombo_pesangon(SSTab1.Tab).Columns("pesangon_code").Value
        txt_pesangon_name(SSTab1.Tab).Text = TDBCombo_pesangon(SSTab1.Tab).Columns("allow_name").Value
    End If
End Sub

Private Sub TDBCombo_masa_kerja_ItemChange(Index As Integer)
    If TDBCombo_masa_kerja(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_masa_kerja(SSTab1.Tab).Text = TDBCombo_masa_kerja(SSTab1.Tab).Columns("masa_kerja_code").Value
        txt_masa_kerja_name(SSTab1.Tab).Text = TDBCombo_masa_kerja(SSTab1.Tab).Columns("masa_kerja_name").Value
    End If
End Sub

Private Sub TDBCombo_pph_ItemChange(Index As Integer)
    If TDBCombo_pph(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_pph(SSTab1.Tab).Text = TDBCombo_pph(SSTab1.Tab).Columns("pph21_code").Value
        txt_pph21_name(SSTab1.Tab).Text = TDBCombo_pph(SSTab1.Tab).Columns("pph21_name").Value
    End If
End Sub

Private Sub load_data_pesangon(Index As Integer)
    If rsAllowRange.State Then rsAllowRange.Close
    SQL = "select allow_code pesangon_code, allow_name, allow_code masa_kerja_code, allow_name masa_kerja_name " & _
            "from m_allow_gen order by allow_code"
    rsAllowRange.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_pesangon(SSTab1.Tab).RowSource = rsAllowRange
    If SSTab1.Tab <> 0 Then
        TDBCombo_masa_kerja(SSTab1.Tab).RowSource = rsAllowRange
    End If
End Sub

Private Sub load_data_pph21(Index As Integer)
    If rspph.State Then rspph.Close
    SQL = "select * from m_pph21 order by pph21_code"
    rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_pph(SSTab1.Tab).RowSource = rspph
End Sub
