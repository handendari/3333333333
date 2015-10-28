VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_summary_formula 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER SALARY SUMMARY FORMULA"
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
   Icon            =   "frm_mst_salary_summary_formula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_entry 
      Height          =   3375
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   11295
      Begin VB.TextBox txt_formula_salary_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox cbo_operation 
         Height          =   315
         ItemData        =   "frm_mst_salary_summary_formula.frx":000C
         Left            =   4800
         List            =   "frm_mst_salary_summary_formula.frx":001C
         TabIndex        =   4
         Text            =   "OP"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txt_variable 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   4800
         TabIndex        =   25
         Top             =   2640
         Width           =   1695
         Begin VB.OptionButton opt_variable_minus 
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   7
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton opt_variable_plus 
            Caption         =   "+"
            Height          =   195
            Left            =   360
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   375
         Left            =   4800
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
         Begin VB.OptionButton opt_salary_plus 
            Caption         =   "+"
            Height          =   195
            Left            =   360
            TabIndex        =   2
            Top             =   120
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton opt_salary_minus 
            Caption         =   "-"
            Height          =   195
            Left            =   1080
            TabIndex        =   3
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.CommandButton cmd_browse_salary 
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
         Left            =   6480
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txt_formula_salary_code 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "OPERATION"
         Height          =   195
         Left            =   3240
         TabIndex        =   27
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "VARIABLE VALUE"
         Height          =   195
         Left            =   3240
         TabIndex        =   26
         Top             =   2400
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "VARIABLE SIGN"
         Height          =   195
         Left            =   3240
         TabIndex        =   23
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SALARY SIGN"
         Height          =   195
         Left            =   3240
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SALARY CODE"
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SALARY NAME"
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   18
      Top             =   5160
      Width           =   11295
      Begin VB.CommandButton cmd_select 
         Caption         =   "&Select"
         Height          =   645
         Left            =   7560
         Picture         =   "frm_mst_salary_summary_formula.frx":002C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   1800
         Picture         =   "frm_mst_salary_summary_formula.frx":05B6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5040
         Picture         =   "frm_mst_salary_summary_formula.frx":0B40
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   9600
         Picture         =   "frm_mst_salary_summary_formula.frx":10CA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   720
         Picture         =   "frm_mst_salary_summary_formula.frx":1654
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_salary_summary_formula.frx":1BDE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   3960
         Picture         =   "frm_mst_salary_summary_formula.frx":2168
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   2880
         Picture         =   "frm_mst_salary_summary_formula.frx":26F2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "SALARY CODE"
      Columns(0).DataField=   "formula_salary_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "SALARY NAME"
      Columns(1).DataField=   "formula_salary_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   16
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   "1"
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue=   "+"
      Columns(2).ValueItems(0).DisplayValue.vt=   8
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(1)._DefaultItem=   0
      Columns(2).ValueItems(1).Value=   "-1"
      Columns(2).ValueItems(1).Value.vt=   8
      Columns(2).ValueItems(1).DisplayValue=   "-"
      Columns(2).ValueItems(1).DisplayValue.vt=   8
      Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   2
      Columns(2).Caption=   "SALARY SIGN"
      Columns(2).DataField=   "formula_salary_sign"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "OPERATION"
      Columns(3).DataField=   "formula_operation"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "VARIABLE"
      Columns(4).DataField=   "formula_variable"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   16
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "1"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "+"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "-1"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   "-"
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   2
      Columns(5).Caption=   "VARIABLE SIGN"
      Columns(5).DataField=   "formula_variable_sign"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3201"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3122"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6562"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6482"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2355"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2037"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1958"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2540"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2461"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
      Caption         =   "LIST OF SALARY SUMMARY FORMULA"
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
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(58)  =   "Named:id=33:Normal"
      _StyleDefs(59)  =   ":id=33,.parent=0"
      _StyleDefs(60)  =   "Named:id=34:Heading"
      _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=34,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=35:Footing"
      _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=36:Selected"
      _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=37:Caption"
      _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(69)  =   "Named:id=38:HighlightRow"
      _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=39:EvenRow"
      _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=40:OddRow"
      _StyleDefs(74)  =   ":id=40,.parent=33"
      _StyleDefs(75)  =   "Named:id=41:RecordSelector"
      _StyleDefs(76)  =   ":id=41,.parent=34"
      _StyleDefs(77)  =   "Named:id=42:FilterBar"
      _StyleDefs(78)  =   ":id=42,.parent=33"
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
   Begin VB.Label lbl_header 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   240
      TabIndex        =   28
      Top             =   240
      Width           =   11295
   End
End
Attribute VB_Name = "frm_mst_salary_summary_formula"
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

str_sql = "select count(formula_salary_code) as rec_count from m_salary_summary_formula " _
    & "where formula_salary_code = '" & Trim(txt_formula_salary_code) & "' and salary_code='" _
    & frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_code").Value & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
txt_formula_salary_code = ""
If txt_formula_salary_code.Enabled = True Then txt_formula_salary_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not txt_formula_salary_code = Adodc1.Recordset.Fields("formula_salary_code").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_formula_salary_code) = "" Then
    MsgBox "Salary Code is empty!", vbOKOnly + vbInformation, headerMSG
    txt_formula_salary_code.SetFocus
    check_validate_new = False
    Exit Function
End If

'If Trim(txt_formula_salary_name) = "" Then
'    MsgBox "Item Name is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_formula_salary_name.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
End Function

Private Sub load_data()
'Timer1.Enabled = True
End Sub

Private Sub cmd_browse_salary_Click()
'frm_mst_salary_summary.public_int_mode = 0
'frm_mst_salary_summary.Show 1
End Sub

Private Sub cmd_select_Click()
On Error GoTo err_capture

Dim i As Integer
Dim item

If Not TDBGrid1.ApproxCount > 0 Then
    Exit Sub
End If


If public_int_mode = 0 Then
    With frm_mst_salary_summary.Adodc2.Recordset
        .AddNew
        .Fields("salary_code").Value = frm_mst_salary_summary.Adodc1.Recordset("salary_code").Value
        .Fields("sequence_number").Value = get_inc_number("d_salary_slip", "sequence_number", "where salary_code='" _
                            & frm_mst_salary_summary.Adodc1.Recordset("salary_code").Value & "'")
        '---------------------------------
        .Fields("salary_name").Value = frm_mst_salary_summary.Adodc1.Recordset("salary_name").Value
        
        .Fields("flag_main_salary").Value = TDBGrid1.Columns("flag_main_salary").Value
        .Fields("formula_salary_code").Value = TDBGrid1.Columns("formula_salary_code").Value
        .Fields("salary_item_name").Value = TDBGrid1.Columns("salary_item_name").Value
        .Fields("flag_sign").Value = TDBGrid1.Columns("flag_sign").Value
        .Fields("flag_detail").Value = TDBGrid1.Columns("flag_detail").Value
        .Fields("flag_per_presence").Value = TDBGrid1.Columns("flag_per_presence").Value
        .Fields("default_value").Value = TDBGrid1.Columns("default_value").Value
        '.Fields("description").Value = TDBGrid1.Columns("description").Value
        .Update
    End With
    
End If

Call CmdExit_Click

Exit Sub
err_capture:
MsgBox "Error selecting data!", vbInformation, headerMSG
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & TDBGrid1.Columns("formula_salary_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_salary_summary_formula where formula_salary_code = '" _
    & Adodc1.Recordset.Fields("formula_salary_code").Value & "' and salary_code = '" _
    & Adodc1.Recordset.Fields("salary_code").Value & "'"
CnG.CommitTrans

Call load_data_item
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    'txt_salary_code = .Fields("salary_code").Value
    txt_formula_salary_code = .Fields("formula_salary_code").Value
    'txt_salary_name = .Fields("salary_name").Value
    
    txt_formula_salary_name = .Fields("formula_salary_name").Value
    If .Fields("formula_salary_sign").Value = 1 Then
        opt_salary_plus = True
    Else
        opt_salary_minus = True
    End If
    
    cbo_operation.Text = .Fields("formula_operation").Value
    txt_variable = .Fields("formula_variable").Value
    
    If .Fields("formula_variable_sign").Value = 1 Then
        opt_variable_plus = True
    Else
        opt_variable_minus = True
    End If
    
    'txt_flag_detail = .Fields("flag_detail").Value
End With
End Sub

Private Sub cmdEdit_Click()
int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
int_mode = 1
Call load_mode
End Sub

Private Sub CmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary_summary_formula where formula_salary_code = 'άφ'", CnG, adOpenKeyset, adLockOptimistic


CnG.BeginTrans

With rsBound
    .AddNew
    
    .Fields("salary_code").Value = frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_code").Value
    .Fields("formula_salary_code").Value = Trim(txt_formula_salary_code)
    '---------------------------------------------------------------------------
    .Fields("salary_name").Value = frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_name").Value
    .Fields("formula_salary_name").Value = get_salary_name(Trim(txt_formula_salary_code))
    
    .Fields("formula_salary_sign").Value = IIf(opt_salary_plus, 1, -1)
    .Fields("formula_operation").Value = Trim(cbo_operation.Text)
    .Fields("formula_variable").Value = Val(DropAllComma(txt_variable))
    .Fields("formula_variable_sign").Value = IIf(opt_variable_plus, 1, -1)
'    .Fields("flag_detail").Value = Trim(txt_flag_detail)
    
    .Update
End With

CnG.CommitTrans
End Sub

Private Function get_salary_name(ByVal str1 As String) As String
Dim rs1 As New ADODB.Recordset

rs1.Open "select * from m_salary_summary where salary_code = '" & str1 & "'", CnG, adOpenStatic, adLockReadOnly

If rs1.RecordCount > 0 Then
    get_salary_name = rs1.Fields("salary_name").Value
Else
    get_salary_name = ""
End If
End Function

Private Sub edit_old_data()
On Error GoTo err_capture

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary_summary_formula where formula_salary_code = '" _
& Adodc1.Recordset.Fields("formula_salary_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic


CnG.BeginTrans
With rsBound

    .Fields("salary_code").Value = frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_code").Value
    .Fields("formula_salary_code").Value = Trim(txt_formula_salary_code)
    '---------------------------------------------------------------------------
    .Fields("salary_name").Value = frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_name").Value
    .Fields("formula_salary_name").Value = get_salary_name(Trim(txt_formula_salary_code))
    
    .Fields("formula_salary_sign").Value = IIf(opt_salary_plus, 1, -1)
    .Fields("formula_operation").Value = Trim(cbo_operation.Text)
    .Fields("formula_variable").Value = Val(DropAllComma(txt_variable))
    .Fields("formula_variable_sign").Value = IIf(opt_variable_plus, 1, -1)
'    .Fields("flag_detail").Value = Trim(txt_flag_detail)
    
    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub CmdSave_Click()
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

Call load_data_item
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
'cmd_refresh.Enabled = g
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
cbo_operation.ListIndex = 0
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_formula_salary_name.Enabled = False
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_formula_salary_code.Enabled = True Then
        txt_formula_salary_code.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    txt_formula_salary_name.Enabled = False
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

Private Sub Form_Load()
Adodc1.ConnectionString = strConn

lbl_header.Caption = "COMPONENT : (" _
    & frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_code").Value _
    & ") " & frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_name").Value
Call load_data_item

'Call load_data_user_access(Me)
int_mode = 0
Call load_mode
End Sub


Private Sub TDBGrid1_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid1, Adodc1)
End Sub

Private Sub load_data_item()
Adodc1.RecordSource = "select * from m_salary_summary_formula " _
& "where salary_code='" & frm_mst_salary_summary.Adodc1.Recordset.Fields("salary_code").Value _
& "' order by formula_salary_code"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub


