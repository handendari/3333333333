VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form sisaStock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proses Tutup Kandang"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.vbButton vbButton1 
      Height          =   375
      Left            =   3690
      TabIndex        =   11
      Top             =   4020
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Hitung Sisa"
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
      MICON           =   "frm_sisaStock.frx":0000
      PICN            =   "frm_sisaStock.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1230
      TabIndex        =   10
      Top             =   3570
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   930
      TabIndex        =   1
      Top             =   810
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   95354883
      CurrentDate     =   40291
   End
   Begin Project1.vbButton vbButton3 
      Height          =   525
      Left            =   1500
      TabIndex        =   3
      Top             =   4920
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   926
      BTYPE           =   14
      TX              =   "Exit"
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
      MICON           =   "frm_sisaStock.frx":0A2E
      PICN            =   "frm_sisaStock.frx":0A4A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.vbButton vbButton2 
      Height          =   525
      Left            =   180
      TabIndex        =   4
      Top             =   4920
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   926
      BTYPE           =   14
      TX              =   "Save"
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
      MICON           =   "frm_sisaStock.frx":1ADC
      PICN            =   "frm_sisaStock.frx":1AF8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.EnterNum txtekor 
      Height          =   315
      Left            =   1590
      TabIndex        =   5
      Top             =   3180
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Decimals        =   0
   End
   Begin Project1.EnterNum txtkg 
      Height          =   315
      Left            =   4320
      TabIndex        =   6
      Top             =   3180
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Decimals        =   1
   End
   Begin Project1.EnterNum txtekorkemarin 
      Height          =   315
      Left            =   1590
      TabIndex        =   12
      Top             =   1440
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   0
   End
   Begin Project1.EnterNum txtkgkemarin 
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Top             =   1440
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   1
   End
   Begin Project1.EnterNum txtekorterima 
      Height          =   315
      Left            =   1590
      TabIndex        =   16
      Top             =   1770
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   0
   End
   Begin Project1.EnterNum txtkgterima 
      Height          =   315
      Left            =   4320
      TabIndex        =   17
      Top             =   1770
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   1
   End
   Begin Project1.EnterNum txtekormatkan 
      Height          =   315
      Left            =   1590
      TabIndex        =   20
      Top             =   2130
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   0
   End
   Begin Project1.EnterNum txtkgmatkan 
      Height          =   315
      Left            =   4320
      TabIndex        =   21
      Top             =   2130
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   1
   End
   Begin Project1.EnterNum txtekorjual 
      Height          =   315
      Left            =   1590
      TabIndex        =   24
      Top             =   2490
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   0
   End
   Begin Project1.EnterNum txtkgjual 
      Height          =   315
      Left            =   4320
      TabIndex        =   25
      Top             =   2490
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Decimals        =   1
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kgs Penjualan :"
      Height          =   225
      Left            =   3030
      TabIndex        =   27
      Top             =   2550
      Width           =   1275
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ekor Penjualan :"
      Height          =   225
      Left            =   180
      TabIndex        =   26
      Top             =   2550
      Width           =   1395
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kgs Mati Kandang :"
      Height          =   225
      Left            =   2760
      TabIndex        =   23
      Top             =   2190
      Width           =   1545
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ekor Mati Kandang :"
      Height          =   225
      Left            =   120
      TabIndex        =   22
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kgs Penerimaan :"
      Height          =   225
      Left            =   3030
      TabIndex        =   19
      Top             =   1830
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ekor Penerimaan :"
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kgs Kemarin :"
      Height          =   225
      Left            =   3030
      TabIndex        =   15
      Top             =   1500
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ekor Kemarin :"
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   1500
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keterangan :"
      Height          =   225
      Left            =   300
      TabIndex        =   9
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sisa Ekor Hari Ini :"
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sisa Kgs Hari Ini :"
      Height          =   225
      Left            =   3030
      TabIndex        =   7
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   -180
      Picture         =   "frm_sisaStock.frx":2B8A
      Stretch         =   -1  'True
      Top             =   4770
      Width           =   5820
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tanggal"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   870
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROSES TUTUP KANDANG"
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
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   -510
      X2              =   12600
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   -90
      Picture         =   "frm_sisaStock.frx":4EF5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5730
   End
End
Attribute VB_Name = "sisaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public edit As Boolean
Dim sqlman As New clsSqlManag

Private Sub DTPicker1_Change()
    isiform
End Sub

Private Sub DTPicker1_Validate(Cancel As Boolean)
    isiform
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set sisaStock = Nothing
End Sub

Private Sub vbButton1_Click()
Dim rsjual As ADODB.Recordset
Dim tanggal1 As String
Dim sisaekor As Double, sisakgs As Double
    
    If sqlman.isSisaStockExist(Format(DTPicker1.Value - 1, "yyyy-MM-dd")) = False Then
        MsgBox "Transaksi Sisa Stock Kemarin Belum di Proses...!!" & Chr(13) & _
            "Periksa Kembali Inputan Sisa Stock Kemarin...", vbExclamation, INFO_USAHA.nmUsaha
        Exit Sub
    End If

    tanggal1 = Format(DTPicker1.Value, "YYYY-MM-dd")
    SQL = "select ifnull(sum(b.jmlekor),0)totekorjual,ifnull(sum(b.jmlkg),0)totkgsjual," _
             & "ifnull((select sum(jmlekorTU) from penerimaan WHERE tglterima = '" & tanggal1 & "'),0) totekormasuk," _
             & "ifnull((select sum(jmlKgTU) from penerimaan WHERE tglterima = '" & tanggal1 & "'),0) totkgmasuk," _
             & "ifnull((select sisaekor from sisastock where tanggal = date_add('" & tanggal1 & "',interval - 1 day)),0) sisaekor," _
             & "ifnull((select sisakgs from sisastock where tanggal = date_add('" & tanggal1 & "',interval - 1 day)),0) sisakgs," _
             & "ifnull((select sum(ekormati) from ayammatikandang where tanggal = '" & tanggal1 & "'),0) totEkorMatiKandang," _
             & "ifnull((select sum(kgmati) from ayammatikandang where tanggal = '" & tanggal1 & "'),0) totKgsMatiKandang " _
        & "from penjualan a join penjualan_detail b on a.nojual = b.nojual " _
        & "WHERE a.tgljual = '" & tanggal1 & "'"
    Set rsjual = New ADODB.Recordset
    rsjual.Open SQL, adoCN, adOpenForwardOnly, adLockReadOnly
    With rsjual
        txtekorkemarin.Value = !sisaekor
        txtkgkemarin.Value = !sisakgs
        txtekorterima.Value = !totekormasuk
        txtkgterima.Value = !totkgmasuk
        txtekormatkan.Value = !totekormatikandang
        txtkgmatkan.Value = !totkgsmatikandang
        txtekorjual.Value = !totekorjual
        txtkgjual.Value = !totkgsjual
        sisaekor = (!totekormasuk + !sisaekor) - (!totekorjual + !totekormatikandang)
        sisakgs = (!totkgmasuk + !sisakgs) - (!totkgsjual + !totkgsmatikandang)
    End With
    txtekor.Value = sisaekor
    txtkg.Value = sisakgs
End Sub

Private Sub vbButton2_Click()
'    If txtekor.Value = 0 Then
'        MsgBox "Jumlah Ekor Ayam Tidak Boleh Kosong...!!" & Chr(13) & _
'            "Periksa Kembali Inputan Anda....", vbInformation, INFO_USAHA.nmUsaha
'        Exit Sub
'    End If
    If edit = False Then
        SQL = "Insert into sisaStock (Tanggal,sisaekor,sisakgs," _
                & "keterangan,inputdate,userinput) " _
            & "VALUES ('" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'," _
                & "" & txtekor.Value & "," & txtkg.Value & "," _
                & "'" & Text1.Text & "',now(),'" & INFO_USER.kdUser & "')"
        adoCN.Execute SQL
    Else
        SQL = "update sisastock set " _
            & "sisaekor = '" & txtekor.Value & "',sisakgs = '" & txtkg.Value & "'," _
            & "keterangan = '" & Text1.Text & "',useredit = '" & INFO_USER.kdUser & "'," _
            & "editdate = now() where tanggal = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
        adoCN.Execute SQL
        Unload Me
        Exit Sub
    End If
    
    edit = False
    txtekor.Value = 0
    txtkg.Value = 0
    Text1.Text = ""
End Sub

Private Sub vbButton3_Click()
    Unload Me
End Sub

Private Sub isiform()
    SQL = "select * from sisaStock where tanggal = '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "'"
    Set rscari = New ADODB.Recordset
    rscari.Open SQL, adoCN, adOpenForwardOnly, adLockReadOnly
    If rscari.RecordCount > 0 Then
        With rscari
            DTPicker1.Value = rscari!tanggal
            txtkg.Value = rscari!sisakgs
            txtekor.Value = rscari!sisaekor
            Text1.Text = rscari!keterangan
        End With
        edit = True
    Else
        txtkg.Value = 0
        txtekor.Value = 0
        Text1.Text = ""
        edit = False
    End If
End Sub
