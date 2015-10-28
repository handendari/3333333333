VERSION 5.00
Begin VB.MDIForm mdi_fgp 
   BackColor       =   &H8000000C&
   Caption         =   "FGP DATA ENROLL MANAGER 1.0"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9120
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_setting 
         Caption         =   "&Setting"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_master 
      Caption         =   "&Master"
      Begin VB.Menu mnu_mst_user 
         Caption         =   "&User"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnu_utility 
      Caption         =   "&Utility"
      Begin VB.Menu mnu_exportr_data 
         Caption         =   "&Export Data"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "mdi_fgp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_exit_Click()
End
End Sub

Private Sub mnu_exportr_data_Click()
frm_export_user.Show
End Sub

Private Sub mnu_mst_user_Click()
frm_master_user.Show
End Sub

Private Sub mnu_setting_Click()
frm_setting.Show 1
End Sub
