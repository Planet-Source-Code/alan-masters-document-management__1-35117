VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "ImagePro V1.10"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu1 
      Caption         =   "&View"
      Begin VB.Menu mnu2 
         Caption         =   "&View Images"
      End
   End
   Begin VB.Menu mnu5 
      Caption         =   "&Add"
      Begin VB.Menu mnu6 
         Caption         =   "Add images to DB"
      End
   End
   Begin VB.Menu Mnu8 
      Caption         =   "&Scan"
      Begin VB.Menu Mnu9 
         Caption         =   "S&can Image"
      End
   End
   Begin VB.Menu mnuSet 
      Caption         =   "Settings"
      Begin VB.Menu mnuSet2 
         Caption         =   "Specify DataSource"
      End
   End
   Begin VB.Menu mnuAbt 
      Caption         =   "A&bout"
      Begin VB.Menu mnuAbt1 
         Caption         =   "About this App"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu2_Click()
Dim frm As New viewer
frm.Show
End Sub

Private Sub mnu6_Click()
Dim frm As New browser
frm.Show
End Sub

Private Sub mnu9_Click()
Dim frm As New frmScan
frm.Show
End Sub

Private Sub mnuAbt1_Click()
Dim frm As New frmAboutScan
frm.Show
End Sub

Private Sub mnuSet2_Click()
Dim frm As New frmDBPath
frm.Show
End Sub
