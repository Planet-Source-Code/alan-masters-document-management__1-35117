VERSION 5.00
Object = "{009541A3-3B81-101C-92F3-040224009C02}#3.0#0"; "IMGADMIN.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "IMGEDIT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScan 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scan Image"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin ScanLibCtl.ImgScan ImgScan1 
      Left            =   0
      Top             =   7560
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
      DestImageControl=   "ImgEdit1"
      PageType        =   3
      CompressionType =   1
      CompressionInfo =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      _Version        =   131074
      _ExtentX        =   14843
      _ExtentY        =   13785
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UndoBufferSize  =   64321280
      OcrZoneVisibility=   -4044
      AnnotationOcrType=   127
      ForceFileLinking1x=   -1  'True
      lReserved1      =   22777984
      lReserved2      =   22777984
      Begin AdminLibCtl.ImgAdmin ImgAdmin1 
         Left            =   480
         Top             =   2280
         _Version        =   196608
         _ExtentX        =   873
         _ExtentY        =   1085
         _StockProps     =   0
         PrintStartPage  =   0
         PrintEndPage    =   0
      End
   End
   Begin MSComctlLib.Toolbar tlbNew 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1111
      ButtonWidth     =   2117
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "cmdSave"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "cmdPrint"
            Object.ToolTipText     =   "Print Current Image to Printer"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Key             =   "cmdSettings"
            Object.ToolTipText     =   "Set your Scanner Settings"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zoom"
            Key             =   "cmdZoom"
            Object.ToolTipText     =   "Zoom into Image 150%"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flip"
            Key             =   "cmdFlip"
            Object.ToolTipText     =   "Flip the current image 180 Degrees"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Scan to Display"
            Key             =   "cmdScanDisplay"
            Object.ToolTipText     =   "Scan a document to the dispaly"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Scan to File"
            Key             =   "cmdScanFile"
            Object.ToolTipText     =   "Scan a document to file"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "cmdSave"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   "cmdPrint"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   "cmdSettings"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0336
            Key             =   "cmdZoom"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0650
            Key             =   "cmdFlip"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14A2
            Key             =   "cmdScanDisplay"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18F4
            Key             =   "cmdScanFile"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "There are more image manipulation options from the drop down menus at the top of this screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8415
   End
   Begin VB.Menu mnu_file1 
      Caption         =   "&File"
      Begin VB.Menu mnuFile2 
         Caption         =   "Save Image"
      End
      Begin VB.Menu mnuFile3 
         Caption         =   "Print Image"
      End
      Begin VB.Menu mnuFile4 
         Caption         =   "Select Source"
      End
      Begin VB.Menu mnuFile5 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu1 
      Caption         =   "&Scan"
      Begin VB.Menu mnu2 
         Caption         =   "Scan to &Display"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Scan to File"
      End
      Begin VB.Menu mnu4 
         Caption         =   "Scan to Display && File"
      End
   End
   Begin VB.Menu mnu5 
      Caption         =   "S&ettings"
      Begin VB.Menu mnu6 
         Caption         =   "Scan Options"
      End
   End
   Begin VB.Menu mnu7 
      Caption         =   "&Zoom Options"
      Begin VB.Menu mnu8 
         Caption         =   "Zoom 50%"
      End
      Begin VB.Menu mnu9 
         Caption         =   "Zoom 100%"
      End
      Begin VB.Menu mnu10 
         Caption         =   "Zoom 150%"
      End
      Begin VB.Menu mnu11 
         Caption         =   "Zoom 200%"
      End
      Begin VB.Menu mnu12 
         Caption         =   "Zoom 500%"
      End
      Begin VB.Menu Mnu 
         Caption         =   "Zoom To Selection"
      End
   End
   Begin VB.Menu mnu13 
      Caption         =   "&Rotate Image"
      Begin VB.Menu mnu14 
         Caption         =   "Rotate Left"
      End
      Begin VB.Menu mnu15 
         Caption         =   "Rotate Right"
      End
      Begin VB.Menu mnu16 
         Caption         =   "Flip Image"
      End
   End
   Begin VB.Menu mnuAbout1 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout2 
         Caption         =   "About ImagePro Scan"
      End
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Mnu_Click()
On Error GoTo NoSelMade
ImgEdit1.ZoomToSelection
ImgEdit1.Refresh
Exit Sub

NoSelMade:
Prompt$ = "You MUST make a selection before clicking this button."
reply = MsgBox(Prompt$, vbOKOnly, "No Selection Made")
End Sub

Private Sub mnu10_Click()
On Error GoTo Opps:
ImgEdit1.Zoom = 150
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu11_Click()
On Error GoTo Opps:
ImgEdit1.Zoom = 200
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu12_Click()
On Error GoTo Opps:
ImgEdit1.Zoom = 500
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu14_Click()
On Error GoTo Opps:
ImgEdit1.RotateLeft
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu15_Click()
On Error GoTo Opps:
ImgEdit1.RotateRight
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu16_Click()
On Error GoTo Opps:
ImgEdit1.Flip
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu3_Click()
 On Error GoTo ScanError
    ImgScan1.ScanTo = FileOnly      'Scan to a file
    ImgScan1.Image = App.Path & "\tmpImg\" & GetUniqueId & ".tif"  'Destination file
    ImgScan1.StartScan
    Exit Sub
    
ScanError:
    MsgBox Err.Description, , "Scan Error"
    Exit Sub

End Sub

Private Sub mnu4_Click()
ImgScan1.ScanTo = DisplayAndFile
    'Set the image property to a file name.
    ImgScan1.Image = App.Path & "\tmpImg\" & GetUniqueId & ".tif"
    'Multipage must be true in order to create files with
    'more than one page.
    ImgScan1.MultiPage = False
    'Do not show the scanner's TWAIN UI.
    ImgScan1.ShowSetupBeforeScan = False
    'Scan without using dialog box.
    ImgScan1.ShowSetupBeforeScan = True
    ImgScan1.StartScan
End Sub
Private Sub mnu2_Click()
'scanner available?
ImgScan1.ScannerAvailable
' open scanner port
ImgScan1.OpenScanner
' start scanning
ImgScan1.StartScan
End Sub

Private Sub mnu6_Click()
ImgScan1.ShowScanPreferences
End Sub

Private Sub mnu8_Click()
On Error GoTo Opps:
ImgEdit1.Zoom = 50
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnu9_Click()
On Error GoTo Opps:
ImgEdit1.Zoom = 100
ImgEdit1.Refresh
Exit Sub
Opps:
End Sub

Private Sub mnuAbout2_Click()
Dim frm As New frmAboutScan
frm.Show
End Sub

Private Sub mnuFile2_Click()
On Error GoTo Errhandler:
ImgAdmin1.Filter = "TIFF Files (*.tif)|*.tif|"
ImgAdmin1.ShowFileDialog SaveDlg
ImgEdit1.SaveAs ImgAdmin1.Image + ".tif"
Exit Sub
Errhandler:
Prompt$ = "Opps. Save File Error. Please try again."
reply = MsgBox(Prompt$, vbOKOnly, "Opps - Save Error")
End Sub

Private Sub mnuFile3_Click()
On Error GoTo Opps:
ImgEdit1.PrintImage
Exit Sub
Opps:
Prompt$ = "Opps. There was an error printing your document. You may have tryed to print without loading an image first."
reply = MsgBox(Prompt$, vbOKOnly, "Opps - Printing Error")
End Sub

Private Sub mnuFile4_Click()
ImgScan1.ShowSelectScanner
End Sub

Private Sub mnuFile5_Click()
Unload Me
End Sub

Private Sub tlbNew_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim msgPress As Integer
Select Case Button.Key
Case Is = "cmdSave":
On Error GoTo Errhandler:
ImgAdmin1.Filter = "TIFF Files (*.tif)|*.tif|"
ImgAdmin1.ShowFileDialog SaveDlg
ImgEdit1.SaveAs ImgAdmin1.Image + ".tif"
Exit Sub
Errhandler:
Prompt$ = "Opps. Save File Error. Please try again."
reply = MsgBox(Prompt$, vbOKOnly, "Opps - Save Error")
Case Is = "cmdPrint":
On Error GoTo Opps1:
ImgEdit1.PrintImage
Exit Sub
Opps1:
Prompt$ = "Opps. There was an error printing your document. You may have tryed to print without loading an image first."
reply = MsgBox(Prompt$, vbOKOnly, "Opps - Printing Error")
Case Is = "cmdSettings":
ImgScan1.ShowScanPreferences
Case Is = "cmdZoom":
On Error GoTo Opps2:
ImgEdit1.Zoom = 150
ImgEdit1.Refresh
Exit Sub
Opps2:
Case Is = "cmdFlip":
On Error GoTo Opps:
ImgEdit1.Flip
ImgEdit1.Refresh
Exit Sub
Opps:
Case Is = "cmdScanDisplay":
'scanner available?
ImgScan1.ScannerAvailable
' open scanner port
ImgScan1.OpenScanner
' start scanning
ImgScan1.StartScan
Case Is = "cmdScanFile":
ImgScan1.ScanTo = DisplayAndFile
    'Set the image property to a file name.
    ImgScan1.Image = App.Path & "\tmpImg\" & GetUniqueId & ".tif"
    'Multipage must be true in order to create files with
    'more than one page.
    ImgScan1.MultiPage = False
    'Do not show the scanner's TWAIN UI.
    ImgScan1.ShowSetupBeforeScan = False
    'Scan without using dialog box.
    ImgScan1.ShowSetupBeforeScan = True
    ImgScan1.StartScan
End Select
End Sub
