VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "IMGEDIT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form browser 
   BackColor       =   &H80000005&
   Caption         =   "ImagePro Browser V1.10"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "browser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "browser.frx":0442
      TabIndex        =   25
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Load Image from location"
      Height          =   4575
      Left            =   3240
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   4080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   3120
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Full Path to File Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "File Name:"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Folder:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Drive Letter:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox new_image_path 
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optRight 
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optright2 
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton optLeft 
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdSelZoom 
      Caption         =   "Selection Zoom"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox cboZoom 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Text            =   "Zoom In / Out"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "Zoom"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Image"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox comments 
      DataField       =   "comments"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "browser.frx":0544
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox image_title 
      DataField       =   "title"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "image_title"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox image_path 
      DataField       =   "location"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "image_path"
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox file_name 
      DataField       =   "filename"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "file_name"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox guid1 
      DataField       =   "ref_no"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "GUID1"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.FileListBox file 
      Height          =   2625
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "images"
      Top             =   8520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin ImgeditLibCtl.ImgEdit img_box 
      Height          =   9855
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      _Version        =   131074
      _ExtentX        =   13150
      _ExtentY        =   17383
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      UndoBufferSize  =   64871680
      ForceFileLinking1x=   -1  'True
      lReserved1      =   22802568
      lReserved2      =   22802568
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      Caption         =   "Flip / Rotate Image"
      Height          =   255
      Left            =   1800
      TabIndex        =   32
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "Image File Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      Caption         =   "Image Title"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "Full Image Path"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "File Name"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "Reference No:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2880
      Picture         =   "browser.frx":054D
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   2280
      Picture         =   "browser.frx":064F
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   2640
      Picture         =   "browser.frx":0751
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   1920
      Picture         =   "browser.frx":0853
      Top             =   1680
      Width           =   240
   End
End
Attribute VB_Name = "browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB As Database
Public rstInfo As Recordset
Dim mpath

Private Sub cmdAdd_Click()
If image_title.Text <> "" Then
Set DB = OpenDatabase(App.Path & "\db\imagepro.mdb")
With DB

    Set rstInfo = .OpenRecordset("images")
        With rstInfo
        .AddNew
             !ref_no = guid1.Text
             !FileName = file_name.Text
             !location = new_image_path.Text
             !Title = image_title.Text
             !comments = comments.Text
        .Update
    End With
    End With
    Source = image_path.Text
    Destination = App.Path & "\img\" & file_name.Text
    FileCopy Source, Destination
    Kill Source
    file.Refresh
    guid1.Enabled = False
file_name.Enabled = False
image_path.Enabled = False
image_title.Enabled = False
comments.Enabled = False
img_box.Enabled = False
cmdAdd.Enabled = False
  Else
Prompt$ = "You must enter at least a Image title before adding this image to the database!"
reply = MsgBox(Prompt$, vbOKOnly, "Add Image Title")
image_title.SetFocus
End If

End Sub

Private Sub cmdclear_Click()
pic1.Visible = False
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Opps:
img_box.PrintImage
Exit Sub
Opps:
Prompt$ = "Opps. There was an error printing your document. You may have tryed to print without loading an image first."
reply = MsgBox(Prompt$, vbOKOnly, "Opps - Printing Error")
End Sub

Private Sub cmdSave_Click()
CommonDialog1.CancelError = True
On Error GoTo Errhandler:

CommonDialog1.Filter = "All Files (*.*)|*.*|Tiff Files" & _
  "(*.tif)|*.tif|JPEG (*.jpg)|*.jpg"
CommonDialog1.ShowSave
img_box.SavePage CommonDialog1.FileName
Errhandler:

End Sub

Private Sub cmdSelZoom_Click()
On Error GoTo NoSelMade
img_box.ZoomToSelection
img_box.Refresh
Exit Sub

NoSelMade:
Prompt$ = "You MUST make a selection before clicking this button."
reply = MsgBox(Prompt$, vbOKOnly, "No Selection Made")
End Sub

Private Sub cmdZoom_Click()
If cboZoom.Text = "Zoom In / Out" Then
img_box.Zoom = 100
Else
img_box.Zoom = cboZoom.Text
End If
img_box.Refresh
End Sub



Private Sub file_Click()
On Error GoTo picerror
img_box.Visible = True
mpath = file.Path + "\" + file.FileName
img_box.Image = mpath
img_box.FitTo 1
img_box.Display
img_box.Refresh
guid1.Text = GetUniqueId
file_name.Text = file.FileName
image_path.Text = mpath
image_title.Enabled = True
comments.Enabled = True
image_title.Text = ""
comments.Text = ""
cboZoom.Enabled = True
cmdZoom.Enabled = True
cmdSelZoom.Enabled = True
optRight.Enabled = True
optLeft.Enabled = True
optright2.Enabled = True
new_image_path = App.Path & "\img\" & file_name.Text
cmdPrint.Enabled = True
cmdSave.Enabled = True
Exit Sub
picerror:
Prompt$ = "You have selected an invalid image. Please select another."
reply = MsgBox(Prompt$, vbOKOnly, "Invalid File Format")


End Sub

Private Sub Form_Load()
file.Path = App.Path & "\tmpImg"
cmdAdd.Enabled = False
cboZoom.AddItem "25"
cboZoom.AddItem "50"
cboZoom.AddItem "75"
cboZoom.AddItem "100"
cboZoom.AddItem "200"
End Sub
Public Function GetUniqueId() As String
    GetUniqueId = Trim(Str(CDbl(Now) * 10000000000#))
End Function

Private Sub image_title_GotFocus()
cmdAdd.Enabled = True
End Sub

Private Sub Image2_Click()
img_box.Visible = False
Frame1.Visible = True
File1.Visible = True
Dir1.Visible = True
Drive1.Visible = True
Text1.Visible = True
End Sub

Private Sub optLeft_Click()
img_box.RotateLeft
img_box.FitTo 1
End Sub

Private Sub optRight_Click()
img_box.RotateRight
img_box.FitTo 1
End Sub

Private Sub optright2_Click()
img_box.Flip
img_box.FitTo 1
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error GoTo picerror
Text1.Text = File1.Path & "\" & File1.FileName
mpath = File1.Path + "\" + File1.FileName
guid1.Text = GetUniqueId
file_name.Text = File1.FileName
image_path.Text = mpath
image_title.Enabled = True
comments.Enabled = True
image_title.Text = ""
comments.Text = ""
cboZoom.Enabled = True
cmdZoom.Enabled = True
cmdSelZoom.Enabled = True
optRight.Enabled = True
optLeft.Enabled = True
optright2.Enabled = True
new_image_path.Text = App.Path & "\img\" & file_name.Text
image_path.Text = Text1.Text
img_box.Image = Text1.Text
img_box.Display
img_box.Refresh
Frame1.Visible = False
Dir1.Visible = False
Drive1.Visible = False
File1.Visible = False
Text1.Visible = False
img_box.Visible = True
cmdSave.Enabled = True
cmdPrint.Enabled = True
Exit Sub
picerror:
Prompt$ = "You have selected an invalid image. Please select another."
reply = MsgBox(Prompt$, vbOKOnly, "Invalid File Format")


End Sub


