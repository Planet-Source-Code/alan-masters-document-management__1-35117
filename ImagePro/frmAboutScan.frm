VERSION 5.00
Begin VB.Form frmAboutScan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ImagePro V1.10"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frmAboutScan.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAboutScan.frx":0489
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   1
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Image Pro V1.10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Programming  && Development by Alan Masters"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAboutScan.frx":05AE
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   2
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAboutScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About ImagePro V1.10"
End Sub
