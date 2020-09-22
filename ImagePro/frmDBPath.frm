VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDBPath 
   BackColor       =   &H80000005&
   Caption         =   "Browse For Database"
   ClientHeight    =   1575
   ClientLeft      =   3090
   ClientTop       =   4875
   ClientWidth     =   8100
   Icon            =   "frmDBPath.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   78.75
   ScaleMode       =   0  'User
   ScaleWidth      =   140.625
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      ToolTipText     =   "Click here to select a database."
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Type in a database name with full path. Then click the 'Refresh' button."
      Top             =   720
      Width           =   5520
   End
   Begin MSComDlg.CommonDialog cdlgPath 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Database Search"
      FileName        =   "*.MDB"
      Filter          =   "*.MDB"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Please select the location of the Database you wish to use for this program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000005&
      Caption         =   "Database Path:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1140
   End
End
Attribute VB_Name = "frmDBPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String

Private Sub cmdBrowse_Click()
cdlgPath.DialogTitle = "Browse for database file..." 'set the dialog title
cdlgPath.Filter = "Database files|*.mdb"
cdlgPath.CancelError = True
On Error GoTo errorhandler
cdlgPath.ShowOpen
  txtPath.Text = cdlgPath.FileName
Exit Sub

errorhandler:
  Exit Sub
  
End Sub

Private Sub saveini()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\DB.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub


' Save your ini parameters
Private Sub cmdSave_Click()

'Save DB path
KeySection = "DB"
KeyKey = "DBPath"
KeyValue = txtPath.Text
saveini


End Sub

