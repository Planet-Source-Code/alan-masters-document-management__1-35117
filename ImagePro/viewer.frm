VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "IMGEDIT.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form viewer 
   Caption         =   "Image Pro V1.10 Viewer"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtComments 
      Height          =   2055
      Left            =   8400
      TabIndex        =   39
      Top             =   8160
      Width           =   2535
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   8400
      TabIndex        =   38
      Top             =   7560
      Width           =   2535
   End
   Begin VB.TextBox txtRef 
      Height          =   285
      Left            =   8400
      TabIndex        =   37
      Top             =   6960
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "viewer.frx":0000
      Height          =   1575
      Left            =   3960
      OleObjectBlob   =   "viewer.frx":0014
      TabIndex        =   1
      Top             =   240
      Width           =   6975
   End
   Begin VB.Frame frmRotate 
      Caption         =   "Rotate"
      Height          =   1575
      Left            =   8400
      TabIndex        =   29
      Top             =   5040
      Width           =   2175
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Flip Image"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Rotate Left"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Rotate Right"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frmZoom 
      Caption         =   "Zoom"
      Height          =   3015
      Left            =   8400
      TabIndex        =   14
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To Selection"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   28
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To 200%"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To 150%"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To 100%"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To 75%"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To 50%"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom To 25%"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   8295
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   8175
      _Version        =   131074
      _ExtentX        =   14420
      _ExtentY        =   14631
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      UndoBufferSize  =   64297216
      ForceFileLinking1x=   -1  'True
      lReserved1      =   22777984
      lReserved2      =   22777984
   End
   Begin VB.TextBox txtTable 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   3000
         Picture         =   "viewer.frx":09E7
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox schString 
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Text            =   "Dave"
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   4
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "In this Field:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Search for this string:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Title"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Ref No"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "Id No"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Text            =   "Image Path & Filename"
      Top             =   6480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox tables 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Text            =   "Images"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label10 
      Caption         =   "Comments:"
      Height          =   255
      Left            =   8400
      TabIndex        =   47
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Image Title:"
      Height          =   255
      Left            =   8400
      TabIndex        =   46
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Reference No:"
      Height          =   255
      Left            =   8400
      TabIndex        =   45
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Records Found:"
      Height          =   255
      Left            =   3960
      TabIndex        =   44
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblCountRecords 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5400
      TabIndex        =   36
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim txtPath As String
Dim RSTemp
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String
Dim dbpath3 As String
Private Sub cmdSearch_Click()

KeySection = "DB"
KeyKey = "DBPATH"
Call DB_Path
dbpath3 = KeyValue
Data1.DatabaseName = dbpath3
Dim MyQuery As String
MyQuery = txtTable.Text & " " & " Like " & " '*" & schString.Text & "*'"
Data1.RecordSource = "Select id_no, ref_no, title, comments, location from images where" & " " & MyQuery
On Error Resume Next
Data1.Refresh
Data1.Recordset.MoveLast: Data1.Recordset.MoveFirst
lblCountRecords.Caption = Data1.Recordset.RecordCount & " records that match your criteria"
'lblCountRecords.Caption = Data1.Recordset.RecordCount & " records"
Exit Sub
OOPS:
MsgBox "No Records Found"
End Sub
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, _
                                  X As Single, Y As Single)
        On Error GoTo Opps5:
         igColumn = DBGrid1.ColContaining(X)   'Set the Cell Column number
         igRow = DBGrid1.RowContaining(Y)      'Set the Cell Row number
   
         vargBookmark = DBGrid1.RowBookmark(igRow) 'Set Bookmark Value
         'Show the contents of the cell in a textbox
         'Text1.Text = DBGrid1.Columns(igColumn).CellValue(vargBookmark)
   Text1.Text = DBGrid1.Columns(4).CellValue(vargBookmark)
   txtRef.Text = DBGrid1.Columns(1).CellValue(vargBookmark)
   txtTitle.Text = DBGrid1.Columns(2).CellValue(vargBookmark)
   txtComments.Text = DBGrid1.Columns(3).CellValue(vargBookmark)
   ImgEdit1.Image = Text1.Text
   ImgEdit1.FitTo 1
   ImgEdit1.Display
   frmZoom.Enabled = True
   frmRotate.Enabled = True
   
   Exit Sub
Opps5:
      End Sub

Private Sub Form_Load()
 Dim igColumn As Integer
      Dim igRow As Integer
      Dim vargBookmark As Variant

Dim X
Data1.DatabaseName = App.Path & "\db\imagepro.mdb"
Data1.RecordSource = "Select id_no, ref_no, title, comments, location from images"
On Error Resume Next
Data1.Refresh
Data1.Recordset.MoveLast: Data1.Recordset.MoveFirst
lblCountRecords.Caption = Data1.Recordset.RecordCount & " records that match your criteria"
frmZoom.Enabled = False
frmRotate.Enabled = False
Exit Sub
OOPS:
MsgBox "No Records Found"
'txtPath = "d:\imagepro\db\imagepro.mdb"
'Set DB = OpenDatabase(txtPath, , True, "Access")
'cbotitles.Clear
'Set RSTemp = DB.OpenRecordset(tables.Text)
'For X = 0 To RSTemp.Fields.Count - 1
'cbotitles.AddItem RSTemp.Fields(X).Name
'Next X
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
txtTable.Text = "ID No"
Case 1
txtTable.Text = "Ref No"
Case 2
txtTable.Text = "Title"
Case 3
txtTable.Text = "Comments"
End Select

End Sub

Private Sub Option2_Click(Index As Integer)
On Error GoTo Opps10:
Select Case Index
Case 0
ImgEdit1.Zoom = 25
ImgEdit1.Refresh
Case 1
ImgEdit1.Zoom = 50
ImgEdit1.Refresh
Case 2
ImgEdit1.Zoom = 75
ImgEdit1.Refresh
Case 3
ImgEdit1.Zoom = 100
ImgEdit1.Refresh
Case 4
ImgEdit1.Zoom = 150
ImgEdit1.Refresh
Case 5
ImgEdit1.Zoom = 200
ImgEdit1.Refresh
Case 6
ImgEdit1.ZoomToSelection
ImgEdit1.Refresh
End Select
Exit Sub
Opps10:

End Sub

Private Sub Option3_Click(Index As Integer)
Select Case Index
Case 0
ImgEdit1.RotateLeft
ImgEdit1.Refresh
Case 1
ImgEdit1.RotateRight
ImgEdit1.Refresh
Case 2
ImgEdit1.Flip
ImgEdit1.Refresh
End Select
End Sub
Public Sub DB_Path()
Dim lngResult As Long
Dim strFileName
Dim strResult As String * 90
strFileName = App.Path & "\DB.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If


End Sub
