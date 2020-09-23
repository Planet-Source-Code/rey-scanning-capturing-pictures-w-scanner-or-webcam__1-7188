VERSION 5.00
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmScan 
   Caption         =   "simple scan example"
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "Scanform1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin ScanLibCtl.ImgScan imgscan 
      Left            =   240
      Top             =   360
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
      StopScanBox     =   -1  'True
      FileType        =   2
      PageType        =   6
      CompressionType =   1
      CompressionInfo =   0
      ScanTo          =   2
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop - Scan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan - Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "C:\temp.bmp"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Save Picture to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error GoTo hell
Dim Bx, Hwn, Scnav As Boolean
Scnav = imgscan.ScannerAvailable
Select Case Scnav
    Case True
'  checks for file type being used
    If InStr(Text1.Text, ".bmp") Then
     imgscan.FileType = BMP_Bitmap
    ElseIf InStr(Text1.Text, ".awa") Then
     imgscan.FileType = AWD_MicrosoftFax
    ElseIf InStr(Text1.Text, ".tif") Then
     imgscan.FileType = TIFF
        End If
    imgscan.ShowScanPreferences
'  displays scan preferences
    imgscan.image = Text1.Text
'  sets file to write to
    imgscan.ShowSelectScanner
'  if you have a webcam you can chose
'  to capture from the scanner or webcam
'  or if you have multiple scanners
    imgscan.StartScan
        Exit Sub
'  end sub before error handler
    Case False
        MsgBox "Scanner is Busy!" + vbCrLf + "Try Again Later.", vbInformation + vbsytemmodal, "Scanner: Busy!"
         Exit Sub
'  if busy gives them this
        End Select
hell: MsgBox "UnExpected Error Occured.", vbCritical, "Error!": Exit Sub

End Sub

Private Sub Command2_Click()
imgscan.StopScan
End Sub

Private Sub Command3_Click()
On Error Resume Next
  Dim hwnd
  Call ShellExecute(hwnd, "Open", Text1.Text, "", App.Path, 1)
End Sub

Private Sub Form_Load()
MsgBox "Images Can Only Be Saved as " + _
vbCrLf + "[1] *.tif" + _
vbCrLf + "[2] *.awd" + _
vbCrLf + "[3] *.bmp", vbInformation, "Important!"

End Sub

