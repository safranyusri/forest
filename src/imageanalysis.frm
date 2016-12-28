VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19875
   LinkTopic       =   "Form1"
   Picture         =   "imageanalysis.frx":0000
   ScaleHeight     =   774.749
   ScaleMode       =   0  'User
   ScaleWidth      =   1325
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   4800
      TabIndex        =   15
      Top             =   5640
      Width           =   4815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   17280
      TabIndex        =   8
      Text            =   "Treshold (0-255)"
      Top             =   7800
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   495
      Left            =   18120
      TabIndex        =   7
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyze Image"
      Height          =   735
      Left            =   18120
      TabIndex        =   6
      Top             =   7680
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   5334
      Left            =   9840
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   120
      Width           =   9600
   End
   Begin VB.PictureBox Picture1 
      Height          =   5334
      Left            =   120
      Picture         =   "imageanalysis.frx":A4C1
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
   Begin VB.Label Label11 
      Caption         =   "Safran Yusri  G051160021"
      Height          =   255
      Left            =   13680
      TabIndex        =   18
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Forest Estimator"
      Height          =   255
      Left            =   9840
      TabIndex        =   17
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   9840
      Picture         =   "imageanalysis.frx":215F4
      Top             =   6480
      Width           =   5730
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   615
      Left            =   9840
      TabIndex        =   16
      Top             =   5640
      Width           =   5895
   End
   Begin VB.Label Label8 
      Caption         =   "%"
      Height          =   375
      Left            =   19200
      TabIndex        =   12
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Percent Cover"
      Height          =   375
      Left            =   16200
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Pixel Count:"
      Height          =   375
      Left            =   16200
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Threshold (0-255)"
      Height          =   615
      Left            =   16080
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Percent Cover"
      Height          =   375
      Left            =   17640
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Pix Width"
      Height          =   375
      Left            =   17160
      TabIndex        =   4
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Pix Height"
      Height          =   495
      Left            =   16080
      TabIndex        =   3
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Pixel Count"
      Height          =   375
      Left            =   17640
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim threshold As Single
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
 If Right(File1.Path, 1) <> "\" Then
 Label9.Caption = File1.Path & "\" & File1.FileName
 Else
 Label9.Caption = File1.Path & File1.FileName
 End If
  Picture1.Picture = LoadPicture(Label9.Caption)
End Sub

Public Sub Form_Activate()
Text1.Text = 127
End Sub

Private Sub Form_Load()
Drive1.Drive = App.Path
Dir1.Path = App.Path
End Sub

Public Sub Text1_Change()
th = Val(Text1.Text)
End Sub
Private Sub Command1_Click()
black = 0
th = Val(Text1.Text)
 For i = 0 To Picture1.Width
  For j = 0 To Picture1.Height
   Colour = Picture1.Point(i, j)
    R = Colour And &HFF
    G = Colour \ 256 And &HFF
    B = Colour \ 256 ^ 2 And &HFF
    grey = (R + G + B) / 3
    If grey < th Then
    Picture1.PSet (i, j), RGB(R, G, B)
    Picture2.PSet (i, j), RGB(0, 0, 0)
    black = black + 1
     
   Else
    Picture1.PSet (i, j), RGB(R, G, B)
    Picture2.PSet (i, j), RGB(255, 255, 255)
   End If
  Next j
 Next i
Label1 = black
Label2 = i
Label3 = j
 Label4 = (black / (i * j)) * 100
End Sub

Private Sub Command2_Click()
Picture2.Picture = LoadPicture("")
Label1 = ""
Label2 = ""
Label3 = ""
 Label4 = ""
End Sub
Private Sub Picture1_Click()
Dir1.Path = Drive1.Drive
End Sub
