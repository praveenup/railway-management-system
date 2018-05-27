VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H008080FF&
   Caption         =   "Home"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18975
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   9615
   ScaleWidth      =   18975
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":4B9F0C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":4BD2EE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":4C1F34
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "PANEL"
      Height          =   6375
      Left            =   4440
      TabIndex        =   0
      Top             =   2040
      Width           =   11175
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   720
         Picture         =   "frmMain.frx":4C6B7A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Height          =   1695
         Left            =   7680
         Picture         =   "frmMain.frx":4D7944
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Height          =   1695
         Left            =   4440
         Picture         =   "frmMain.frx":4E870E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CommandButton Command4 
         Height          =   1695
         Left            =   720
         Picture         =   "frmMain.frx":4F94D8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   4440
         Picture         =   "frmMain.frx":50F7AA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         Height          =   5775
         Left            =   0
         Picture         =   "frmMain.frx":524C18
         ScaleHeight     =   5715
         ScaleWidth      =   11115
         TabIndex        =   9
         Top             =   600
         Width           =   11175
      End
      Begin VB.Image Image5 
         Height          =   645
         Left            =   0
         Picture         =   "frmMain.frx":60471C
         Top             =   0
         Width           =   11190
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RAILWAY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   5400
      TabIndex        =   10
      Top             =   600
      Width           =   10935
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   4440
      Picture         =   "frmMain.frx":61BF9E
      Top             =   480
      Width           =   11415
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   10440
      Picture         =   "frmMain.frx":644B2C
      Top             =   0
      Width           =   14565
   End
   Begin VB.Image Image8 
      Height          =   1530
      Left            =   0
      Picture         =   "frmMain.frx":65DA1A
      Top             =   480
      Width           =   2910
   End
   Begin VB.Image Image3 
      Height          =   525
      Left            =   0
      Picture         =   "frmMain.frx":66C30C
      Top             =   0
      Width           =   14565
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6840
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer

Private Sub Command1_Click()
frmPnrUser.Show 1
End Sub

Private Sub Command2_Click()
frmTrainScheduleUser.Show 1
End Sub

Private Sub Command3_Click()
frmFareEnquiryUser.Show 1
End Sub

Private Sub Command4_Click()
frmSeatAvailablityUser.Show 1
End Sub

Private Sub Command5_Click()
frmTrainBetStnUser.Show
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
If flag = 0 Then
    Command6.Visible = True
    Command8.Visible = True
    Image8.Visible = True
    flag = 1
Else
    flag = 0
    Command6.Visible = False
    Command8.Visible = False
    Image8.Visible = False
End If
End Sub



Private Sub Command8_Click()
Command6.Visible = False
Command8.Visible = False
Image8.Visible = False
frmLogin.Show
End Sub



Private Sub Form_Click()
flag = 0
Command6.Visible = False
Command8.Visible = False
Image8.Visible = False
End Sub

Private Sub Form_activate()
Command6.Visible = False
Command8.Visible = False
Image8.Visible = False
End Sub

Private Sub Frame1_Click()
flag = 0
Command6.Visible = False
Command8.Visible = False
Image8.Visible = False
End Sub

