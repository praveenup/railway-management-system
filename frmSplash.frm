VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4755
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   7200
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   10
         X1              =   0
         X2              =   7215
         Y1              =   4560
         Y2              =   4575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   10
         X1              =   0
         X2              =   7215
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "PLATFORM"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5190
         TabIndex        =   7
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "WINDOWS"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   5175
         TabIndex        =   6
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Image Image2 
         Height          =   1920
         Left            =   4800
         Picture         =   "frmSplash.frx":000C
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   600
         Picture         =   "frmSplash.frx":C04E
         Top             =   3720
         Width           =   5985
      End
      Begin VB.Image imgLogo 
         Height          =   2985
         Left            =   120
         Picture         =   "frmSplash.frx":1C260
         Stretch         =   -1  'True
         Top             =   195
         Width           =   3015
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H8000000E&
         Caption         =   "Copyrights To "
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H8000000E&
         Caption         =   "MR. PRAVEEN KUMAR"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H8000000E&
         Caption         =   "Warning:For Study Purpose only."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3300
         Width           =   6855
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Gill Sans Ultra Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "INDIAN RAILWAYS"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   3315
         TabIndex        =   5
         Top             =   465
         Width           =   3795
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmMain.Show
End Sub



Private Sub Frame1_Click()
    Unload Me
    frmMain.Show
End Sub
