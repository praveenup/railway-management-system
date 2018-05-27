VERSION 5.00
Begin VB.Form frmCancellation 
   BackColor       =   &H8000000E&
   Caption         =   "Ticket Cancellation"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   480
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Ticket Details"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   6495
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   16455
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Passenger Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   56
         Top             =   2040
         Width           =   15975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   360
         Width           =   15975
      End
      Begin VB.Label lblcoach 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   4
         Left            =   9600
         TabIndex        =   54
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblcoach 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   9600
         TabIndex        =   53
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblcoach 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   9600
         TabIndex        =   52
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblcoach 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   9600
         TabIndex        =   51
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblcoach 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   9600
         TabIndex        =   50
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   4
         Left            =   8400
         TabIndex        =   49
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   8400
         TabIndex        =   48
         Top             =   4560
         Width           =   375
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   8400
         TabIndex        =   47
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   8400
         TabIndex        =   46
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   45
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   4
         Left            =   6960
         TabIndex        =   44
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   43
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   6960
         TabIndex        =   42
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   6960
         TabIndex        =   41
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   40
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   4
         Left            =   3360
         TabIndex        =   39
         Top             =   5040
         Width           =   3255
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   38
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   37
         Top             =   4080
         Width           =   3255
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   36
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   35
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Label lblContact 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7680
         TabIndex        =   34
         Top             =   5760
         Width           =   3615
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         X1              =   1920
         X2              =   1920
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Line Line8 
         BorderWidth     =   5
         X1              =   14280
         X2              =   14280
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Train No. :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey date :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12720
         TabIndex        =   31
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "From Station :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "To Station :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   29
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808000&
         Caption         =   $"frmCancellation.frx":0000
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   2520
         Width           =   12375
      End
      Begin VB.Line Line6 
         BorderWidth     =   5
         X1              =   1920
         X2              =   14280
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No. :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   22
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Line Line7 
         BorderWidth     =   5
         X1              =   1920
         X2              =   14280
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblTrainNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblClass 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   13680
         TabIndex        =   20
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblDateJourney 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   7800
         TabIndex        =   19
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblToStn 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblFromStn 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblBook 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   13680
         TabIndex        =   15
         Top             =   960
         Width           =   2655
      End
      Begin VB.Line Line9 
         BorderWidth     =   5
         X1              =   12120
         X2              =   12120
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Line Line10 
         BorderWidth     =   5
         X1              =   10560
         X2              =   10560
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Label lblSeatNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   11040
         TabIndex        =   14
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblSeatNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   11040
         TabIndex        =   13
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label lblSeatNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   11040
         TabIndex        =   12
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label lblSeatNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   11040
         TabIndex        =   11
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblSeatNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   4
         Left            =   11040
         TabIndex        =   10
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label lblSeatType 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   12240
         TabIndex        =   9
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblSeatType 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   12240
         TabIndex        =   8
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label lblSeatType 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   12240
         TabIndex        =   7
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblSeatType 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   12240
         TabIndex        =   6
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label lblSeatType 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Journey Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   4
         Left            =   12240
         TabIndex        =   5
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   3000
         X2              =   3000
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   6720
         X2              =   6720
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Line Line5 
         BorderWidth     =   5
         X1              =   8040
         X2              =   8040
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   9360
         X2              =   9360
         Y1              =   2520
         Y2              =   5640
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   240
         Picture         =   "frmCancellation.frx":008E
         Top             =   360
         Width           =   16005
      End
      Begin VB.Image Image7 
         Height          =   405
         Left            =   240
         Picture         =   "frmCancellation.frx":152BC
         Top             =   2040
         Width           =   16005
      End
      Begin VB.Image Image9 
         Height          =   4245
         Left            =   240
         Picture         =   "frmCancellation.frx":2A4EA
         Top             =   2040
         Width           =   16005
      End
      Begin VB.Image Image8 
         Height          =   1485
         Left            =   240
         Picture         =   "frmCancellation.frx":107B18
         Top             =   720
         Width           =   16005
      End
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   7560
      Picture         =   "frmCancellation.frx":155266
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   3135
   End
   Begin VB.CommandButton cmdDetail 
      Height          =   375
      Left            =   9840
      Picture         =   "frmCancellation.frx":1595BC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtPnr 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7200
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   720
      Picture         =   "frmCancellation.frx":15C75A
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Cancellation"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   57
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmCancellation.frx":15D0A5
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11520
      Picture         =   "frmCancellation.frx":15D57F
      Top             =   0
      Width           =   11535
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000E&
      Caption         =   "PNR Number. :"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   5985
      Left            =   0
      Picture         =   "frmCancellation.frx":15DA59
      Top             =   1560
      Width           =   825
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   17280
      Picture         =   "frmCancellation.frx":16E073
      Top             =   1560
      Width           =   825
   End
End
Attribute VB_Name = "frmCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPnr As ADODB.Recordset
Dim rstPnr1 As ADODB.Recordset
Dim rstStn As ADODB.Recordset
Dim rstSeat As ADODB.Recordset
Dim rstPass As ADODB.Recordset
Dim cmdCan As ADODB.Command
Private Sub cmdCancel_Click()
If vbYes = MsgBox("Are you sure want to Cancel The Ticket ?", vbQuestion + vbYesNo, "") Then
    Dim seat As Variant
        Set rstPnr = New ADODB.Recordset
        rstPnr.CursorLocation = adUseClient
        rstPnr.Open "select * from reservation,passenger where reservation.pnrno=passenger.pnrno and reservation.pnrno='" & txtPnr.Text & "'", railCn
        If rstPnr.RecordCount > 0 Then
            rstPnr.MoveFirst
            For i = 0 To rstPnr.RecordCount - 1
                Set rstSeat = New ADODB.Recordset
                rstSeat.CursorLocation = adUseClient
                rstSeat.Open "select * from trainseat where trainno='" & rstPnr("trainno") & "' and journeydate=#" & rstPnr("journeydate") & "# and coachname='" & rstPnr("coachname") & "'", railCn
                rstSeat.MoveFirst
                If rstPnr("seatno") < 10 Then
                    seat = "00" & rstPnr("seatno")
                Else
                    seat = "0" & rstPnr("seatno")
                End If
                Set rstPass = New ADODB.Recordset
                rstPass.CursorLocation = adUseClient
                rstPass.Open "select * from reservation,passenger where reservation.pnrno=passenger.pnrno and trainno='" & rstPnr("trainno") & "' and coachname='" & rstPnr("coachname") & "' and journeydate = #" & rstPnr("journeydate") & "# and seatno = " & rstPnr("seatno") & " and reservation.pnrno<>'" & txtPnr.Text & "' and cancel=" & False & "", railCn
                If rstPass.RecordCount = 0 Then
                    pos = InStr(1, rstSeat("availableseat"), seat)
                    First = Mid(rstSeat("availableseat"), 1, pos + 3)
                    length = Len(rstSeat("availableseat")) - Len(First)
                    Last = Mid(rstSeat("availableseat"), pos + 4, length)
                    newseat = Mid(First, 1, Len(First) - 1) & "N" & Last
                    Set cmdCan = New ADODB.Command
                    cmdCan.CommandType = adCmdText
                    cmdCan.ActiveConnection = railCn
                    cmdCan.CommandText = "update trainseat set availableseat='" & newseat & "' where trainno='" & rstPnr("trainno") & "' and coachname='" & rstPnr("coachname") & "' and journeydate = #" & rstPnr("journeydate") & "#"
                    cmdCan.Execute
                End If
                rstPnr.MoveNext
            Next
            Set cmdCan = New ADODB.Command
            cmdCan.CommandType = adCmdText
            cmdCan.ActiveConnection = railCn
            cmdCan.CommandText = "update reservation set cancel=" & True & " where pnrno='" & txtPnr.Text & "'"
            cmdCan.Execute
            MsgBox "Ticket Cancelled Successfully.", vbInformation
            Frame1.Visible = False
            Image4.Visible = False
            cmdCancel.Visible = False
            Image5.Visible = False
        Else
            MsgBox "PNR Doesn't Exists.", vbCritical
        End If
End If
End Sub

Private Sub cmdDetail_Click()
If txtPnr.Text <> "" Then
    Set rstPnr = New ADODB.Recordset
    rstPnr.CursorLocation = adUseClient
    rstPnr.Open "select * from reservation,passenger,coach,traincoach,coachberth,berthtype where cancel=" & False & " and coachberth.berthtypeid=berthtype.berthid and coachberth.coachtypeid=coach.coachtypeid and coachberth.seatno=passenger.seatno and traincoach.coachtypeid=coach.coachtypeid and traincoach.coachname=passenger.coachname and traincoach.trainno=reservation.trainno and reservation.pnrno=passenger.pnrno and reservation.pnrno='" & txtPnr.Text & "'", railCn
    
    Set rstPnr1 = New ADODB.Recordset
    rstPnr1.CursorLocation = adUseClient
    rstPnr1.Open "select * from reservation,passenger,coach,traincoach,coachberth,berthtype where cancel=" & True & " and coachberth.berthtypeid=berthtype.berthid and coachberth.coachtypeid=coach.coachtypeid and coachberth.seatno=passenger.seatno and traincoach.coachtypeid=coach.coachtypeid and traincoach.coachname=passenger.coachname and traincoach.trainno=reservation.trainno and reservation.pnrno=passenger.pnrno and reservation.pnrno='" & txtPnr.Text & "'", railCn
    If rstPnr.RecordCount > 0 Then
        For i = 0 To 4
            lblName(i).Visible = False
            lblAge(i).Visible = False
            lblGender(i).Visible = False
            lblcoach(i).Visible = False
            lblSeatType(i).Visible = False
            lblSeatNo(i).Visible = False
        Next
        Frame1.Visible = True
        Image4.Visible = True
        Image5.Visible = True
        cmdCancel.Visible = True
        rstPnr.MoveFirst
        lblTrainNo.Caption = rstPnr("reservation.trainno")
        lblDateJourney.Caption = rstPnr("journeydate")
        lblClass.Caption = rstPnr("coachtypename")
        lblBook.Caption = rstPnr("bookdate")
        Set rstStn = New ADODB.Recordset
        rstStn.CursorLocation = adUseClient
        rstStn.Open "select * from station where stnid=" & rstPnr("fromstn") & "", railCn
        lblFromStn.Caption = rstStn("stnname") & "(" & rstStn("stncode") & ")"
        Set rstStn = New ADODB.Recordset
        rstStn.CursorLocation = adUseClient
        rstStn.Open "select * from station where stnid=" & rstPnr("tostn") & "", railCn
        lblToStn.Caption = rstStn("stnname") & "(" & rstStn("stncode") & ")"
        lblContact.Caption = rstPnr("contactno")
        Frame1.Caption = "Ticket Details of PNR (" & txtPnr.Text & ")"
        For i = 0 To rstPnr.RecordCount - 1
            lblName(i).Visible = True
            lblAge(i).Visible = True
            lblGender(i).Visible = True
            lblcoach(i).Visible = True
            lblSeatType(i).Visible = True
            lblSeatNo(i).Visible = True
            lblName(i).Caption = rstPnr("passengername")
            lblAge(i).Caption = rstPnr("age")
            lblGender(i).Caption = rstPnr("gender")
            lblcoach(i).Caption = rstPnr("passenger.coachname")
            lblSeatNo(i).Caption = rstPnr("passenger.seatno")
            lblSeatType(i).Caption = rstPnr("berthtype")
            rstPnr.MoveNext
        Next
    ElseIf rstPnr1.RecordCount > 0 Then
        MsgBox "Ticket Already Cancelled " & txtPnr & " PNR Number", vbCritical
        Frame1.Visible = False
        Image4.Visible = False
        cmdCancel.Visible = False
        Image5.Visible = False
    Else
        MsgBox "No Ticket is Issued To " & txtPnr & " PNR Number", vbCritical
        Frame1.Visible = False
        Image4.Visible = False
        cmdCancel.Visible = False
        Image5.Visible = False
    
    End If
Else
    MsgBox "Please Enter PNR Number.", vbCritical
    Frame1.Visible = False
    Image4.Visible = False
    cmdCancel.Visible = False
    Image5.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
    If Label5.ForeColor = vbYellow Then
        Label5.ForeColor = vbHighlight
    ElseIf Label5.ForeColor = &H8080FF Then
        Label5.ForeColor = vbYellow
    ElseIf Label5.ForeColor = vbHighlight Then
        Label5.ForeColor = vbGreen
    Else
        Label5.ForeColor = &H8080FF
    End If

End Sub

Private Sub Form_load()
    Frame1.Visible = False
    Image4.Visible = False
    cmdCancel.Visible = False
    Image5.Visible = False
End Sub

Private Sub txtPnr_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtPnr.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtPnr.Text = txtPnr.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtPnr.Text = txtPnr.Text & Chr(KeyAscii)
    End If
End If
End Sub
